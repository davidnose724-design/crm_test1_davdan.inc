"""
app.py — Interfaz Streamlit del CRM + asistente de prospección B2B.

Ejecutar:  streamlit run app.py

Toda la lógica vive en crm_core.py. Esta capa solo orquesta:
selección de industria, configuración de campaña, vista previa con aprobación
humana, marcado de envíos, bandeja de respuestas, follow-ups por lote, estados
y dashboard con las 8 gráficas. La app NUNCA envía sola: prepara el mensaje, la
persona aprueba y envía por su cuenta, y luego marca "Enviado".
"""

import datetime as dt
from pathlib import Path

import pandas as pd
import streamlit as st

from crm_core import (
    CRM, CHANNEL_STEPS, STAGE_COLS, INDUSTRIES, RESPONSE_CLASSES,
    RESPONSE_RESULTS, OUTCOME_VALUES, ACTIVITY_CHANNELS, IMPORT_FIELDS,
    TEMPLATE_COLUMNS, MANUAL_STATES, PROSPECT_CHANNELS, WARMTH_LEVELS,
    SCHEDULE_STATUS, StateStore, _is_marked, _norm,
)
import pandas as pd
import integrations

st.set_page_config(page_title="Prospecting CRM", layout="wide")

# --------------------------------------------------------------------------- #
# Carga del CRM (en sesión)
# --------------------------------------------------------------------------- #

DEFAULT_XLSX = "Manufacturing_Prospecting_CRM_v4_nuevos_leads.xlsx"

st.sidebar.title("⚙️ Configuración")
xlsx_path = st.sidebar.text_input("Ruta del Excel", DEFAULT_XLSX)

@st.cache_resource(show_spinner="Cargando CRM…")
def get_crm(path, mtime):
    return CRM(path)

if not Path(xlsx_path).exists():
    st.error(f"No encuentro el archivo: {xlsx_path}")
    st.stop()

crm = get_crm(xlsx_path, Path(xlsx_path).stat().st_mtime)

def persist():
    """Guarda con backup previo y avisa claramente si el Excel está abierto."""
    crm.ensure_formatting()
    out, bkp, err = crm.save()
    if err:
        st.error("⚠️ " + err)
        return False
    msg = "Excel y estado guardados ✅"
    if bkp:
        msg += f"  ·  backup: {Path(bkp).name}"
    st.toast(msg)
    return True

# --------------------------------------------------------------------------- #
# Navegación
# --------------------------------------------------------------------------- #

page = st.sidebar.radio(
    "Sección",
    ["📣 Campaña", "📤 Importar leads", "🔁 Follow-ups", "📥 Respuestas",
     "🔔 Notifications", "🏷️ Estados", "📊 Dashboard"],
)
st.sidebar.caption("Aprobación humana obligatoria antes de cada envío. "
                   "La app no automatiza LinkedIn ni evade límites. "
                   "Backup automático antes de cada guardado.")

# =========================================================================== #
# 1) CAMPAÑA
# =========================================================================== #
if page == "📣 Campaña":
    st.header("📣 Nueva campaña")

    # --- Tabla de leads disponibles (con filtros) ANTES de prospectar ---
    st.subheader("Leads disponibles")

    @st.cache_data(show_spinner="Cargando leads disponibles…")
    def _available_table(ind, mtime):
        rows = []
        for l in crm.all_leads(ind):
            stt = crm.state.get(l.key)
            last_contact = crm._cell(crm.wb[l.sheet], crm.maps[l.sheet],
                                     l.row, "Last Contact")
            rows.append({
                "Nombre": l.full_name, "Empresa": l.company, "Puesto": l.job_title,
                "Seniority": l.seniority_level or "—", "Industria": l.industry,
                "LinkedIn": l.linkedin, "Email": l.email,
                "Prioridad": crm.lead_priority(l),
                "Estado": crm.lead_status(l, stt),
                "Último contacto": str(last_contact or ""),
                "Siguiente acción": crm.next_action(l, stt),
            })
        return pd.DataFrame(rows)

    industry = st.selectbox("Industria", INDUSTRIES)
    avail = _available_table(industry, Path(xlsx_path).stat().st_mtime)

    fc = st.columns(4)
    f_sen = fc[0].multiselect("Seniority", sorted(avail["Seniority"].unique()))
    f_comp = fc[1].multiselect("Empresa", sorted(avail["Empresa"].dropna().unique())[:500])
    f_prio = fc[2].multiselect("Prioridad", ["Alta", "Media-Alta", "Media", "Baja"])
    f_est = fc[3].multiselect("Estado", sorted(avail["Estado"].unique()))
    view = avail.copy()
    if f_sen: view = view[view["Seniority"].isin(f_sen)]
    if f_comp: view = view[view["Empresa"].isin(f_comp)]
    if f_prio: view = view[view["Prioridad"].isin(f_prio)]
    if f_est: view = view[view["Estado"].isin(f_est)]
    st.dataframe(view, use_container_width=True, hide_index=True)
    st.caption(f"{len(view)} leads en la vista (de {len(avail)} en {industry}).")

    # ======================================================================= #
    # Crear campaña filtrada (selección + programación de mensajes)
    # ======================================================================= #
    st.divider()
    st.subheader("🎯 Crear campaña filtrada")

    cf = st.columns(4)
    cf_channel = cf[0].selectbox("Canal de prospección", PROSPECT_CHANNELS)
    cf_country = cf[1].text_input("País/ubicación contiene")
    cf_warm = cf[2].multiselect("Nivel de calentamiento", WARMTH_LEVELS)
    cf_n = cf[3].number_input("Nº de leads", 1, 200, 30)

    # Construir el universo filtrado (reutiliza los filtros de arriba + estos)
    filtered = []
    for l in crm.all_leads(industry):
        stt = crm.state.get(l.key)
        if f_sen and (l.seniority_level or "—") not in f_sen: continue
        if f_comp and l.company not in f_comp: continue
        if f_prio and crm.lead_priority(l) not in f_prio: continue
        if f_est and crm.lead_status(l, stt) not in f_est: continue
        if cf_warm and crm.classify_warmth(l, stt) not in cf_warm: continue
        loc = crm._cell(crm.wb[l.sheet], crm.maps[l.sheet], l.row, "Location") or ""
        if cf_country and _norm(cf_country) not in _norm(str(loc)): continue
        filtered.append((l, stt, loc))
    filtered = filtered[:cf_n]

    if not filtered:
        st.info("Ningún lead cumple esos filtros. Ajusta los parámetros.")
    else:
        st.write(f"**{len(filtered)} leads** cumplen los filtros. "
                 "Selecciona/deselecciona y prepara el mensaje de cada uno.")
        msg_tmpl = st.text_area(
            "Plantilla de mensaje sugerido (usa {first}, {company}, {title})",
            "Hola {first}, vi tu rol en {company} y me gustaría conectar.")
        c_i1, c_i2 = st.columns(2)
        start = c_i1.text_input("Fecha/hora inicio (YYYY-MM-DD HH:MM)",
                                dt.datetime.now().strftime("%Y-%m-%d %H:%M"))
        interval = c_i2.number_input("Intervalo entre leads (min)", 1, 1440, 45)
        try:
            t0 = dt.datetime.strptime(start, "%Y-%m-%d %H:%M")
        except ValueError:
            t0 = dt.datetime.now()

        selected = []
        for i, (l, stt, loc) in enumerate(filtered):
            first = (l.full_name or "").split(" ")[0]
            last_msg = ""
            for e in reversed(stt.log):
                if e.get("action") == "sent":
                    last_msg = (e.get("message") or "")[:50]; break
            last_contact = crm._cell(crm.wb[l.sheet], crm.maps[l.sheet],
                                     l.row, "Last Contact")
            send_at = t0 + dt.timedelta(minutes=i * interval)
            with st.expander(
                f"{i+1:>2}. {l.full_name or '(sin nombre)'} · {l.company} · "
                f"{crm.lead_priority(l)} · {crm.classify_warmth(l, stt)} · ⏱ {send_at:%m-%d %H:%M}"):
                cc = st.columns(2)
                cc[0].write(f"**Puesto:** {l.job_title or '—'}  \n"
                            f"**Seniority:** {l.seniority_level or '—'}  \n"
                            f"**Industria:** {l.industry}  \n"
                            f"**Ubicación:** {loc or '—'}")
                cc[1].write(f"**LinkedIn:** {l.linkedin or '—'}  \n"
                            f"**Email:** {l.email or '—'}  \n"
                            f"**Teléfono:** {l.phone or '—'}  \n"
                            f"**Canal recomendado:** {crm.recommend_channel(l)}")
                st.caption(f"Follow-up actual: {crm.lead_stage(l, stt)} · "
                           f"Último contacto: {last_contact or '—'} · "
                           f"Último mensaje: {last_msg or '—'}")
                default_msg = msg_tmpl.format(first=first, company=l.company,
                                              title=l.job_title)
                msg = st.text_area("Mensaje sugerido (editable)", default_msg,
                                   key=f"cf_msg{i}", height=80)
                pick = st.checkbox("Incluir en la campaña", value=True, key=f"cf_pick{i}")
                if pick:
                    selected.append((l, msg, send_at))

        st.write(f"Seleccionados: **{len(selected)}**")
        cbtn = st.columns(2)
        if cbtn[0].button("📝 Guardar como BORRADORES (Draft)", key="cf_draft"):
            for l, msg, send_at in selected:
                crm.schedule_message(l, cf_channel, msg, send_at.isoformat(timespec="minutes"),
                                     status="Draft")
            if persist():
                st.success(f"{len(selected)} mensajes guardados como Draft en "
                           "Scheduled_Messages.")
        if cbtn[1].button("📅 Programar (Scheduled)", type="primary", key="cf_sched"):
            for l, msg, send_at in selected:
                crm.schedule_message(l, cf_channel, msg, send_at.isoformat(timespec="minutes"),
                                     status="Scheduled")
            if persist():
                st.success(f"{len(selected)} mensajes programados. Revísalos y márcalos "
                           "como enviados manualmente en 'Agenda de mensajes'.")

    # --- Agenda de mensajes (revisar y marcar enviado manualmente) ---
    st.divider()
    st.subheader("🗓️ Agenda de mensajes (Scheduled_Messages)")
    st.caption("La app NO envía nada sola. Aquí revisas la cola y marcas cada mensaje "
               "como enviado tú mismo cuando lo mandaste.")
    sched = crm.read_scheduled()
    if not sched:
        st.info("No hay mensajes en la agenda todavía.")
    else:
        fstat = st.multiselect("Filtrar por estado", SCHEDULE_STATUS,
                               default=["Draft", "Scheduled"])
        shown = [s for s in sched if not fstat or s.get("Status") in fstat]
        st.dataframe(pd.DataFrame([{k: v for k, v in s.items() if k != "_row"}
                                   for s in shown]),
                     use_container_width=True, hide_index=True)
        ids = {f"{s['Schedule ID']} · {s['Full Name']} · {s['Status']}": s for s in shown}
        if ids:
            pick = st.selectbox("Actualizar mensaje", list(ids.keys()))
            row = ids[pick]["_row"]
            a1, a2, a3 = st.columns(3)
            if a1.button("Marcar ENVIADO (registro real)", type="primary"):
                crm.update_schedule_status(row, "Sent")
                if persist():
                    st.success("Marcado como enviado y registrada la fecha real.")
            if a2.button("Marcar Scheduled"):
                crm.update_schedule_status(row, "Scheduled")
                if persist(): st.success("Marcado como Scheduled.")
            if a3.button("Cancelar mensaje"):
                crm.update_schedule_status(row, "Cancelled")
                if persist(): st.success("Cancelado.")

    st.divider()
    st.subheader("Configurar campaña")
    c2, c3 = st.columns(2)
    channel = c2.selectbox("Canal", list(CHANNEL_STEPS.keys()))
    step_label = c3.selectbox("Paso", CHANNEL_STEPS[channel])
    step_index = CHANNEL_STEPS[channel].index(step_label)

    # --- Apartado Seniority Level (lee el del Excel y filtra la campaña) ---
    st.subheader("Seniority Level")
    sen_values = sorted({(l.seniority_level or "—") for l in crm.all_leads(industry)})
    sel_sen = st.multiselect(
        "Filtrar por Seniority Level (vacío = todos). Se usa para calcular prioridad.",
        sen_values, default=f_sen)
    st.caption("Prioridad: Owner/Founder/CEO/President = **Alta** · "
               "VP/Director/Head/Manager = **Media-Alta** · "
               "Buyer/Purchasing/Sourcing/Procurement/Engineer = **Media** · "
               "Assistant/Coordinator/Intern = **Baja**.")

    c4, c5, c6 = st.columns(3)
    n_leads = c4.number_input("Leads por campaña", 1, 200, 30)
    interval = c5.number_input("Intervalo entre leads (min)", 1, 1440, 45)
    start = c6.text_input("Inicio (YYYY-MM-DD HH:MM)",
                          dt.datetime.now().strftime("%Y-%m-%d %H:%M"))
    gap_days = st.number_input("Programar siguiente follow-up en (días)", 1, 60, 3)

    tmpl = st.text_area(
        "Plantilla del mensaje (usa {first}, {company}, {title})",
        "Hola {first}, vi tu rol en {company} y me gustaría conectar.")

    if st.button("Construir cola", type="primary"):
        q = crm.build_queue(industry, channel, step_index, n_leads * 3)
        if sel_sen:
            q = [l for l in q if (l.seniority_level or "—") in sel_sen]
        st.session_state.queue = q[:n_leads]
        st.session_state.qmeta = dict(industry=industry, channel=channel,
                                      step=step_label, step_index=step_index,
                                      interval=interval, start=start, gap_days=gap_days,
                                      tmpl=tmpl)

    queue = st.session_state.get("queue", [])
    if queue:
        meta = st.session_state.qmeta
        st.success(f"{len(queue)} leads en cola · {meta['channel']} · {meta['step']} · "
                   f"1 cada {meta['interval']} min")
        # Distribución de prioridad en la cola
        prio_counts = {}
        for l in queue:
            p = crm.lead_priority(l)
            prio_counts[p] = prio_counts.get(p, 0) + 1
        st.caption("Prioridad en la cola: " +
                   " · ".join(f"{k}: {v}" for k, v in prio_counts.items()))
        try:
            t0 = dt.datetime.strptime(meta["start"], "%Y-%m-%d %H:%M")
        except ValueError:
            t0 = dt.datetime.now()

        st.subheader("Vista previa — aprueba lead por lead")
        approvals = {}
        for i, lead in enumerate(queue):
            send_at = t0 + dt.timedelta(minutes=i * meta["interval"])
            first = (lead.full_name or "").split(" ")[0]
            msg = meta["tmpl"].format(first=first, company=lead.company,
                                      title=lead.job_title)
            prio = crm.lead_priority(lead)
            with st.expander(
                f"{i+1:>2}. {lead.full_name} · {lead.company} · "
                f"{lead.seniority_level or '—'} · 🎯 {prio} · ⏱ {send_at:%H:%M}"):
                msg = st.text_area("Mensaje", msg, key=f"msg{i}", height=90)
                approvals[i] = st.checkbox("✅ Aprobar este mensaje", key=f"ap{i}")
                st.session_state[f"finalmsg{i}"] = msg

        st.divider()
        if st.button("Marcar aprobados como ENVIADOS", type="primary"):
            sent = 0
            for i, lead in enumerate(queue):
                if approvals.get(i):
                    msg = st.session_state.get(f"finalmsg{i}", "")
                    crm.mark_sent(lead, meta["channel"], meta["step"], message=msg)
                    crm.schedule_next(lead, meta["channel"], meta["step_index"],
                                      meta["gap_days"])
                    sent += 1
            if persist():
                st.success(f"{sent} marcados como enviados y siguiente follow-up programado.")
                st.session_state.queue = crm.build_queue(
                    meta["industry"], meta["channel"], meta["step_index"], len(queue))

# =========================================================================== #
# 2) IMPORTAR LEADS
# =========================================================================== #
elif page == "📤 Importar leads":
    st.header("📤 Importar leads desde Excel/CSV")

    with st.expander("ℹ️ Formato esperado del archivo", expanded=True):
        st.markdown(
            "El archivo puede ser **.xlsx** o **.csv**. No necesita todas las columnas: "
            "se importa cada lead **con lo que tenga**. Columnas recomendadas (los nombres "
            "pueden variar; la app las detecta y puedes corregir el mapeo):")
        st.dataframe(pd.DataFrame([
            {"First Name": "Jane", "Last Name": "Doe", "Full Name": "Jane Doe",
             "Job Title": "VP Procurement", "Company Name": "Acme Manufacturing",
             "Industry": "Automotive", "Seniority Level": "VP",
             "LinkedIn URL": "linkedin.com/in/janedoe", "Email": "jane@acme.com",
             "Phone": "+1 555 0100", "Location": "Detroit, MI",
             "Value Proposition": "Reducimos 15% tooling", "Pain Point": "Lead times",
             "Notes": "Feria 2025"},
            {"First Name": "", "Last Name": "", "Full Name": "John Smith",
             "Job Title": "Buyer", "Company Name": "Globex", "Industry": "Packaging",
             "Seniority Level": "", "LinkedIn URL": "linkedin.com/in/johnsmith",
             "Email": "john@globex.com", "Phone": "", "Location": "",
             "Value Proposition": "", "Pain Point": "Empaque sustentable", "Notes": ""},
        ]), use_container_width=True, hide_index=True)
        tmpl_path = Path(xlsx_path).resolve().parent / "plantilla_importacion_leads.xlsx"
        crm.make_import_template(str(tmpl_path.parent))
        with open(tmpl_path, "rb") as fh:
            st.download_button("⬇️ Descargar plantilla_importacion_leads.xlsx", fh.read(),
                               file_name="plantilla_importacion_leads.xlsx",
                               mime="application/vnd.openxmlformats-officedocument."
                                    "spreadsheetml.sheet")

    ups = st.file_uploader("Sube uno o varios archivos .xlsx/.csv",
                           type=["xlsx", "xls", "csv"], accept_multiple_files=True)
    if ups:
        frames = []
        for up in ups:
            try:
                d = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
                d["__archivo__"] = up.name
                frames.append(d)
            except Exception as e:
                st.error(f"No pude leer {up.name}: {e}")
        if frames:
            df = pd.concat(frames, ignore_index=True)
            st.subheader("Vista previa consolidada")
            st.caption(f"{len(ups)} archivo(s) · {len(df)} filas combinadas")
            st.dataframe(df.head(30), use_container_width=True)

            src_cols = [c for c in df.columns if c != "__archivo__"]
            auto = crm.detect_import_mapping(src_cols)
            st.subheader("Mapeo de columnas")
            st.caption("Detecté estas correspondencias. Ajusta manualmente si hace falta.")
            none_opt = "—(ninguna)—"
            options = [none_opt] + src_cols
            mapping = {}
            cols = st.columns(2)
            for i, field in enumerate(IMPORT_FIELDS):
                default = auto.get(field)
                idx = options.index(default) if default in options else 0
                with cols[i % 2]:
                    sel = st.selectbox(field, options, index=idx, key=f"map_{field}")
                if sel != none_opt:
                    mapping[field] = sel

            # Pre-cálculo de duplicados en el consolidado (informativo)
            def row_key(r):
                g = lambda f: (str(r[mapping[f]]).strip()
                               if mapping.get(f) and pd.notna(r.get(mapping[f])) else "")
                full = g("Full Name") or (g("First Name") + " " + g("Last Name")).strip()
                return StateStore.key(g("LinkedIn Profile") or None, g("Email/Gmail") or None,
                                      full or None, g("Company Name") or None,
                                      g("Phone") or None)
            if mapping:
                keys = [row_key(r) for _, r in df.iterrows()]
                dup_in_file = len(keys) - len(set(keys))
                st.caption(f"Duplicados detectados dentro del archivo: {dup_in_file} "
                           "(se omitirán; también se omiten los que ya existan en el CRM).")

            c1, c2 = st.columns(2)
            route = c1.checkbox("Enrutar por columna de industria", value=True)
            default_ind = c2.selectbox("Industria por defecto (si falta o no enruta)",
                                       INDUSTRIES, index=INDUSTRIES.index("Other"))
            st.caption("Se agregan los leads con lo disponible; solo se omiten filas "
                       "totalmente vacías y duplicados (LinkedIn / email / nombre+empresa).")

            if st.button("Importar al Excel principal", type="primary"):
                if not mapping:
                    st.error("Mapea al menos una columna para poder importar.")
                else:
                    if not (mapping.get("LinkedIn Profile") or mapping.get("Email/Gmail")
                            or mapping.get("Full Name") or mapping.get("First Name")
                            or mapping.get("Company Name") or mapping.get("Phone")):
                        st.warning("Sin columna de identidad: los leads se agregarán pero "
                                   "no podrán deduplicarse ni seguirse individualmente.")
                    added, dups, per = crm.import_leads(
                        df.drop(columns=["__archivo__"], errors="ignore"),
                        mapping, route_by_industry=route, default_industry=default_ind)
                    if persist():
                        st.success(f"✅ {added} leads importados · {dups} duplicados omitidos.")
                        if per:
                            st.table(pd.DataFrame(
                                [{"Industria": k, "Agregados": v} for k, v in per.items()]))

# =========================================================================== #
# 3) FOLLOW-UPS
# =========================================================================== #
elif page == "🔁 Follow-ups":
    st.header("🔁 Follow-ups")
    industry = st.selectbox("Industria", INDUSTRIES)
    channel = st.selectbox("Canal", ACTIVITY_CHANNELS)
    seq_channel = "Cold Call" if channel in ("Cold Call", "Llamada") else \
                  ("Email" if channel == "Email" else "LinkedIn")
    today = dt.date.today()

    # --- Vista general con todas las columnas + filtros por fase/estado ---
    st.subheader("Vista general")
    rows = []
    for lead in crm.all_leads(industry):
        stt = crm.state.get(lead.key)
        steps = CHANNEL_STEPS[seq_channel]
        current = "—"
        for s in steps:
            if _is_marked(lead.stage_values.get(s)):
                current = s
        nxt_step, nxt_date = "—", "—"
        if stt.scheduled:
            k = sorted(stt.scheduled.items(), key=lambda kv: kv[1])[0]
            nxt_step = k[0]
            nxt_date = k[1][:10]
        last_msg = "—"
        for e in reversed(stt.log):
            if e.get("action") == "sent":
                last_msg = (e.get("message") or "—")[:60]
                break
        last_contact = crm._cell(crm.wb[lead.sheet], crm.maps[lead.sheet],
                                 lead.row, "Last Contact")
        rows.append({
            "Nombre": lead.full_name, "Empresa": lead.company, "Puesto": lead.job_title,
            "Seniority": lead.seniority_level or "—", "Industria": lead.industry,
            "Follow up actual": current,
            "Último contacto": str(last_contact or "—"),
            "Canal último contacto": crm.lead_channel(lead, stt),
            "Mensaje usado": last_msg,
            "Siguiente follow-up": f"{nxt_step} ({nxt_date})" if nxt_step != "—" else
                                   crm.next_action(lead, stt),
            "Estado": crm.lead_status(lead, stt),
        })
    fdf = pd.DataFrame(rows)

    # Filtros por fase / estado
    phase_filters = ["Follow Up 1 pendiente", "Follow Up 1 enviado",
                     "Follow Up 2 pendiente", "Follow Up 2 enviado",
                     "Follow Up 3 pendiente", "Follow Up 3 enviado",
                     "respondieron", "prospectar después"]
    fp = st.multiselect("Ver", phase_filters)

    def _match(row):
        if not fp:
            return True
        st_ = row["Estado"]; cur = row["Follow up actual"]
        for f in fp:
            if f == "respondieron" and st_ == "respondió":
                return True
            if f == "prospectar después" and st_ == "prospectar después":
                return True
            if f.startswith("Follow Up"):
                n = f.split(" ")[2]
                enviado = cur == f"Follow Up {n}"
                if "enviado" in f and enviado:
                    return True
                if "pendiente" in f and not enviado and st_ == "pendiente":
                    return True
        return False

    if not fdf.empty:
        view = fdf[fdf.apply(_match, axis=1)] if fp else fdf
        st.dataframe(view, use_container_width=True, hide_index=True)
        st.caption(f"{len(view)} de {len(fdf)} leads.")
    else:
        st.info("Aún no hay leads en esta industria.")

    # --- Editor por lead: marcar enviado / cambiar canal / estado / nota ---
    st.divider()
    st.subheader("Actualizar un lead desde aquí")
    leads = crm.all_leads(industry)
    if leads:
        labels = [f"{l.full_name} · {l.company} · [{crm.lead_stage(l)}]" for l in leads]
        si = st.selectbox("Lead", range(len(leads)), format_func=lambda i: labels[i],
                          key="fu_lead")
        lead = leads[si]
        e1, e2 = st.columns(2)
        upd_channel = e1.selectbox("Canal", ACTIVITY_CHANNELS, key="fu_ch")
        # próximo paso a marcar en el canal de secuencia elegido
        seq2 = "Cold Call" if upd_channel in ("Cold Call", "Llamada") else \
               ("Email" if upd_channel == "Email" else "LinkedIn")
        nxt = None
        for s in CHANNEL_STEPS[seq2]:
            if not _is_marked(lead.stage_values.get(s)):
                nxt = s
                break
        e2.write(f"Próximo paso a marcar: **{nxt or '— (secuencia completa)'}**")
        note = st.text_input("Nota / mensaje usado", key="fu_note")
        b1, b2 = st.columns(2)
        if b1.button("Marcar follow-up como ENVIADO", type="primary") and nxt:
            crm.mark_sent(lead, upd_channel, nxt, message=note or "")
            idx = CHANNEL_STEPS[seq2].index(nxt)
            crm.schedule_next(lead, seq2, idx, gap_days=3)
            crm.add_notification(crm.read_lead(lead.sheet, lead.row),
                                 "Follow-up enviado", f"{nxt} por {upd_channel}", upd_channel)
            if persist():
                st.success(f"{nxt} marcado como enviado por {upd_channel}.")
        new_state = b2.selectbox("Cambiar estado a", ["—"] + MANUAL_STATES, key="fu_state")
        if st.button("Aplicar cambio de estado") and new_state != "—":
            crm.set_manual_state(lead, new_state, channel=upd_channel, note=note or None)
            if persist():
                st.success(f"Estado → {new_state}.")

    # --- Aprobación por lote de follow-ups programados ---
    st.divider()
    st.subheader("Aprobar y marcar enviados (por lote)")
    due = []
    for lead in crm.all_leads(industry):
        stt = crm.state.get(lead.key)
        for step, when_iso in stt.scheduled.items():
            if step not in CHANNEL_STEPS[seq_channel]:
                continue
            if crm._is_blocked(lead, stt):
                continue
            when = dt.datetime.fromisoformat(when_iso)
            due.append((lead, step, when, when.date() <= today))

    if not due:
        st.info("No hay follow-ups programados para este canal.")
    else:
        st.write(f"{len(due)} programados ({sum(d[3] for d in due)} vencidos):")
        tmpl = st.text_area("Plantilla follow-up", "Hola {first}, retomo mi mensaje anterior.")
        chosen = {}
        for i, (lead, step, when, overdue) in enumerate(due):
            flag = "🔴 vencido" if overdue else f"🗓 {when:%Y-%m-%d}"
            chosen[i] = st.checkbox(
                f"{flag} · {step} · {lead.full_name} · {lead.company}", key=f"fu{i}")
        if st.button("Aprobar lote y marcar ENVIADOS", type="primary"):
            sent = 0
            for i, (lead, step, when, _o) in enumerate(due):
                if chosen.get(i):
                    first = (lead.full_name or "").split(" ")[0]
                    msg = tmpl.format(first=first, company=lead.company)
                    crm.mark_sent(lead, channel, step, message=msg)
                    idx = CHANNEL_STEPS[seq_channel].index(step)
                    crm.schedule_next(lead, seq_channel, idx, gap_days=3)
                    stt = crm.state.get(lead.key); stt.scheduled.pop(step, None)
                    crm.state.put(lead.key, stt)
                    sent += 1
            if persist():
                st.success(f"{sent} follow-ups marcados como enviados.")

    st.divider()
    st.subheader("Sugerencias de escalamiento de canal")
    any_s = False
    for lead in crm.all_leads(industry):
        s = crm.escalation_suggestion(lead)
        if s:
            any_s = True
            st.write(f"• **{lead.full_name}** ({lead.company}): {s}")
    if not any_s:
        st.caption("Sin sugerencias de escalamiento por ahora.")

# =========================================================================== #
# 3) RESPUESTAS
# =========================================================================== #
elif page == "📥 Respuestas":
    st.header("📥 Registrar respuesta de un lead")

    # --- Buscador / autocompletado por nombre / email / LinkedIn / empresa ---
    term = st.text_input("Buscar lead (nombre, email, LinkedIn o empresa)")
    if term:
        t = _norm(term)
        matches = []
        for sheet in crm.maps:
            for l in crm.all_leads(sheet):
                blob = _norm(" ".join([l.full_name, l.email, l.linkedin, l.company]))
                if t in blob:
                    matches.append(l)
                    if len(matches) >= 50:
                        break
            if len(matches) >= 50:
                break
        if not matches:
            st.warning("Sin coincidencias.")
        else:
            labels = [f"{l.full_name or '(sin nombre)'} · {l.company or '—'} · "
                      f"{l.email or l.linkedin or '—'}" for l in matches]
            sel = st.selectbox("Resultados", range(len(matches)),
                               format_func=lambda i: labels[i])
            lead = matches[sel]
            st.session_state.resp_lead = (lead.sheet, lead.row)

    # --- Autocompletado de datos del lead seleccionado ---
    if "resp_lead" in st.session_state:
        sheet, row = st.session_state.resp_lead
        lead = crm.read_lead(sheet, row)
        st_state = crm.state.get(lead.key)
        fu_actual = crm.lead_stage(lead, st_state)
        st.subheader("Lead seleccionado")
        st.dataframe(pd.DataFrame([{
            "Nombre": lead.full_name, "Empresa": lead.company, "Puesto": lead.job_title,
            "Industria": lead.industry, "Seniority": lead.seniority_level,
            "Email": lead.email, "LinkedIn": lead.linkedin, "Follow-up actual": fu_actual,
        }]).T.rename(columns={0: "Valor"}), use_container_width=True)

        st.subheader("Captura de la respuesta")
        message = st.text_area("Mensaje recibido")
        c1, c2 = st.columns(2)
        channel = c1.selectbox("Canal de respuesta", ACTIVITY_CHANNELS)
        result = c2.selectbox("Resultado", RESPONSE_RESULTS)
        recontact = None
        if result == "Prospectar después":
            recontact = st.date_input(
                "Fecha de recontacto (la decides tú)",
                dt.date.today() + dt.timedelta(days=30)).isoformat()
        note = st.text_input("Nota adicional")

        if st.button("Guardar respuesta", type="primary"):
            crm.apply_response_result(lead, result, channel=channel, message=message,
                                      note=note or None, recontact_date=recontact)
            if persist():
                st.success(f"✅ Respuesta guardada · resultado: {result} · "
                           "actualizado en la tabla principal, dashboard y Activity_Log.")
                del st.session_state.resp_lead

# =========================================================================== #
# NOTIFICATIONS
# =========================================================================== #
elif page == "🔔 Notifications":
    st.header("🔔 Notifications")

    tabA, tabB, tabC = st.tabs(["📝 Registrar respuesta (manual)",
                                "🔔 Alertas", "🔌 Integraciones (Modo B)"])

    # ---- Modo A: registro manual de respuesta ---- #
    with tabA:
        st.caption("Registro manual y seguro. Busca el lead, pega el mensaje recibido "
                   "y clasifícalo. No se conecta a ningún canal automáticamente.")
        term = st.text_input("Buscar lead (nombre, email, LinkedIn o empresa)",
                             key="noti_search")
        lead = None
        if term:
            t = _norm(term)
            matches = []
            for sheet in crm.maps:
                for l in crm.all_leads(sheet):
                    blob = _norm(" ".join([l.full_name, l.email, l.linkedin, l.company]))
                    if t in blob:
                        matches.append(l)
                        if len(matches) >= 50: break
                if len(matches) >= 50: break
            if matches:
                labels = [f"{l.full_name or '(sin nombre)'} · {l.company or '—'} · "
                          f"{l.email or l.linkedin or '—'}" for l in matches]
                si = st.selectbox("Resultados", range(len(matches)),
                                  format_func=lambda i: labels[i], key="noti_pick")
                lead = matches[si]
            else:
                st.warning("Sin coincidencias.")

        if lead is not None:
            stt = crm.state.get(lead.key)
            st.info(f"**{lead.full_name}** · {lead.company} · {lead.job_title or '—'} · "
                    f"{lead.industry} · {lead.seniority_level or '—'}  \n"
                    f"Follow-up actual: **{crm.lead_stage(lead, stt)}**")
            msg = st.text_area("Mensaje recibido (pega aquí)")
            r1, r2 = st.columns(2)
            channel = r1.selectbox("Canal", PROSPECT_CHANNELS + ["Otro"])
            rtype = r2.selectbox("Tipo de respuesta", RESPONSE_CLASSES + [
                "Pidió cotización", "Reunión agendada", "Won", "Lost",
                "Prospectar después", "Blacklist"])
            result = st.selectbox("Resultado", RESPONSE_RESULTS)
            recontact = None
            if result == "Prospectar después":
                recontact = st.date_input("Fecha próximo contacto",
                    dt.date.today() + dt.timedelta(days=30)).isoformat()
            note = st.text_input("Nota adicional")
            if st.button("Guardar respuesta", type="primary"):
                crm.apply_response_result(lead, result, channel=channel, message=msg,
                                          note=note or None, recontact_date=recontact)
                crm.add_notification(crm.read_lead(lead.sheet, lead.row),
                                     f"Respuesta: {rtype}", (msg or result)[:120], channel)
                if persist():
                    st.success(f"✅ Guardado. Actualizado estado, Current Stage, "
                               "Activity_Log, Notifications y dashboard.")

    # ---- Alertas (lista de notificaciones) ---- #
    with tabB:
        if st.button("🔄 Detectar follow-ups vencidos y generar alertas"):
            n = crm.scan_due_followups()
            if persist():
                st.success(f"{n} alerta(s) de follow-up creada(s).")
        notis = crm.read_notifications()
        if not notis:
            st.info("No hay notificaciones todavía.")
        else:
            pend = [n for n in notis if str(n.get("Visto")) != "Vista"]
            st.subheader(f"Pendientes: {len(pend)} · Total: {len(notis)}")
            only_pend = st.checkbox("Ver solo pendientes", value=True)
            shown = pend if only_pend else notis
            for n in reversed(shown):
                icon = "🟡" if n.get("Visto") != "Vista" else "✅"
                with st.expander(f"{icon} {n.get('Fecha')} · {n.get('Lead')} · "
                                 f"{n.get('Tipo de evento')}"):
                    st.write(f"**Empresa:** {n.get('Empresa') or '—'}  ·  "
                             f"**Canal:** {n.get('Canal') or '—'}  ·  "
                             f"**Estado:** {n.get('Estado actual') or '—'}")
                    st.write(f"**Resumen:** {n.get('Mensaje/Resumen') or '—'}")
                    if n.get("Visto") != "Vista":
                        if st.button("Marcar como vista", key=f"seen{n['_row']}"):
                            crm.set_notification_seen(n["_row"], True)
                            if persist(): st.rerun()
                    else:
                        if st.button("Marcar como pendiente", key=f"pend{n['_row']}"):
                            crm.set_notification_seen(n["_row"], False)
                            if persist(): st.rerun()
            st.divider()
            st.dataframe(pd.DataFrame([{k: v for k, v in n.items() if k != "_row"}
                                       for n in notis]),
                         use_container_width=True, hide_index=True)

    # ---- Modo B: estado de integraciones (sin credenciales) ---- #
    with tabC:
        st.caption("Arquitectura preparada para integraciones futuras con APIs oficiales. "
                   "No hay credenciales en el código, no se envía ni se lee nada "
                   "automáticamente, y no se hace scraping.")
        for s in integrations.channel_statuses():
            badge = "🟢 API lista" if s.configured else "⚪ Manual"
            st.markdown(f"**{s.channel}** — {badge}")
            st.caption(s.note)
        st.info("LinkedIn: solo registro manual o importación CSV (política de LinkedIn). "
                "Gmail: futura vía Gmail API/OAuth. WhatsApp: futura vía WhatsApp Business "
                "API. Las credenciales, cuando existan, vendrán de variables de entorno.")

# =========================================================================== #
# 4) ESTADOS
# =========================================================================== #
elif page == "🏷️ Estados":
    st.header("🏷️ Estados")

    # --- Tabla resumen agrupada con filtros ---
    st.subheader("Resumen agrupado")

    @st.cache_data(show_spinner="Calculando estados…")
    def _states_table(mtime):
        rows = []
        for sheet in crm.maps:
            for l in crm.all_leads(sheet):
                stt = crm.state.get(l.key)
                status = crm.lead_status(l, stt)
                nxt = ""
                if stt.scheduled:
                    nxt = min(stt.scheduled.values())[:10]
                rows.append({
                    "Industria": l.industry, "Seniority": l.seniority_level or "—",
                    "Paso actual": crm.lead_stage(l, stt),
                    "Canal actual": crm.lead_channel(l, stt),
                    "Estado": status, "Próximo follow-up": nxt,
                    "_contactado": status not in ("pendiente",),
                    "_pendiente": status == "pendiente",
                    "_respondio": status == "respondió",
                    "_won": status == "won", "_lost": status == "lost",
                    "_blacklist": status == "blacklist",
                })
        return pd.DataFrame(rows)

    df = _states_table(Path(xlsx_path).stat().st_mtime)

    f1, f2, f3 = st.columns(3)
    fi = f1.multiselect("Industria", sorted(df["Industria"].unique()))
    fs = f2.multiselect("Seniority", sorted(df["Seniority"].unique()))
    fc = f3.multiselect("Canal actual", sorted(df["Canal actual"].unique()))
    f4c, f5c, f6c = st.columns(3)
    fp = f4c.multiselect("Etapa/fase (paso actual)", sorted(df["Paso actual"].unique()))
    fe = f5c.multiselect("Estado", sorted(df["Estado"].unique()))
    finals = ["won", "lost", "blacklist", "prospectar después"]
    ff = f6c.multiselect("Resultado final", finals)
    fdf = df.copy()
    if fi: fdf = fdf[fdf["Industria"].isin(fi)]
    if fs: fdf = fdf[fdf["Seniority"].isin(fs)]
    if fc: fdf = fdf[fdf["Canal actual"].isin(fc)]
    if fp: fdf = fdf[fdf["Paso actual"].isin(fp)]
    if fe: fdf = fdf[fdf["Estado"].isin(fe)]
    if ff: fdf = fdf[fdf["Estado"].isin(ff)]

    grouped = (fdf.groupby(["Industria", "Seniority", "Paso actual", "Canal actual"])
               .agg(total_leads=("Estado", "size"),
                    contactados=("_contactado", "sum"),
                    pendientes=("_pendiente", "sum"),
                    respondieron=("_respondio", "sum"),
                    won=("_won", "sum"), lost=("_lost", "sum"),
                    blacklist=("_blacklist", "sum"),
                    proximo_follow_up=("Próximo follow-up",
                                       lambda s: min([x for x in s if x], default="")))
               .reset_index())
    st.dataframe(grouped, use_container_width=True, hide_index=True)
    st.caption(f"{len(fdf)} leads en la vista · {len(grouped)} grupos")

    st.divider()
    # --- Editor: cambiar manualmente a cualquiera de los 12 estados ---
    st.subheader("Cambiar estado de un lead")
    industry = st.selectbox("Industria", INDUSTRIES, key="est_ind")
    leads = crm.all_leads(industry)
    labels = [f"{l.full_name} · {l.company} · [{crm.lead_stage(l)}]" for l in leads]
    sel = st.selectbox("Lead", range(len(leads)), format_func=lambda i: labels[i])
    lead = leads[sel]

    st.write(f"**{lead.full_name}** — {lead.job_title} @ {lead.company} · "
             f"estado actual: **{crm.lead_stage(lead)}**")
    new_state = st.selectbox("Current Stage / Estado", MANUAL_STATES)
    ch = st.selectbox("Canal", ACTIVITY_CHANNELS)
    recontact = None
    if new_state == "Prospectar después":
        recontact = st.date_input("Fecha de recontacto (la decides tú)",
                                  dt.date.today() + dt.timedelta(days=30)).isoformat()
    note = st.text_input("Nota")

    if st.button("Aplicar estado", type="primary"):
        crm.set_manual_state(lead, new_state, channel=ch, note=note or None,
                             recontact_date=recontact)
        if persist():
            st.success(f"✅ Estado → {new_state}. Actualizado en tabla, Activity_Log, "
                       "Notifications y dashboard.")
            _states_table.clear()

    # --- Campos editables adicionales ---
    with st.expander("✏️ Editar más campos (warmth, prioridad, canal, próxima acción…)"):
        e1, e2, e3 = st.columns(3)
        warmth = e1.selectbox("Lead Warmth Level", ["—"] + WARMTH_LEVELS)
        priority = e2.selectbox("Priority", ["—", "Alta", "Media", "Baja"])
        rec_ch = e3.selectbox("Recommended Channel", ["—"] + PROSPECT_CHANNELS)
        e4, e5 = st.columns(2)
        fu_step = e4.selectbox("Follow Up Step", ["—"] + CHANNEL_STEPS["LinkedIn"])
        next_date = e5.date_input("Fecha de próximo contacto",
                                  dt.date.today() + dt.timedelta(days=3))
        next_act = st.text_input("Próxima acción")
        more_notes = st.text_input("Notas", key="est_notes")
        if st.button("Guardar campos"):
            crm.edit_lead_fields(
                lead,
                warmth=None if warmth == "—" else warmth,
                priority=None if priority == "—" else priority,
                channel=None if rec_ch == "—" else rec_ch,
                fu_step=None if fu_step == "—" else fu_step,
                next_action=next_act or None,
                next_date=next_date.isoformat(),
                notes=more_notes or None)
            if persist():
                st.success("Campos guardados y registrados en Activity_Log.")
                _states_table.clear()

    st.divider()
    st.subheader("Historial de actividad del lead")
    log = crm.state.get(lead.key).log
    if log:
        st.dataframe(pd.DataFrame(log), use_container_width=True, hide_index=True)
    else:
        st.caption("Sin actividad registrada todavía.")

# =========================================================================== #
# 5) DASHBOARD
# =========================================================================== #
elif page == "📊 Dashboard":
    st.header("📊 Dashboard")

    # --- Métricas globales y desgloses (lead-level) ---
    @st.cache_data(show_spinner="Calculando métricas…")
    def _dash_metrics(mtime):
        total = nuevos = msgs = fus = resp = won = lost = prosp = black = 0
        by_ind, by_sen, by_chan, by_stage = {}, {}, {}, {}
        for sheet in crm.maps:
            for l in crm.all_leads(sheet):
                stt = crm.state.get(l.key)
                total += 1
                stage = crm.lead_stage(l, stt)
                chan = crm.lead_channel(l, stt)
                by_ind[l.industry] = by_ind.get(l.industry, 0) + 1
                sen = l.seniority_level or "—"
                by_sen[sen] = by_sen.get(sen, 0) + 1
                by_chan[chan] = by_chan.get(chan, 0) + 1
                by_stage[stage] = by_stage.get(stage, 0) + 1
                sv = l.stage_values
                fu_sent = sum(_is_marked(sv.get(s)) for s in CHANNEL_STEPS["LinkedIn"])
                em_sent = sum(_is_marked(sv.get(s)) for s in CHANNEL_STEPS["Email"])
                cc_sent = sum(_is_marked(sv.get(s)) for s in CHANNEL_STEPS["Cold Call"])
                msgs += fu_sent + em_sent + cc_sent
                fus += max(0, fu_sent - 1) + max(0, em_sent - 1)  # follow-ups (no el 1°)
                if stage == "Nuevo lead":
                    nuevos += 1
                if stt.responded or l.outcome == "Respondió":
                    resp += 1
                won += l.outcome == "Won"; lost += l.outcome == "Lost"
                prosp += l.outcome == "Prospectar Después"
                black += l.outcome == "Blacklist"
        return dict(total=total, nuevos=nuevos, msgs=msgs, fus=fus, resp=resp,
                    won=won, lost=lost, prosp=prosp, black=black,
                    by_ind=by_ind, by_sen=by_sen, by_chan=by_chan, by_stage=by_stage)

    M = _dash_metrics(Path(xlsx_path).stat().st_mtime)
    a, b, c, d = st.columns(4)
    a.metric("Total leads", M["total"]); b.metric("Leads nuevos", M["nuevos"])
    c.metric("Mensajes enviados", M["msgs"]); d.metric("Follow-ups enviados", M["fus"])
    e, f, g, h = st.columns(4)
    e.metric("Respuestas", M["resp"]); f.metric("Won", M["won"])
    g.metric("Lost", M["lost"]); h.metric("Prospectar después", M["prosp"])
    st.metric("Blacklist", M["black"])

    st.subheader("Distribuciones")
    d1, d2 = st.columns(2)
    with d1:
        st.markdown("**Leads por industria**")
        st.bar_chart(pd.Series(M["by_ind"]))
        st.markdown("**Leads por canal**")
        st.bar_chart(pd.Series(M["by_chan"]))
    with d2:
        st.markdown("**Leads por seniority level**")
        st.bar_chart(pd.Series(M["by_sen"]))
        st.markdown("**Leads por etapa**")
        st.bar_chart(pd.Series(M["by_stage"]))

    st.divider()
    st.subheader("📊 Dashboard por industria")
    rows = []
    for ind in INDUSTRIES:
        leads = crm.all_leads(ind)
        total = len(leads)
        def cnt(step):
            return sum(_is_marked(l.stage_values.get(step)) for l in leads)
        contacted = sum(
            any(_is_marked(l.stage_values.get(s)) for s in
                ["Follow Up 1", "Email 1", "Cold Call 1"])
            for l in leads)
        outcome = lambda v: sum(l.outcome == v for l in leads)
        row = {"Industria": ind, "Total": total, "Contactados": contacted}
        for ch, steps in CHANNEL_STEPS.items():
            for s in steps:
                row[s] = cnt(s)
        for s in STAGE_COLS:
            row[s] = cnt(s)
        row["Won"] = outcome("Won"); row["Lost"] = outcome("Lost")
        row["Blacklist"] = outcome("Blacklist")
        row["Prospectar Después"] = outcome("Prospectar Después")
        row["Conversión %"] = round(100 * row["Won"] / total, 1) if total else 0.0
        row["% Contactado"] = round(100 * contacted / total, 1) if total else 0.0
        rows.append(row)
    df = pd.DataFrame(rows)

    st.subheader("Resumen (Total · % sobre leads · conversión)")
    st.dataframe(df, use_container_width=True, hide_index=True)

    # % sobre contactados para cada métrica
    st.subheader("Las 8 gráficas")
    df_i = df.set_index("Industria")
    groups = {
        "1 · Follow-ups": CHANNEL_STEPS["LinkedIn"],
        "2 · Emails": CHANNEL_STEPS["Email"],
        "3 · Cold Calls": CHANNEL_STEPS["Cold Call"],
        "4 · Meeting/RFQ/Quote": STAGE_COLS,
        "5 · Won": ["Won"], "6 · Lost": ["Lost"],
        "7 · Blacklist": ["Blacklist"], "8 · Prospectar Después": ["Prospectar Después"],
    }
    cols = st.columns(2)
    for i, (title, metrics) in enumerate(groups.items()):
        with cols[i % 2]:
            st.markdown(f"**{title}**")
            st.bar_chart(df_i[metrics])

    st.divider()
    st.subheader("📅 Reporte diario")
    st.caption("Genera un reporte con 8 hojas: Leads_Actualizados, Activity_Log, "
               "Follow_Up, Emails_Enviados, Respondieron, Blacklist, "
               "Pendientes_Mañana y Resumen_Dia.")

    fmt = st.radio("Formato", ["Excel (.xlsx)", "CSV (.zip)"], horizontal=True)

    # Vista previa (sin generar archivo): muestra cada hoja del reporte
    if st.checkbox("Previsualizar contenido"):
        data = crm.build_daily_report_data()
        for name, sh in data.items():
            with st.expander(f"{name} · {len(sh['rows'])} filas"):
                st.dataframe(pd.DataFrame(sh["rows"], columns=sh["headers"]),
                             use_container_width=True, hide_index=True)

    if st.button("Exportar reporte del día", type="primary"):
        out_dir = str(Path(xlsx_path).resolve().parent)
        kind = "csv" if fmt.startswith("CSV") else "xlsx"
        rep = crm.export_daily_report(out_dir=out_dir, fmt=kind)
        st.success(f"Reporte generado: {Path(rep).name}")
        mime = ("application/zip" if kind == "csv" else
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with open(rep, "rb") as fh:
            st.download_button("Descargar reporte", fh.read(),
                               file_name=Path(rep).name, mime=mime)

    st.divider()
    st.subheader("📊 Dashboard exportable (multi-hoja)")
    st.caption("Excel con hojas separadas por Industria, Seniority, Current Stage, "
               "Follow Up Step, Canal, Won, Lost, Prospectar después, Blacklist, "
               "Scheduled_Messages, Activity_Log y Notifications, más un Resumen. "
               "Incluye automáticamente los leads nuevos importados.")
    if st.button("Exportar dashboard completo", type="primary"):
        out_dir = str(Path(xlsx_path).resolve().parent)
        rep = crm.export_dashboard(out_dir=out_dir)
        st.success(f"Dashboard generado: {Path(rep).name}")
        with open(rep, "rb") as fh:
            st.download_button(
                "Descargar dashboard", fh.read(), file_name=Path(rep).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
