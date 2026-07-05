# Manufacturing Prospecting CRM — App de prospección B2B asistida

CRM + asistente de prospección conectado a `Manufacturing_Prospecting_CRM_v4_nuevos_leads.xlsx`.
Carga leads por industria, arma campañas controladas con **aprobación humana**, respeta
límites e intervalos, gestiona follow-ups multicanal, detecta respuestas y mantiene el
dashboard del Excel cuadrado.

> **Postura de cumplimiento.** La app **no envía sola** y **no automatiza LinkedIn**. Prepara
> el mensaje, tú lo apruebas y lo envías por tu cuenta, y luego la app lo marca como "Enviado".
> No evade límites, no usa proxies, no rota cuentas, no hace scraping ni automatización
> agresiva. El intervalo y el tope por campaña son límites que la app respeta, no que burla.

---

## Cómo correr

```bash
pip install -r requirements.txt
streamlit run app.py
```

En la barra lateral indica la ruta del Excel (por defecto el nombre del archivo en el
directorio actual). La app crea junto al Excel un archivo `*.state.json` (estado lateral).

---

## Arquitectura

```
crm_core.py   ← motor (sin UI): se puede importar y testear solo
app.py        ← interfaz Streamlit (5 secciones)
*.state.json  ← estado que el Excel NO puede guardar (sidecar)
```

**Por qué un sidecar JSON.** El Excel no tiene columnas para: si un lead "Respondió" como
evento, la fecha de recontacto que tú das, los follow-ups *programados* con su timestamp, el
texto de cada mensaje, ni un log de actividad rico. Para **no alterar la estructura del
archivo** (orden de columnas y fórmulas del Dashboard son intocables), ese estado vive en el
sidecar, indexado por la URL de LinkedIn como llave primaria (con respaldo email o
nombre+empresa). El Excel sigue siendo la fuente de verdad de todo lo que sí tiene columna.

---

## Mapeo del Excel (verificado contra el archivo real)

- **11 hojas**: 9 de industria + `Dashboard` + `Lead Queue`. La hoja de software se llama
  `Software IT` (no "Software / IT").
- Las 9 hojas de industria comparten **40 columnas idénticas**.
- El estado de color vive en **`Outcome Status` = columna AH**. No existe "Lead Status".
- **"Message 1" = `Follow Up 1`** (columna P). El primer toque de LinkedIn no tiene columna
  propia; es el primer paso de la secuencia LinkedIn.
- Secuencias por canal: LinkedIn = `Follow Up 1–5` (P–T) · Email = `Email 1–5` (U–Y) ·
  Cold Call = `Cold Call 1–3` (Z–AB). Etapas: `Meeting` (AC), `RFQ` (AD), `Quote` (AE).
- **Detección de columnas tolerante a nombres**: normalización + tabla de sinónimos hacia
  nombres canónicos, anclada a la posición. Si cambias ligeramente un encabezado, se sigue
  reconociendo.

### Ordenamiento de la cola
`Lead Score` es categórico (A+…C) y `Seniority Level` es texto, así que el orden real usa:
`Seniority Rank` (numérico, 1 = C-Level) → prioridad de `Lead Score` (A+>A>B+>B>C+>C) →
`Company Name` → `Full Name`. Es la traducción fiel de "ordenar por Seniority Level y Lead
Score" a los datos que existen.

---

## Estados y colores (acordados)

Los colores salen del **formato condicional sobre AH**; la app nunca pinta celdas a mano.
A las 4 reglas existentes se suman 4 nuevas, sin tocar las viejas:

| Outcome Status        | Color de fila          | Origen     |
|-----------------------|------------------------|------------|
| Blacklist             | rojo claro             | existente  |
| Prospectar Después    | amarillo               | existente  |
| Won                   | verde                  | existente  |
| Lost                  | **morado** (opción C)  | existente  |
| Respondió             | azul claro             | **nuevo**  |
| Meeting / RFQ / Quote | verde claro tenue      | **nuevo**  |

La validación (dropdown) de AH se extiende de `Active,Blacklist,Prospectar Después,Won,Lost`
a incluir `Respondió,Meeting,RFQ,Quote`.

### Reglas de bloqueo / pausa
- **Blacklist** → nunca se vuelve a contactar (bloqueo duro).
- **Won** → no más prospección.
- **Lost** → detiene la secuencia activa.
- **Prospectar Después** → sin envíos hasta la **fecha de recontacto que tú das** (nunca se
  calcula sola).
- **Respondió** → pausa **todos** los follow-ups hasta que clasifiques y decidas.

### Meeting / RFQ / Quote sin romper el dashboard
Se marca **la columna de etapa** (AC/AD/AE), que es lo que el Dashboard cuenta por posición,
**y además** se escribe `Outcome Status` para teñir la fila de verde claro. Como el Dashboard
lee las etapas por posición (no por AH), los conteos siguen cuadrando. Estrictamente mejor
que la opción que rompía el conteo.

---

## "Respondió" = ambas
Cuando un lead responde: (1) se dispara la notificación con nombre, empresa, industria,
cargo, canal, mensaje y acción sugerida; se pausan **todos** los canales; **y** (2) se escribe
`Outcome Status = "Respondió"` con fila azul. No es terminal: queda en pausa hasta que
clasifiques. La clasificación la decides tú; la app solo **sugiere**:

| Clasificación        | Efecto                                  |
|----------------------|-----------------------------------------|
| No contactar         | Blacklist (rojo)                        |
| Pide cotización      | marca RFQ (verde claro)                 |
| Quiere reunión       | marca Meeting (verde claro)             |
| No interesado        | Lost (morado)                           |
| Interesado / Más info / Otro | queda en pausa; nota en `Notes` |

---

## Follow-ups y escalamiento de canal
- Cadena FU1→FU5: cada paso se programa **X días** después del anterior, **solo si no hubo
  respuesta**. El reloj del siguiente arranca cuando el actual se marca "Enviado".
- Aprobación **individual o por lote** (sección Follow-ups).
- Escalamiento **sugerido** (nunca automático): LinkedIn sin respuesta tras FU3 → Email 1;
  Email 2 sin respuesta → Cold Call 1; Cold Call 3 sin respuesta → Lost o Prospectar Después.
- Responder por **cualquier** canal pausa **los tres**.
- Límites e intervalo se aplican **globales por lead/campaña**, no por canal aislado, para no
  exceder lo configurado al combinar canales.

---

## Dashboard
La hoja `Dashboard` del Excel ya viene viva (COUNTIF/COUNTA por posición) y **no se
reconstruye**: la app solo escribe en las celdas correctas y los conteos se actualizan al
recalcular. La sección Dashboard de la app además muestra, por industria: Total, % sobre
leads, % sobre contactados, conversión (Won ÷ total), y **8 gráficas** (Follow-ups, Emails,
Cold Calls, Meeting/RFQ/Quote, Won, Lost, Blacklist, Prospectar Después). "Contactados" =
leads con ≥1 toque en cualquier canal (conteo único, igual que el `Activity Rate` del
archivo). Las gráficas viven en la app; **no se incrustan en el Excel** para no alterarlo.

> Nota sobre el archivo original: el KPI "Total Leads" del Dashboard usa `COUNTA …$A$2:$A$5000`,
> que puede inflar el total si hay celdas residuales en ese rango. Es una fórmula
> **preexistente del archivo**; no se modificó para respetar la regla de no tocar el formato.
> Los conteos por industria (fila por industria) son correctos.

---

## Recalcular fórmulas
Tras escribir con la app, abre el Excel en Excel/LibreOffice (recalcula solo) o corre el
script de recálculo incluido en el entorno de skills. La app preserva fórmulas y formato; solo
escribe valores en celdas de datos.

## Demo incluida
`Manufacturing_Prospecting_CRM_v4_DEMO.xlsx` muestra una campaña aplicada (30 FU1 en
Automotive + 10 Email 1 en Pharmaceutical) con los 7 desenlaces y sus colores, ya recalculada
y con cero errores de fórmula.

---

# Actualización v2 — Importador, Activity_Log, reporte diario, backups

Esta versión agrega seis bloques de funcionalidad **sin romper** lo anterior.

## Archivos modificados
- **`crm_core.py`** (motor): se añadieron importador de leads con deduplicación,
  asignación automática (industria/seniority), inferencia de seniority desde el Job
  Title, hoja `Activity_Log` en el Excel, reporte diario, backup automático, guardado
  con manejo de "archivo abierto", y re-anclaje seguro de colores y validaciones para
  que los leads importados hereden formato y dropdowns.
- **`app.py`** (interfaz): nueva sección **📤 Importar leads**, **🔁 Follow-ups**
  reescrita con vista general (estado, último mensaje, último contacto, siguiente acción
  y fecha sugerida), botón **Exportar reporte del día** en el Dashboard, canales LinkedIn/
  Email/Llamada/WhatsApp, y avisos claros cuando el Excel está abierto.
- **`README.md`**: esta sección.

## 1. Importador de leads (📤 Importar leads)
Sube `.xlsx`/`.csv`, muestra preview, **auto-detecta** columnas (full/first/last name,
company, title, industry, location, LinkedIn, email, phone) y permite **mapear a mano**
las que no reconozca. Al importar, agrega los leads al Excel principal respetando las 40
columnas, y **evita duplicados** por LinkedIn URL, email, nombre+empresa o teléfono.

Se agrega cada lead **con lo que tenga disponible**: si le falta nombre completo, email o
teléfono, igual se importa y queda listo para trabajar con lo que sí tenga (solo se llenan
los espacios posibles; lo demás queda vacío). Únicamente se omiten las filas totalmente
vacías y los duplicados. Si un lead no trae **ninguna** señal de identidad
(LinkedIn/email/nombre/empresa/teléfono), se agrega igual, pero la app avisa que no podrá
deduplicarse ni seguirse individualmente.

## 2. Asignación automática
Cada lead se enruta a la hoja de su **industria** (o a una por defecto si no viene), se
infiere **Seniority Level/Rank** desde el Job Title cuando falta, y se marca `Active`. El
formato original del Excel se conserva; los leads nuevos **heredan colores y dropdowns**
porque la app re-ancla los rangos de formato condicional y validación a las filas nuevas.

> Nota: el `Seniority Rank` inferido para leads importados usa una escala propia
> (C-Level=1 … Other=9). Puede diferir levemente de los ranks que ya traían los leads
> originales; solo afecta el orden relativo de los importados.

## 3. Vista de follow-ups (🔁 Follow-ups)
Tabla con quién está en cada paso, mostrando nombre, empresa, puesto, industria, último
mensaje enviado, fecha de último contacto, siguiente acción, fecha sugerida del próximo
follow-up y **estado** (pendiente / enviado / respondió / no respondió / blacklist /
prospectar después). Debajo, aprobación por lote y sugerencias de escalamiento.

## 4. Registro de actividad (hoja `Activity_Log`)
Cada vez que marcas un mensaje como enviado, se agrega una fila a la hoja **`Activity_Log`**
del Excel con: timestamp, persona, empresa, industria, fecha, **canal** (LinkedIn / email /
llamada / WhatsApp), número de follow-up, paso, mensaje y resultado.

## 5. Exportación diaria
Botón **Exportar reporte del día** (en el Dashboard) genera el reporte con **8 hojas**,
exportable y visible como **Excel** (`CRM_reporte_YYYY-MM-DD.xlsx`) o como **CSV**
(`CRM_reporte_YYYY-MM-DD.zip`, con un .csv por hoja). En la app puedes **previsualizar**
cada hoja antes de descargar. No modifica el CRM; es un archivo aparte. Hojas:

1. `Leads_Actualizados` — todos los leads con estado actual.
2. `Activity_Log` — historial de actividad acumulado.
3. `Follow_Up` — quién va en qué Follow Up (paso actual y los enviados).
4. `Emails_Enviados` — a quién se le envió cada Email y su correo.
5. `Respondieron` — quiénes respondieron, canal y mensaje.
6. `Blacklist` — leads que pidieron no ser contactados, con razón.
7. `Pendientes_Mañana` — follow-ups programados para mañana o antes.
8. `Resumen_Dia` — totales: leads cargados, contactados hoy, Follow Ups enviados hoy
   (con desglose por paso), correos enviados hoy (con desglose), no accedieron a recibir
   mensaje (Blacklist), respondieron, Blacklist y pendientes mañana.

## 6. Seguridad y estabilidad
- **No envía nada automáticamente**: solo prepara mensajes y registra actividad manual.
- **Backup automático** antes de cada guardado, en la carpeta `backups/` junto al Excel,
  con nombre `<archivo>_YYYYMMDD_HHMMSS.xlsx`.
- Si el Excel **está abierto** en otra app, el guardado no truena: muestra una advertencia
  clara ("ciérralo y reintenta") y conserva tus cambios en memoria y el backup intacto.

---

# Cómo correr en Windows 11

Abre **PowerShell** (o Símbolo del sistema) en la carpeta del proyecto. Necesitas Python
3.10+ instalado (desde python.org o Microsoft Store).

```powershell
:: 1. Entrar a la carpeta del proyecto
cd C:\ruta\a\crm

:: 2. Crear y activar un entorno virtual (recomendado)
python -m venv .venv
.\.venv\Scripts\activate

:: 3. Instalar dependencias
pip install -r requirements.txt

:: 4. Colocar el Excel en la misma carpeta (o indicar su ruta en la barra lateral)
::    Manufacturing_Prospecting_CRM_v4_nuevos_leads.xlsx

:: 5. Lanzar la app (se abre en el navegador en http://localhost:8501)
streamlit run app.py
```

Para volver a usarla otro día: entra a la carpeta, activa el entorno
(`.\.venv\Scripts\activate`) y corre `streamlit run app.py`.

Si PowerShell bloquea la activación del entorno, ejecútalo una vez como:
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

---

# Actualización v3 — Importación avanzada, prioridad por seniority, búsqueda de respuestas, estados agrupados y dashboard ampliado

## Archivos modificados
- **`crm_core.py`**: columnas nuevas `Value Proposition` y `Pain Point` que se crean
  **al final** de la hoja (posición ≥41, sin mover nada que el Dashboard lea por
  posición) vía `ensure_columns`; `make_import_template` (plantilla); `lead_priority`,
  `lead_stage`, `lead_channel` (cálculos en vivo, no se almacenan); `apply_response_result`
  (mapeo de resultado de respuesta a acciones); el importador ahora respeta el Seniority
  Level del archivo y escribe Notes/Value Proposition/Pain Point.
- **`app.py`**: las 5 secciones (Importar / Campañas / Respuestas / Estados / Dashboard)
  se ampliaron como se detalla abajo, sin tocar el núcleo de abrir/mostrar/guardar.
- **`README.md`**: esta sección.

## Qué cambió por sección
1. **Importar leads**: explicación del formato, **tabla ilustrativa**, botón para
   descargar `plantilla_importacion_leads.xlsx` (con First/Last/Full Name, Job Title,
   Company, Industry, Seniority Level, LinkedIn URL, Email, Phone, Location, Value
   Proposition, Pain Point, Notes), **carga de varios archivos a la vez**, **preview
   consolidado**, conteo de duplicados y dedup por LinkedIn / Email / Full Name+Company.
2. **Campañas**: apartado **Seniority Level** que lee el del Excel, **filtra** la campaña
   por seniority y calcula **prioridad** (Alta / Media-Alta / Media / Baja) mostrada junto
   a cada lead y resumida en la cola.
3. **Respuestas**: **buscador/autocompletado** por nombre/email/LinkedIn/empresa; al
   seleccionar, autocompleta nombre, empresa, puesto, industria, seniority, email,
   LinkedIn y follow-up actual; solo capturas mensaje, canal, **resultado** (Won / Lost /
   Respondió interesado / Respondió no interesado / Prospectar después / Blacklist /
   Reunión agendada / Cotización solicitada) y nota. Al guardar, actualiza la tabla
   principal, el dashboard y el **Activity_Log**.
4. **Estados**: **tabla resumen** agrupada por Industria, Seniority, Paso actual y Canal
   actual, con total de leads, contactados, pendientes, respondieron, won, lost, blacklist
   y próximo follow-up, más **filtros** por industria, seniority, canal y estado. Debajo,
   el editor de un lead.
5. **Dashboard**: métricas headline (total, nuevos, mensajes, follow-ups, respuestas, won,
   lost, prospectar después, blacklist) y desgloses **por industria, seniority, canal y
   etapa**, además del resumen por industria y las 8 gráficas previas.

## Compatibilidad y seguridad (sin cambios de comportamiento)
No se automatiza ningún envío; solo se registra lo que hiciste manualmente. La carga del
Excel actual no cambia. El mapeo tolera columnas faltantes y las necesarias se crean al
final automáticamente. Backup automático antes de guardar y aviso claro si el Excel está
abierto siguen activos.

---

# Actualización v4 — Notifications, leads disponibles en Campaña, estados manuales y Follow-ups editables

## Archivos modificados
- **`crm_core.py`**: hoja **`Notifications`** (crear/leer/marcar visto), `add_notification`,
  `scan_due_followups` (alertas de follow-up vencido), `set_manual_state` (los 12 estados),
  `next_action` (siguiente acción sugerida). `apply_response_result` ahora también crea
  notificación.
- **`app.py`**: nueva sección **🔔 Notifications**; tabla de **leads disponibles con
  filtros** en Campaña; **12 estados manuales** + filtros de fase/resultado en Estados;
  Follow-ups con **tabla completa filtrable** y **edición por lead** (marcar enviado,
  cambiar canal, cambiar estado, nota).
- **`README.md`**: esta sección.

## Qué cambió por sección
1. **Campaña**: antes de prospectar, tabla de leads disponibles con nombre, empresa,
   puesto, seniority, industria, LinkedIn, email, prioridad, estado actual, último contacto
   y siguiente acción, con filtros por industria, seniority, empresa, prioridad y estado.
2. **Notifications** (nueva): alertas por respuestas registradas, cambios de estado y
   follow-ups vencidos (ej. "… debe recibir Follow Up 2 hoy", "… pidió cotización", "Lead
   marcado como Won"). Cada una con fecha, lead, empresa, canal, tipo de evento, resumen y
   estado actual; se pueden marcar **vista/pendiente**. Se guardan en la hoja `Notifications`
   del Excel. No se conecta con LinkedIn/Gmail/WhatsApp: todo se basa en registros manuales.
3. **Estados**: además de ver el estado, se cambia manualmente a cualquiera de los 12
   (Nuevo lead, Mensaje inicial enviado, Follow Up 1–3, Respondió, Won, Lost, Prospectar
   después, Blacklist, Reunión agendada, Cotización solicitada). Cada cambio actualiza la
   tabla, registra en Activity_Log, crea notificación y actualiza el dashboard. Filtros por
   industria, seniority, etapa/fase, canal y resultado final.
4. **Follow-ups**: tabla filtrable por fase (Follow Up 1/2/3 pendiente/enviado,
   respondieron, prospectar después) con nombre, empresa, puesto, seniority, industria,
   follow-up actual, último contacto, canal, mensaje usado, siguiente follow-up y estado.
   Desde la tabla se puede marcar follow-up como enviado, cambiar canal, cambiar estado y
   agregar nota; todo queda en Activity_Log.

Compatibilidad y seguridad sin cambios: sin envíos automáticos, columnas faltantes se
crean al final, backup antes de guardar y aviso si el Excel está abierto.

---

# Actualización v5 — Campañas filtradas + programación, importación inteligente, dashboard multi-hoja, Notifications con modos e integraciones

## Archivos modificados / nuevos
- **`crm_core.py`**: columnas nuevas `Lead Warmth Level`, `Priority`, `Recommended Channel`,
  `Current Stage` (creadas al final, ≥43); clasificadores `classify_warmth`,
  `recommend_channel`, `classify_stage`, `classify_lead`; hoja **`Scheduled_Messages`** con
  `schedule_message`/`read_scheduled`/`update_schedule_status`; `edit_lead_fields`;
  `build_dashboard_data`/`export_dashboard` (multi-hoja). El importador ahora clasifica y
  persiste warmth/priority/canal/stage.
- **`integrations.py`** (NUEVO): arquitectura de integraciones (Modo B) sin credenciales,
  sin envíos, sin scraping. LinkedIn = solo manual/CSV; Gmail = stub Gmail API/OAuth;
  WhatsApp = stub WhatsApp Business API; Cold Call = manual. Las credenciales, si algún día
  existen, se leerían de variables de entorno.
- **`app.py`**: Campaña con **"Crear campaña filtrada"** (industria, canal, seniority,
  empresa, país/ubicación, estado, calentamiento, nº de leads) + selección + **programación
  de mensajes** y **agenda** (Draft/Scheduled/Sent/Cancelled) con marcado manual de enviado;
  Importar con clasificación automática; **Notifications** con pestañas Modo A (registro
  manual con buscador, tipo de respuesta y resultado) / Alertas / Modo B (estado de
  integraciones); **Estados** con campos editables (warmth, priority, canal, follow-up step,
  próxima acción, fecha próximo contacto, notas); **Dashboard** con botón de exportación
  multi-hoja.
- **`README.md`**: esta sección.

## Cómo se guardan las hojas
- **Scheduled_Messages**: hoja con Schedule ID, Lead ID, Full Name, Company, Seniority,
  Industry, Channel, Message, Scheduled DateTime, Status (Draft/Scheduled/Sent/Cancelled),
  Sent DateTime, Follow Up Step, Notes. Al marcar **Sent** (manual) se sella la fecha real y
  se marca la etapa del lead. **Nunca se envía solo.**
- **Activity_Log** y **Notifications**: como antes, append por evento; las notificaciones se
  marcan vista/pendiente.

## Seguridad
Sin envíos automáticos, sin scraping, sin credenciales en el código; integraciones futuras
solo por APIs oficiales. Compatibilidad con el Excel actual, columnas faltantes creadas al
final, backup antes de guardar y aviso si el Excel está abierto.

---

# Actualización v6 — Integración Gmail API (OAuth 2.0 vía st.secrets)

## Archivos nuevos / modificados
- **`gmail_service.py`** (NUEVO): módulo aislado de Gmail — OAuth 2.0 (authorization code
  flow), lectura de correos recientes (readonly), matching remitente→lead, envío de
  respuestas con MIME correcto (threading con In-Reply-To/References). Degradación limpia
  si faltan librerías o secrets.
- **`app.py`**: pestaña **📧 Gmail** dentro de Notifications, con botón **Conectar Gmail**,
  lectura de correos recientes, **detección de respuestas de leads** (cruza remitente vs
  Email/Gmail del CRM), registro de la respuesta (pausa follow-ups + notificación +
  Activity_Log) y **responder desde la app** (envío manual con botón, registrado en
  Activity_Log y Notifications).
- **`secrets.toml.example`** (NUEVO): plantilla con placeholders.
- **`.gitignore`** (NUEVO): excluye `.streamlit/secrets.toml`, estado y backups.
- **`requirements.txt`**: + google-auth, google-auth-oauthlib, google-api-python-client.
- **`integrations.py`**: Gmail reporta "API lista" cuando hay `[google]` en st.secrets.

## Configuración (una sola vez)
1. En **Google Cloud Console** → APIs & Services → habilita **Gmail API**.
2. En **Credentials** → OAuth client ID → tipo **Web application** → agrega tu
   redirect URI EXACTO (tu URL de Streamlit Cloud).
3. Pantalla de consentimiento (OAuth consent screen): agrega los scopes
   `openid`, `userinfo.email`, `gmail.readonly`, `gmail.send`, y tu cuenta como
   test user si la app está en modo Testing.
4. Secrets:
   - **Streamlit Cloud**: App → Settings → Secrets → pega el bloque `[google]`
     con tus valores (client_id, client_secret, redirect_uri).
   - **Local**: copia `secrets.toml.example` a `.streamlit/secrets.toml` y edítalo.
5. `pip install -r requirements.txt` y listo: pestaña 📧 Gmail → **Conectar Gmail**.

## Seguridad
- **Cero credenciales en el código o en GitHub**: todo vive en st.secrets; el
  `.gitignore` protege `secrets.toml`; la plantilla solo trae placeholders.
- Si tu client_secret se compartió en algún chat/captura/commit, **rótalo** en Google
  Cloud Console (Credentials → tu OAuth client → Reset secret).
- Scopes mínimos (`gmail.readonly` + `gmail.send`); tokens solo en memoria de sesión
  (st.session_state), no se escriben a disco ni al repo.
- La app **nunca envía correos sola**: cada envío requiere que presiones el botón.
- Solo API oficial de Google; sin scraping.

---

# Actualización v7 — Reprogramación de follow-ups, Cold Call, colores configurables y estados personalizados

## Archivos modificados
- **`crm_core.py`**: `reschedule_followup` (+ hoja **`Follow_Up_History`** con las 12
  columnas), `register_cold_call` (9 resultados con sus acciones), configuración de
  colores en hoja **`State_Color_Config`** (`read_state_colors`/`upsert_state_color`),
  `resolve_display_state` y `apply_state_colors` (pintado con openpyxl `PatternFill`,
  fila completa o solo celda de estado, blanco = limpiar). Columnas nuevas al final:
  `Follow Up Step`, `Follow Up Channel`, `Next Follow Up Time`, `Follow Up Reason`,
  `Owner/User`.
- **`app.py`**: formulario de **reprogramación reutilizable** presente en Follow-ups,
  Gmail (con acciones rápidas: Won/Lost/RFQ/Contactar después/Blacklist), Estados y
  Notifications; **formulario de Cold Call** (fecha, resultado, nota, próxima acción,
  próxima fecha); nueva sección **🎨 Configuración de Estados** (ver estados, elegir
  color, fila-o-celda, crear estados personalizados con categoría/acción/prioridad,
  botón **Aplicar colores ahora**); **Dashboard** con follow-ups vencidos/hoy/próximos
  7 días, RFQ activos, cold calls pendientes, gmail pendientes y leads por estado/color.

## Cómo se aplican los colores (y el conflicto con el formato heredado)
El Excel original trae **formato condicional** por `Outcome Status` (Won verde, Lost
morado, Blacklist rojo, Prospectar amarillo…), que en Excel **prevalece visualmente**
sobre el pintado directo. La sección de Configuración incluye el checkbox
**"Reemplazar formato condicional heredado"**:
- **Desmarcado** (default): el CF clásico sigue mandando en esos estados; el pintado
  estático cubre el resto (Follow Up pendiente, Mensaje enviado, estados custom…).
- **Marcado**: se eliminan las reglas CF y tu configuración es la única fuente de color
  (defaults sembrados: Won verde 00B050, Lost rojo FF0000, RFQ azul, Reunión morado,
  Prospectar amarillo, Respondió interesado verde claro, No interesado naranja,
  Blacklist gris, Follow Up pendiente amarillo claro, Mensaje enviado azul claro,
  Nuevo lead blanco). Siempre con backup previo.

## Historial de reprogramaciones
Cada reprogramación (desde cualquier sección) agrega una fila a `Follow_Up_History`
con timestamp, lead, fecha/canal anteriores y nuevos, motivo, nota, usuario y status;
además actualiza la fila del lead (Next Follow-up, hora, canal, motivo, owner),
re-sincroniza el scheduler interno, registra en Activity_Log y crea una notificación.
El historial por lead se consulta en la sección Estados.

---

# Actualización v8 — Notificaciones configurables, Email Campaigns (auto-send opt-in), manejo de respuestas/rebotes, LinkedIn Manager asistido y Workflow Config

## Archivos modificados
- **`crm_core.py`**: hojas `Notification_Settings` y `Workflow_Config` (clave-valor con
  defaults), `scan_no_response` (alertas por horas configurables: lead sin respuesta,
  campaña lista para FU2/FU3, follow-up vencido, lead debe descartarse/reprogramarse);
  hoja **`Campaigns`** con el workflow asistido completo (`create_campaign`,
  `next_pending_lead`, `mark_campaign_sent` con avance automático de etapa y cálculo del
  siguiente follow-up según config, auto-creación de "Follow Up N+1 - <campaña>" al
  terminar excluyendo respondidos/blacklist/Won, `reschedule_campaign`,
  `scan_campaign_reminders`); clasificador de respuestas Gmail (`classify_email_reply`:
  bounce/blacklist/not_interested/rfq/meeting/later/interested), `handle_email_event`
  (Email Bounced/Blocked/Do Not Contact detienen todo y cancelan la cola),
  `process_due_emails` (cola con tope diario, skip+cancel de bloqueados, marca Sent antes
  de enviar para evitar duplicados y revierte si el envío falla; función de envío
  inyectable). Outcome nuevos: `Do Not Contact`, `Email Bounced`, `Blocked` (bloqueo duro).
- **`gmail_service.py`**: `send_email` (correo nuevo para campañas).
- **`app.py`**: secciones **📧 Email Campaigns** (filtros, variables {first}/{name}/
  {company}/{title}/{industry}, cola en Scheduled_Messages, toggle **Auto-send enabled**
  con advertencia de límites/spam/cumplimiento, "Procesar cola ahora", lectura y
  clasificación de respuestas), **💼 LinkedIn Manager** (cola inteligente de un-lead-a-la-
  vez con todos los datos, "Abrir perfil", mensaje sugerido editable/copiable, "Marcar
  mensaje como enviado" → registra fecha/hora/usuario/canal/mensaje, avanza al siguiente,
  reprogramación mañana/3 días/semana/personalizada, "Iniciar campaña", registro de
  respuesta) y **⚙️ Workflow Config** (horas entre FU1→2/2→3/3→4, máx FUs, canal default,
  horario/días permitidos, intervalo, y "Follow Up Notification Settings" con 24/30/48/
  72/76/personalizado). Notifications con 3 escaneos; Dashboard con las 11 métricas nuevas.

## Cómo se evita el desastre
- **Duplicados**: la cola marca `Sent` ANTES del envío real (y revierte si falla), el
  guard de etapa rechaza celdas ya marcadas y el tope diario se verifica por envío.
- **Pausa al responder**: `Respondió` entra al guard `_is_blocked`; el lead sale de la
  cola de campaña, del scheduler y del auto-send (sus mensajes pendientes se cancelan).
- **Bounce/Blocked/Blacklist/Do Not Contact/Lost**: bloqueo duro — ningún flujo (manual
  o automático) puede volver a enviarles; sus mensajes en cola se cancelan y se notifica.

## Honestidad técnica del auto-send
Streamlit no corre en segundo plano: el envío automático procesa los correos vencidos
cuando la app está abierta (botón "Procesar cola ahora" con el toggle activo). Para envío
24/7 haría falta un worker externo (cron/Cloud Scheduler), que puedo agregar aparte.

---

# Actualización v9 — Gmail Campaigns (login, sender, preview editable, Gmail Follow Up 1, respuestas)

## Archivos modificados
- **`crm_core.py`**: hojas **`Gmail_Accounts`** (upsert de cuenta conectada con Last
  Sync), **`Gmail_Campaigns`**, **`Gmail_Campaign_Leads`** y **`Gmail_Follow_Ups`**
  (`ensure_gmail_sheets` las crea si faltan, junto con Scheduled_Messages/
  Follow_Up_History/Notifications/Activity_Log); `personalize` con variables
  {{first_name}}, {{full_name}}, {{company}}, {{job_title}}, {{industry}},
  {{seniority_level}}; `gmail_followup1_candidates` (recibieron email inicial, sin
  respuesta, no bloqueados, ≥ horas configuradas, con horas transcurridas, último
  mensaje, asunto anterior y próxima acción). Fix general: los registros de envío ahora
  guardan el timestamp real del envío (antes usaban el momento del registro).
- **`app.py`**: sección **📮 Gmail Campaigns** con 5 pestañas — Conexión (correo
  conectado, estado, última sincronización, desconectar; cuenta guardada en
  Gmail_Accounts), Crear campaña (los 13 filtros con exclusiones automáticas de
  blacklist/lost/do-not-contact/sin-email, asunto/mensaje con variables, inicio,
  intervalo, límite diario, confirmación manual y auto-send), Preview/cola (tabla con
  las 10 columnas, editar mensaje por lead, excluir, guardar borrador, programar,
  enviar lead por lead, enviar seleccionados con confirmación; registro en
  Gmail_Campaigns + Gmail_Campaign_Leads), **Gmail Follow Up 1** (candidatos por
  tiempo configurado con todas las columnas pedidas; seleccionar, editar, programar,
  enviar con confirmación, reprogramar, nota; registro en Gmail_Follow_Ups) y
  Respuestas ("Buscar respuestas recientes" → marca Respondió, detiene follow-ups,
  clasificación de 9 opciones incluyendo Won y Do Not Contact). El formulario de
  reprogramación ahora tiene opciones rápidas: mañana / 24h / 48h / 72h / 7 días /
  personalizada.

Reglas: credenciales solo en st.secrets, nunca a leads bloqueados o sin email, la
secuencia se detiene al responder, todo va a Activity_Log, backup antes de guardar.
