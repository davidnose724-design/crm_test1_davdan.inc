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
