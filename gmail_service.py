"""
gmail_service.py — Integración Gmail vía API oficial (OAuth 2.0).

Postura de seguridad (innegociable):
- CERO credenciales en el código. Todo viene de st.secrets["google"]
  (.streamlit/secrets.toml en local, o el Secrets manager en Streamlit Cloud).
- .streamlit/secrets.toml NUNCA se sube a GitHub (ver .gitignore).
- Solo API oficial de Google. Sin scraping.
- La app NUNCA envía correos sola: cada envío lo dispara el usuario con un botón.
- Scopes mínimos: openid, userinfo.email, gmail.readonly, gmail.send.

Requiere: google-auth, google-auth-oauthlib, google-api-python-client
(en requirements.txt). Si faltan, la app lo indica sin romperse.
"""

from __future__ import annotations
import base64
import datetime as dt
from email.mime.text import MIMEText
from email.utils import parseaddr

SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
]

_IMPORT_ERROR = None
try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import Flow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except Exception as e:
    _IMPORT_ERROR = e
    Credentials = Flow = build = HttpError = None


def libs_available() -> bool:
    return _IMPORT_ERROR is None


def libs_error() -> str:
    return ("Faltan librerías de Google. Instala con: "
            "pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client"
            if _IMPORT_ERROR else "")


# --------------------------------------------------------------------------- #
# Configuración desde st.secrets (nunca hardcodeada)
# --------------------------------------------------------------------------- #

def read_google_secrets(st):
    """Lee client_id/client_secret/redirect_uri de st.secrets['google'].
    Devuelve (dict, None) o (None, mensaje_de_error)."""
    try:
        g = st.secrets["google"]
        cfg = {
            "client_id": g["client_id"],
            "client_secret": g["client_secret"],
            "redirect_uri": g["redirect_uri"],
        }
        if not all(cfg.values()):
            return None, "st.secrets['google'] tiene campos vacíos."
        return cfg, None
    except Exception:
        return None, ("No encontré st.secrets['google']. Crea .streamlit/secrets.toml "
                      "(local) o configura Secrets en Streamlit Cloud con: "
                      "[google] client_id, client_secret, redirect_uri.")


def _client_config(cfg):
    return {
        "web": {
            "client_id": cfg["client_id"],
            "client_secret": cfg["client_secret"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [cfg["redirect_uri"]],
        }
    }


# --------------------------------------------------------------------------- #
# OAuth (authorization code flow para app web)
# --------------------------------------------------------------------------- #

def build_auth_url(cfg) -> tuple[str, str]:
    """Devuelve (auth_url, state). El usuario abre la URL, autoriza, y Google
    redirige a redirect_uri con ?code=..."""
    flow = Flow.from_client_config(_client_config(cfg), scopes=SCOPES,
                                   redirect_uri=cfg["redirect_uri"])
    auth_url, state = flow.authorization_url(
        access_type="offline", include_granted_scopes="true", prompt="consent")
    return auth_url, state


def exchange_code(cfg, code: str):
    """Intercambia el código de autorización por credenciales. Devuelve un dict
    serializable para guardar en st.session_state (no en disco, no en Git)."""
    flow = Flow.from_client_config(_client_config(cfg), scopes=SCOPES,
                                   redirect_uri=cfg["redirect_uri"])
    flow.fetch_token(code=code)
    c = flow.credentials
    return {
        "token": c.token, "refresh_token": c.refresh_token,
        "token_uri": c.token_uri, "client_id": c.client_id,
        "client_secret": c.client_secret, "scopes": list(c.scopes or SCOPES),
    }


def credentials_from_dict(d):
    return Credentials(**d)


def gmail_client(creds_dict):
    return build("gmail", "v1", credentials=credentials_from_dict(creds_dict),
                 cache_discovery=False)


def whoami(creds_dict) -> str:
    """Email de la cuenta conectada."""
    svc = gmail_client(creds_dict)
    prof = svc.users().getProfile(userId="me").execute()
    return prof.get("emailAddress", "")


# --------------------------------------------------------------------------- #
# Lectura de correos recientes y detección de respuestas de leads
# --------------------------------------------------------------------------- #

def _header(headers, name):
    for h in headers or []:
        if h.get("name", "").lower() == name.lower():
            return h.get("value", "")
    return ""


def list_recent_messages(creds_dict, days=7, max_results=25, query_extra=""):
    """Lista correos recientes del inbox (solo lectura). Devuelve lista de dicts:
    id, threadId, from_email, from_name, subject, date, snippet, message_id_hdr."""
    svc = gmail_client(creds_dict)
    q = f"in:inbox newer_than:{days}d {query_extra}".strip()
    resp = svc.users().messages().list(userId="me", q=q,
                                       maxResults=max_results).execute()
    out = []
    for m in resp.get("messages", []) or []:
        full = svc.users().messages().get(
            userId="me", id=m["id"], format="metadata",
            metadataHeaders=["From", "Subject", "Date", "Message-ID"]).execute()
        headers = full.get("payload", {}).get("headers", [])
        raw_from = _header(headers, "From")
        name, email = parseaddr(raw_from)
        out.append({
            "id": m["id"], "threadId": full.get("threadId", ""),
            "from_email": (email or "").lower(), "from_name": name or email,
            "subject": _header(headers, "Subject"),
            "date": _header(headers, "Date"),
            "snippet": full.get("snippet", ""),
            "message_id_hdr": _header(headers, "Message-ID"),
        })
    return out


def match_messages_to_leads(messages, crm):
    """Cruza remitentes con los leads del CRM por Email/Gmail.
    Devuelve lista de (message, lead) solo para coincidencias."""
    email_index = {}
    for sheet in crm.maps:
        for lead in crm.all_leads(sheet):
            if lead.email:
                email_index[str(lead.email).strip().lower()] = lead
    matched = []
    for msg in messages:
        lead = email_index.get(msg["from_email"])
        if lead is not None:
            matched.append((msg, lead))
    return matched


# --------------------------------------------------------------------------- #
# Envío de respuesta (SIEMPRE disparado manualmente por el usuario)
# --------------------------------------------------------------------------- #

def send_email(creds_dict, to_email, subject, body):
    """Envía un correo nuevo (para campañas). Solo se invoca desde la cola con
    'Auto-send enabled' activado explícitamente por el usuario, o desde un botón."""
    mime = MIMEText(body, "plain", "utf-8")
    mime["To"] = to_email
    mime["Subject"] = subject
    raw = base64.urlsafe_b64encode(mime.as_bytes()).decode("utf-8")
    svc = gmail_client(creds_dict)
    sent = svc.users().messages().send(userId="me", body={"raw": raw}).execute()
    return sent.get("id", "")


def send_reply(creds_dict, to_email, subject, body,
               thread_id=None, in_reply_to=None):
    """Envía una respuesta. Debe llamarse solo desde un botón que el usuario
    presiona; la app nunca lo invoca de forma automática."""
    mime = MIMEText(body, "plain", "utf-8")
    mime["To"] = to_email
    mime["Subject"] = subject if subject.lower().startswith("re:") else f"Re: {subject}"
    if in_reply_to:
        mime["In-Reply-To"] = in_reply_to
        mime["References"] = in_reply_to
    raw = base64.urlsafe_b64encode(mime.as_bytes()).decode("utf-8")
    payload = {"raw": raw}
    if thread_id:
        payload["threadId"] = thread_id
    svc = gmail_client(creds_dict)
    sent = svc.users().messages().send(userId="me", body=payload).execute()
    return sent.get("id", "")
