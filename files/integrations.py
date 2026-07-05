"""
integrations.py — Arquitectura de integraciones de canal (Modo B).

IMPORTANTE — postura de seguridad:
- Este módulo NO envía mensajes, NO lee mensajes privados y NO hace scraping.
- NO contiene credenciales. Las credenciales, cuando existan, deben venir de
  variables de entorno o de un gestor de secretos del usuario, nunca del código.
- Son *stubs* (interfaces) para una integración futura vía APIs OFICIALES.
  Mientras no se implementen, la app funciona 100% en modo manual.

Canales:
- LinkedIn  -> SOLO registro manual o importación CSV manual (sin API de mensajes,
               sin scraping). No hay integración automática por política de LinkedIn.
- Gmail     -> preparado para Gmail API oficial vía OAuth 2.0 (no incluido).
- WhatsApp  -> preparado para WhatsApp Business API oficial (no incluido).
- Cold Call -> registro manual del resultado de la llamada.
"""

from __future__ import annotations
import os
from dataclasses import dataclass


class IntegrationNotConfigured(Exception):
    """Se lanza cuando se intenta usar una integración que no está configurada."""


@dataclass
class ChannelStatus:
    channel: str
    mode: str            # "manual" | "api"
    configured: bool
    note: str


class ChannelIntegration:
    """Interfaz base. Ninguna implementación real de envío/lectura vive aquí."""
    name = "base"
    supports_api = False

    def status(self) -> ChannelStatus:
        return ChannelStatus(self.name, "manual", False,
                             "Registro manual. Integración por API no configurada.")

    def send(self, *args, **kwargs):
        raise IntegrationNotConfigured(
            f"{self.name}: el envío automático no está disponible. Usa el registro manual.")

    def fetch(self, *args, **kwargs):
        raise IntegrationNotConfigured(
            f"{self.name}: la lectura automática no está disponible.")


class LinkedInIntegration(ChannelIntegration):
    name = "LinkedIn"
    supports_api = False  # por política: solo manual / CSV

    def status(self):
        return ChannelStatus(
            self.name, "manual", False,
            "Solo registro manual o importación CSV. Sin scraping ni lectura de "
            "mensajes privados (política de LinkedIn).")


class GmailIntegration(ChannelIntegration):
    name = "Gmail"
    supports_api = True

    def is_configured(self) -> bool:
        # 1) st.secrets (Streamlit) — 2) variables de entorno. Nunca en el código.
        try:
            import streamlit as st
            g = st.secrets.get("google", {})
            if g.get("client_id") and g.get("client_secret"):
                return True
        except Exception:
            pass
        return bool(os.environ.get("GMAIL_OAUTH_CLIENT_ID")
                    and os.environ.get("GMAIL_OAUTH_CLIENT_SECRET"))

    def status(self):
        cfg = self.is_configured()
        return ChannelStatus(
            self.name, "api" if cfg else "manual", cfg,
            "Gmail API (OAuth 2.0) " + ("configurada vía st.secrets: usa la pestaña "
            "📧 Gmail en Notifications." if cfg else
            "no configurada. Agrega [google] a st.secrets. Hoy: registro manual."))


class WhatsAppIntegration(ChannelIntegration):
    name = "WhatsApp"
    supports_api = True

    def is_configured(self) -> bool:
        return bool(os.environ.get("WHATSAPP_BUSINESS_TOKEN")
                    and os.environ.get("WHATSAPP_PHONE_NUMBER_ID"))

    def status(self):
        cfg = self.is_configured()
        return ChannelStatus(
            self.name, "api" if cfg else "manual", cfg,
            "Listo para WhatsApp Business API oficial cuando se configuren credenciales "
            "por variables de entorno. Hoy: registro manual.")


class ColdCallIntegration(ChannelIntegration):
    name = "Cold Call"

    def status(self):
        return ChannelStatus(self.name, "manual", False,
                             "Registro manual del resultado de la llamada.")


REGISTRY = {
    "LinkedIn": LinkedInIntegration(),
    "Gmail": GmailIntegration(),
    "Email": GmailIntegration(),
    "WhatsApp": WhatsAppIntegration(),
    "Cold Call": ColdCallIntegration(),
}


def channel_statuses():
    """Devuelve el estado de todos los canales, para mostrar en la UI."""
    seen, out = set(), []
    for integ in REGISTRY.values():
        if integ.name in seen:
            continue
        seen.add(integ.name)
        out.append(integ.status())
    return out
