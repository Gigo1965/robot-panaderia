"""Fecha y hora de Costa Rica, sin importar donde corra el robot.

GitHub Actions y Render corren en hora UTC (6 horas adelante de Costa Rica).
Si se usara datetime.now() directo, una venta de las 7 p.m. quedaria con
fecha del dia siguiente. Por eso todo el robot usa ahora_local().
"""
from datetime import datetime
from zoneinfo import ZoneInfo

from . import config


def zona():
    return ZoneInfo(config.ZONA_HORARIA)


def ahora_local() -> datetime:
    """Fecha y hora actual de Costa Rica, sin zona adjunta (como la guarda Excel)."""
    return datetime.now(zona()).replace(tzinfo=None, microsecond=0)


def a_hora_local(momento: datetime) -> datetime:
    """Convierte un momento que venga con zona (ej. UTC desde el SaaS) a hora de Costa Rica."""
    if momento.tzinfo is not None:
        momento = momento.astimezone(zona()).replace(tzinfo=None)
    return momento.replace(microsecond=0)
