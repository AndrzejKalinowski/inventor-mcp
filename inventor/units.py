"""Birim donusumleri.

Inventor COM API'si uzunluklari her zaman santimetre, acilari her zaman
radyan cinsinden alir/dondurur - dokumanin goruntu birimi ne olursa olsun.
Bu MCP araclari kullaniciya/Claude'a karsi milimetre ve derece kullanir.
"""

import math


def mm_to_cm(mm: float) -> float:
    return mm / 10.0


def cm_to_mm(cm: float) -> float:
    return cm * 10.0


def deg_to_rad(deg: float) -> float:
    return math.radians(deg)


def rad_to_deg(rad: float) -> float:
    return math.degrees(rad)


def cm2_to_mm2(cm2: float) -> float:
    return cm2 * 100.0


def cm3_to_mm3(cm3: float) -> float:
    return cm3 * 1000.0
