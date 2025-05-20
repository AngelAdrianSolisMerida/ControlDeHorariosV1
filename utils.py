import pandas as pd
from datetime import datetime, timedelta

def validar_fecha(fecha_str, formato="%d/%m/%Y"):
    try:
        return datetime.strptime(fecha_str, formato)
    except ValueError:
        return None

def es_dia_habil(fecha):
    # 0-4 es lunes a viernes, 5-6 es fin de semana
    return fecha.weekday() < 5

def generar_rango_fechas(inicio, fin):
    delta = fin - inicio
    return [inicio + timedelta(days=i) for i in range(delta.days + 1)]