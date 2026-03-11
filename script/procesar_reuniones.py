import pandas as pd
from datetime import timedelta
import calendar
import os

# ==================================
# RUTA ARCHIVOS
# ==================================

archivo_entrada = "input/ProyectosBogota.xlsx"
archivo_salida = "output/resultado_reuniones.xlsx"

df = pd.read_excel(archivo_entrada)

print("Archivo cargado correctamente")

# ==================================
# PARAMETROS DE ANALISIS
# ==================================

anio = 2026
mes = 3

festivos = [
    "2026-03-23"
]

festivos = [pd.to_datetime(f) for f in festivos]

# ==================================
# MAPA DE DIAS
# ==================================

mapa_dias = {
    "LUNES":0,
    "MARTES":1,
    "MIERCOLES":2,
    "MIÉRCOLES":2,
    "JUEVES":3,
    "VIERNES":4,
    "SABADO":5,
    "SÁBADO":5,
    "DOMINGO":6
}

# ==================================
# CALENDARIO DEL MES
# ==================================

fechas_mes = pd.date_range(
    start=f"{anio}-{mes:02d}-01",
    end=f"{anio}-{mes:02d}-{calendar.monthrange(anio, mes)[1]}"
)

# ==================================
# FUNCIONES
# ==================================

def es_habil(fecha):

    if fecha.weekday() == 6:
        return False

    if fecha in festivos:
        return False

    return True


def siguiente_habil(fecha):

    siguiente = fecha + timedelta(days=1)

    while not es_habil(siguiente):
        siguiente += timedelta(days=1)

    return siguiente


def calcular_posibles(dia_base):

    if pd.isna(dia_base):
        return ""

    dia_base = str(dia_base).upper().strip()

    if dia_base not in mapa_dias:
        return ""

    numero_dia = mapa_dias[dia_base]

    posibles = []

    for fecha in fechas_mes:

        if fecha.weekday() == numero_dia:

            dias_extra = 1

            if fecha in festivos:
                dias_extra = 2

            actual = fecha

            for i in range(dias_extra + 1):

                if es_habil(actual):
                    posibles.append(actual)

                actual = siguiente_habil(actual)

                if actual.month != mes:
                    break

    posibles = sorted(set(posibles))

    return ", ".join([f.strftime("%Y-%m-%d") for f in posibles])

# ==================================
# CALCULO
# ==================================

df["PosibleIntermedia"] = df["DiaIntermedia"].apply(calcular_posibles)
df["PosibleSemanal"] = df["DiaSemanal"].apply(calcular_posibles)

# ==================================
# CREAR CARPETA OUTPUT
# ==================================

os.makedirs("output", exist_ok=True)

# ==================================
# GUARDAR RESULTADO
# ==================================

df.to_excel(archivo_salida, index=False)

print("Archivo generado:", archivo_salida)
