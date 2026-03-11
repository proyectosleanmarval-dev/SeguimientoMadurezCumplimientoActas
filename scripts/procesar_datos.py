import pandas as pd
from datetime import timedelta
import calendar
import os

# ==================================
# ARCHIVO
# ==================================

archivo = "input/Reuniones.xlsx"

salida_teorico = "output/calendario_teorico.xlsx"
salida_real = "output/calendario_real.xlsx"

# ==================================
# PARAMETROS
# ==================================

ANIO = 2026
MES = 3

festivos = [
    "2026-03-23"
]

festivos = [pd.to_datetime(f) for f in festivos]

# ==================================
# LEER HOJAS
# ==================================

print("Leyendo archivo")

df_proyectos = pd.read_excel(archivo, sheet_name="ProyectosBogota")
df_intermedia = pd.read_excel(archivo, sheet_name="ReunionesIntermedias")
df_semanal = pd.read_excel(archivo, sheet_name="ReunionesSemanales")

# ==================================
# MAPA DIAS
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
# CALENDARIO MES
# ==================================

fechas_mes = pd.date_range(
    start=f"{ANIO}-{MES:02d}-01",
    end=f"{ANIO}-{MES:02d}-{calendar.monthrange(ANIO, MES)[1]}"
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

                if actual.month != MES:
                    break

    posibles = sorted(set(posibles))

    return ", ".join([f.strftime("%Y-%m-%d") for f in posibles])

# ==================================
# CALENDARIO TEORICO
# ==================================

print("Calculando calendario teorico")

df_proyectos["PosibleIntermedia"] = df_proyectos["DiaIntermedia"].apply(calcular_posibles)
df_proyectos["PosibleSemanal"] = df_proyectos["DiaSemanal"].apply(calcular_posibles)

# ==================================
# LIMPIEZA REUNIONES
# ==================================

df_intermedia["Proyecto"] = df_intermedia["Proyecto"].str.upper()
df_semanal["Proyecto"] = df_semanal["Proyecto"].str.upper()

df_intermedia["Fecha de Fin"] = pd.to_datetime(df_intermedia["Fecha de Fin"], errors="coerce")
df_semanal["Fecha de Fin"] = pd.to_datetime(df_semanal["Fecha de Fin"], errors="coerce")

df_intermedia = df_intermedia[df_intermedia["Tipo de Reunión"].str.upper() == "INTERMEDIA"]
df_semanal = df_semanal[df_semanal["Tipo de Reunión"].str.upper() == "SEMANAL"]

df_intermedia = df_intermedia[
    (df_intermedia["Fecha de Fin"].dt.month == MES) &
    (df_intermedia["Fecha de Fin"].dt.year == ANIO)
]

df_semanal = df_semanal[
    (df_semanal["Fecha de Fin"].dt.month == MES) &
    (df_semanal["Fecha de Fin"].dt.year == ANIO)
]

# ==================================
# AGRUPAR REUNIONES
# ==================================

intermedia = (
    df_intermedia
    .sort_values("Fecha de Fin")
    .groupby("Proyecto")["Fecha de Fin"]
    .apply(lambda x: ", ".join(x.dt.strftime("%Y-%m-%d")))
    .reset_index()
)

intermedia.columns = ["Proyecto", "Fechas_Intermedia"]

semanal = (
    df_semanal
    .sort_values("Fecha de Fin")
    .groupby("Proyecto")["Fecha de Fin"]
    .apply(lambda x: ", ".join(x.dt.strftime("%Y-%m-%d")))
    .reset_index()
)

semanal.columns = ["Proyecto", "Fechas_Semanal"]

resultado_real = pd.merge(intermedia, semanal, on="Proyecto", how="outer")

# ==================================
# GUARDAR
# ==================================

os.makedirs("output", exist_ok=True)

df_proyectos.to_excel(salida_teorico, index=False)
resultado_real.to_excel(salida_real, index=False)

print("Archivos generados")
