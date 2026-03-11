import pandas as pd
import os

# ================================
# PARAMETROS
# ================================

MES_OBJETIVO = 3
ANIO_OBJETIVO = 2026

archivo = "input/Reuniones.xlsx"
archivo_salida = "output/calendario_real.xlsx"

print("Procesando reuniones reales")

# ================================
# LEER HOJAS
# ================================

df_intermedia = pd.read_excel(archivo, sheet_name="ReunionesIntermedias")
df_semanal = pd.read_excel(archivo, sheet_name="ReunionesSemanales")

# ================================
# LIMPIEZA
# ================================

df_intermedia["Proyecto"] = df_intermedia["Proyecto"].str.upper()
df_semanal["Proyecto"] = df_semanal["Proyecto"].str.upper()

df_intermedia["Fecha de Fin"] = pd.to_datetime(df_intermedia["Fecha de Fin"], errors="coerce")
df_semanal["Fecha de Fin"] = pd.to_datetime(df_semanal["Fecha de Fin"], errors="coerce")

# ================================
# FILTRO TIPO
# ================================

df_intermedia = df_intermedia[df_intermedia["Tipo de Reunión"].str.upper() == "INTERMEDIA"]
df_semanal = df_semanal[df_semanal["Tipo de Reunión"].str.upper() == "SEMANAL"]

# ================================
# FILTRO MES
# ================================

df_intermedia = df_intermedia[
    (df_intermedia["Fecha de Fin"].dt.month == MES_OBJETIVO) &
    (df_intermedia["Fecha de Fin"].dt.year == ANIO_OBJETIVO)
]

df_semanal = df_semanal[
    (df_semanal["Fecha de Fin"].dt.month == MES_OBJETIVO) &
    (df_semanal["Fecha de Fin"].dt.year == ANIO_OBJETIVO)
]

# ================================
# AGRUPAR INTERMEDIA
# ================================

intermedia = (
    df_intermedia
    .sort_values("Fecha de Fin")
    .groupby("Proyecto")["Fecha de Fin"]
    .apply(lambda x: ", ".join(x.dt.strftime("%Y-%m-%d")))
    .reset_index()
)

intermedia.columns = ["Proyecto", "Fechas_Intermedia"]

# ================================
# AGRUPAR SEMANAL
# ================================

semanal = (
    df_semanal
    .sort_values("Fecha de Fin")
    .groupby("Proyecto")["Fecha de Fin"]
    .apply(lambda x: ", ".join(x.dt.strftime("%Y-%m-%d")))
    .reset_index()
)

semanal.columns = ["Proyecto", "Fechas_Semanal"]

# ================================
# UNIR
# ================================

resultado = pd.merge(intermedia, semanal, on="Proyecto", how="outer")

# ================================
# GUARDAR
# ================================

os.makedirs("output", exist_ok=True)

resultado.to_excel(archivo_salida, index=False)

print("Archivo generado:", archivo_salida)
