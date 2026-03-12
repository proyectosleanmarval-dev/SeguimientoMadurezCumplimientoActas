import pandas as pd
import numpy as np
from datetime import timedelta
import calendar
import os

# ==================================
# ARCHIVO DE ENTRADA
# ==================================

archivo = "input/Reuniones.xlsx"

salida_teorico = "output/calendario_teorico.xlsx"
salida_real = "output/calendario_real.xlsx"
salida_comparado = "output/calendario_comparado.xlsx"

# ==================================
# PARAMETROS
# ==================================

ANIO = 2026
MES = 2

festivos = [
    "2026-03-11"
]

festivos = [pd.to_datetime(f) for f in festivos]

# ==================================
# LEER HOJAS
# ==================================

print("Leyendo archivo...")

df_proyectos = pd.read_excel(archivo, sheet_name="ProyectosBogota")
df_intermedia = pd.read_excel(archivo, sheet_name="ReunionesIntermedias")
df_semanal = pd.read_excel(archivo, sheet_name="ReunionesSemanales")

# ==================================
# NORMALIZAR PROYECTO
# ==================================

df_proyectos["Proyecto"] = df_proyectos["Proyecto"].str.upper()
df_intermedia["Proyecto"] = df_intermedia["Proyecto"].str.upper()
df_semanal["Proyecto"] = df_semanal["Proyecto"].str.upper()

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

def obtener_dos_siguientes(fecha):
    siguientes = []
    for i in range(1, 3):
        siguiente = fecha + timedelta(days=i)
        if siguiente.month == MES:
            siguientes.append(siguiente)
    return siguientes

def siguiente_habil(fecha):
    siguiente = fecha + timedelta(days=1)
    while not es_habil(siguiente):
        siguiente += timedelta(days=1)
    return siguiente

# ==================================
# CALCULO FECHAS POSIBLES
# ==================================

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
            
            if fecha in festivos:
                
                dias_siguientes = obtener_dos_siguientes(fecha)
                
                for dia_sig in dias_siguientes:
                    if dia_sig.month == MES:
                        posibles.append(dia_sig)
            
            else:
                
                posibles.append(fecha)
                
                siguiente = siguiente_habil(fecha)
                
                if siguiente.month == MES:
                    posibles.append(siguiente)
    
    posibles = sorted(set(posibles))
    
    return ", ".join([f.strftime("%Y-%m-%d") for f in posibles])

# ==================================
# CALENDARIO TEORICO
# ==================================

print("Calculando calendario teórico...")

df_proyectos["PosibleIntermedia"] = df_proyectos["DiaIntermedia"].apply(calcular_posibles)
df_proyectos["PosibleSemanal"] = df_proyectos["DiaSemanal"].apply(calcular_posibles)

# ==================================
# CONTEO DE FECHAS POSIBLES
# ==================================

def contar_fechas_y_dividir(valor):

    if pd.isna(valor) or valor == "":
        return 0
    
    lista = [x.strip() for x in str(valor).split(",") if x.strip() != ""]
    
    return len(lista) // 2


df_proyectos["ConteoIntermedia"] = df_proyectos["PosibleIntermedia"].apply(contar_fechas_y_dividir)
df_proyectos["ConteoSemanal"] = df_proyectos["PosibleSemanal"].apply(contar_fechas_y_dividir)

# ==================================
# REUNIONES REALES
# ==================================

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
# AGRUPAR INTERMEDIAS
# ==================================

intermedia = (
    df_intermedia
    .sort_values("Fecha de Fin")
    .groupby("Proyecto")["Fecha de Fin"]
    .apply(lambda x: ", ".join(x.dt.strftime("%Y-%m-%d")))
    .reset_index()
)

intermedia.columns = ["Proyecto", "Fechas_Intermedia"]

# ==================================
# AGRUPAR SEMANALES
# ==================================

semanal = (
    df_semanal
    .sort_values("Fecha de Fin")
    .groupby("Proyecto")["Fecha de Fin"]
    .apply(lambda x: ", ".join(x.dt.strftime("%Y-%m-%d")))
    .reset_index()
)

semanal.columns = ["Proyecto", "Fechas_Semanal"]

# ==================================
# RESULTADO REAL
# ==================================

resultado_real = pd.merge(
    intermedia,
    semanal,
    on="Proyecto",
    how="outer"
)

# ==================================
# FULL JOIN TEORICO VS REAL
# ==================================

comparacion = pd.merge(
    df_proyectos,
    resultado_real,
    on="Proyecto",
    how="outer"
)

# ==================================
# FUNCION COINCIDENCIAS
# ==================================

def coincidencias(lista1, lista2):

    if pd.isna(lista1) or pd.isna(lista2):
        return ""
    
    set1 = set([x.strip() for x in str(lista1).split(",")])
    set2 = set([x.strip() for x in str(lista2).split(",")])
    
    inter = sorted(set1.intersection(set2))
    
    return ", ".join(inter)

# ==================================
# CALCULAR COINCIDENCIAS
# ==================================

comparacion["Coincidencias_Intermedia"] = comparacion.apply(
    lambda row: coincidencias(row["PosibleIntermedia"], row["Fechas_Intermedia"]),
    axis=1
)

comparacion["Coincidencias_Semanal"] = comparacion.apply(
    lambda row: coincidencias(row["PosibleSemanal"], row["Fechas_Semanal"]),
    axis=1
)

# ==================================
# CONTEO DE COINCIDENCIAS
# ==================================

def contar_fechas(valor):

    if pd.isna(valor) or valor == "":
        return 0
    
    lista = [x.strip() for x in str(valor).split(",") if x.strip() != ""]
    
    return len(lista)


comparacion["ConteoCoincidenciasIntermedia"] = comparacion["Coincidencias_Intermedia"].apply(contar_fechas)
comparacion["ConteoCoincidenciasSemanal"] = comparacion["Coincidencias_Semanal"].apply(contar_fechas)

# ==================================
# CALCULO DE CUMPLIMIENTO
# ==================================

comparacion["CumplimientoIntermedia"] = np.where(
    comparacion["ConteoIntermedia"] == 0,
    0,
    comparacion["ConteoCoincidenciasIntermedia"] / comparacion["ConteoIntermedia"]
)

comparacion["CumplimientoSemanal"] = np.where(
    comparacion["ConteoSemanal"] == 0,
    0,
    comparacion["ConteoCoincidenciasSemanal"] / comparacion["ConteoSemanal"]
)

# ==================================
# GUARDAR RESULTADOS
# ==================================

os.makedirs("output", exist_ok=True)

df_proyectos.to_excel(salida_teorico, index=False)
resultado_real.to_excel(salida_real, index=False)
comparacion.to_excel(salida_comparado, index=False)

print("Archivos generados correctamente")
