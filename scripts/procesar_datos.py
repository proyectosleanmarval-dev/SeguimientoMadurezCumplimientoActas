import pandas as pd
from datetime import timedelta
import calendar
import os
import math

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
MES = 3

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
    """Verifica si una fecha es hábil (no domingo y no festivo)"""
    if fecha.weekday() == 6:  # Domingo
        return False
    if fecha in festivos:
        return False
    return True

def obtener_dos_siguientes(fecha):
    """
    Obtiene los dos días siguientes a una fecha dada
    (sin importar si son hábiles o no)
    """
    siguientes = []
    for i in range(1, 3):  # i = 1, 2
        siguiente = fecha + timedelta(days=i)
        if siguiente.month == MES:  # Solo si está en el mismo mes
            siguientes.append(siguiente)
    return siguientes

def siguiente_habil(fecha):
    """Encuentra el siguiente día hábil"""
    siguiente = fecha + timedelta(days=1)
    while not es_habil(siguiente):
        siguiente += timedelta(days=1)
    return siguiente

def contar_fechas(lista_fechas_str):
    """
    Cuenta el número de fechas en una cadena separada por comas
    """
    if pd.isna(lista_fechas_str) or lista_fechas_str == "":
        return 0
    return len([x.strip() for x in str(lista_fechas_str).split(",") if x.strip()])

def calcular_indicador(posibles, coincidencias):
    """
    Calcula el indicador de cumplimiento:
    ((Cantidad de fechas posibles / 2) / Cantidad de coincidencias) * 100
    
    Ejemplo: Si hay 10 fechas posibles y 5 coincidencias:
    ((10/2) / 5) * 100 = (5/5) * 100 = 100%
    """
    count_posibles = contar_fechas(posibles)/2
    count_coincidencias = contar_fechas(coincidencias)
    
    # Evitar división por cero
    if count_posibles == 0 or count_coincidencias == 0:
        return 0.0
    
    # Calcular explícitamente: (posibles/2) / coincidencias * 100
    posibles_entre_dos = count_posibles / 2
    indicador = (posibles_entre_dos / count_coincidencias) * 100
    
    # Redondear a 2 decimales
    return round(indicador, 2)

# ==================================
# CALCULO FECHAS POSIBLES
# ==================================

def calcular_posibles(dia_base):
    """
    Calcula las fechas posibles para reuniones basado en el día de la semana.
    Si el día estipulado es festivo, considera los dos días siguientes
    (sin importar si son hábiles) y excluye el día festivo.
    """
    if pd.isna(dia_base):
        return ""
    
    dia_base = str(dia_base).upper().strip()
    
    if dia_base not in mapa_dias:
        return ""
    
    numero_dia = mapa_dias[dia_base]
    posibles = []
    
    for fecha in fechas_mes:
        if fecha.weekday() == numero_dia:
            
            # Verificar si el día estipulado es festivo
            if fecha in festivos:
                # Caso festivo: tomar los dos días siguientes (sin filtrar por hábiles)
                dias_siguientes = obtener_dos_siguientes(fecha)
                for dia_sig in dias_siguientes:
                    if dia_sig.month == MES:  # Verificar que esté en el mismo mes
                        posibles.append(dia_sig)
            else:
                # Caso normal: tomar el día y el siguiente hábil
                posibles.append(fecha)
                
                # Obtener siguiente hábil
                siguiente = siguiente_habil(fecha)
                if siguiente.month == MES:
                    posibles.append(siguiente)
    
    # Eliminar duplicados y ordenar
    posibles = sorted(set(posibles))
    
    return ", ".join([f.strftime("%Y-%m-%d") for f in posibles])

# ==================================
# CALENDARIO TEORICO
# ==================================

print("Calculando calendario teórico...")

df_proyectos["PosibleIntermedia"] = df_proyectos["DiaIntermedia"].apply(calcular_posibles)
df_proyectos["PosibleSemanal"] = df_proyectos["DiaSemanal"].apply(calcular_posibles)

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
    """Encuentra las fechas que coinciden entre dos listas"""
    if pd.isna(lista1) or pd.isna(lista2):
        return ""
    
    set1 = set([x.strip() for x in str(lista1).split(",") if x.strip()])
    set2 = set([x.strip() for x in str(lista2).split(",") if x.strip()])
    
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
# CALCULAR INDICADORES DE CUMPLIMIENTO
# ==================================

print("Calculando indicadores de cumplimiento...")

comparacion["Indicador_Intermedia"] = comparacion.apply(
    lambda row: calcular_indicador(row["PosibleIntermedia"], row["Coincidencias_Intermedia"]),
    axis=1
)

comparacion["Indicador_Semanal"] = comparacion.apply(
    lambda row: calcular_indicador(row["PosibleSemanal"], row["Coincidencias_Semanal"]),
    axis=1
)

# Agregar también los conteos para referencia (opcional)
comparacion["Count_Posibles_Intermedia"] = comparacion["PosibleIntermedia"].apply(contar_fechas)
comparacion["Count_Coincidencias_Intermedia"] = comparacion["Coincidencias_Intermedia"].apply(contar_fechas)
comparacion["Count_Posibles_Semanal"] = comparacion["PosibleSemanal"].apply(contar_fechas)
comparacion["Count_Coincidencias_Semanal"] = comparacion["Coincidencias_Semanal"].apply(contar_fechas)

# ==================================
# GUARDAR RESULTADOS
# ==================================

os.makedirs("output", exist_ok=True)

df_proyectos.to_excel(salida_teorico, index=False)
resultado_real.to_excel(salida_real, index=False)
comparacion.to_excel(salida_comparado, index=False)

print("Archivos generados correctamente")
print("Se agregaron las columnas de indicadores:")
print("- Indicador_Intermedia: ((#Posibles/2) / #Coincidencias) * 100")
print("- Indicador_Semanal: ((#Posibles/2) / #Coincidencias) * 100")
