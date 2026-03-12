import pandas as pd
import numpy as np
from datetime import timedelta
import calendar
import os
import matplotlib.pyplot as plt

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
    "2026-01-01","2026-01-12","2026-03-23","2026-04-02","2026-04-03",
    "2026-05-01","2026-05-18","2026-06-08","2026-06-15","2026-06-29",
    "2026-07-20","2026-08-07","2026-08-17","2026-10-12","2026-11-02",
    "2026-11-16","2026-12-08","2026-12-25"
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
"LUNES":0,"MARTES":1,"MIERCOLES":2,"MIÉRCOLES":2,
"JUEVES":3,"VIERNES":4,"SABADO":5,"SÁBADO":5,"DOMINGO":6
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
    for i in range(1,3):
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
# CONTEO FECHAS POSIBLES
# ==================================

def contar_fechas_y_dividir(valor):

    if pd.isna(valor) or valor == "":
        return 0

    lista = [x.strip() for x in str(valor).split(",") if x.strip() != ""]
    return len(lista)//2

df_proyectos["ConteoIntermedia"] = df_proyectos["PosibleIntermedia"].apply(contar_fechas_y_dividir)
df_proyectos["ConteoSemanal"] = df_proyectos["PosibleSemanal"].apply(contar_fechas_y_dividir)

# ==================================
# REUNIONES REALES
# ==================================

df_intermedia["Fecha de Fin"] = pd.to_datetime(df_intermedia["Fecha de Fin"],errors="coerce")
df_semanal["Fecha de Fin"] = pd.to_datetime(df_semanal["Fecha de Fin"],errors="coerce")

df_intermedia = df_intermedia[df_intermedia["Tipo de Reunión"].str.upper()=="INTERMEDIA"]
df_semanal = df_semanal[df_semanal["Tipo de Reunión"].str.upper()=="SEMANAL"]

df_intermedia = df_intermedia[
(df_intermedia["Fecha de Fin"].dt.month==MES) &
(df_intermedia["Fecha de Fin"].dt.year==ANIO)
]

df_semanal = df_semanal[
(df_semanal["Fecha de Fin"].dt.month==MES) &
(df_semanal["Fecha de Fin"].dt.year==ANIO)
]

# ==================================
# ELIMINAR DUPLICADOS
# ==================================

df_intermedia = df_intermedia.drop_duplicates(subset=["Proyecto","Fecha de Fin"])
df_semanal = df_semanal.drop_duplicates(subset=["Proyecto","Fecha de Fin"])

# ==================================
# AGRUPAR
# ==================================

intermedia = (
df_intermedia
.sort_values("Fecha de Fin")
.groupby("Proyecto")["Fecha de Fin"]
.apply(lambda x:", ".join(sorted(x.dt.strftime("%Y-%m-%d").unique())))
.reset_index()
)

intermedia.columns = ["Proyecto","Fechas_Intermedia"]

semanal = (
df_semanal
.sort_values("Fecha de Fin")
.groupby("Proyecto")["Fecha de Fin"]
.apply(lambda x:", ".join(sorted(x.dt.strftime("%Y-%m-%d").unique())))
.reset_index()
)

semanal.columns = ["Proyecto","Fechas_Semanal"]

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
# COMPARACION
# ==================================

comparacion = pd.merge(
df_proyectos,
resultado_real,
on="Proyecto",
how="outer"
)

# ==================================
# COINCIDENCIAS
# ==================================

def coincidencias(lista1,lista2):

    if pd.isna(lista1) or pd.isna(lista2):
        return ""

    set1 = set([x.strip() for x in str(lista1).split(",")])
    set2 = set([x.strip() for x in str(lista2).split(",")])

    inter = sorted(set1.intersection(set2))

    return ", ".join(inter)

comparacion["Coincidencias_Intermedia"] = comparacion.apply(
lambda row: coincidencias(row["PosibleIntermedia"],row["Fechas_Intermedia"]),axis=1)

comparacion["Coincidencias_Semanal"] = comparacion.apply(
lambda row: coincidencias(row["PosibleSemanal"],row["Fechas_Semanal"]),axis=1)

# ==================================
# CONTEOS
# ==================================

def contar_fechas(valor):

    if pd.isna(valor) or valor=="":
        return 0

    lista=[x.strip() for x in str(valor).split(",") if x.strip()!=""]
    return len(lista)

comparacion["ConteoCoincidenciasIntermedia"] = comparacion["Coincidencias_Intermedia"].apply(contar_fechas)
comparacion["ConteoCoincidenciasSemanal"] = comparacion["Coincidencias_Semanal"].apply(contar_fechas)

# ==================================
# CUMPLIMIENTO
# ==================================

comparacion["CumplimientoIntermedia"] = np.where(
comparacion["ConteoIntermedia"]==0,
0,
comparacion["ConteoCoincidenciasIntermedia"]/comparacion["ConteoIntermedia"]
)

comparacion["CumplimientoSemanal"] = np.where(
comparacion["ConteoSemanal"]==0,
0,
comparacion["ConteoCoincidenciasSemanal"]/comparacion["ConteoSemanal"]
)

# ==================================
# GUARDAR EXCEL
# ==================================

os.makedirs("output",exist_ok=True)

df_proyectos.to_excel(salida_teorico,index=False)
resultado_real.to_excel(salida_real,index=False)
comparacion.to_excel(salida_comparado,index=False)

# ==================================
# GRAFICO DASHBOARD
# ==================================

print("Generando gráfico...")

df_grafico = comparacion.copy()
df_grafico = df_grafico[df_grafico["Estado"].str.upper()=="EJECUCIÓN"]

df_grafico["CumplimientoSemanal"]*=100
df_grafico["CumplimientoIntermedia"]*=100

df_grafico["Promedio"]=(
df_grafico["CumplimientoSemanal"]+
df_grafico["CumplimientoIntermedia"]
)/2

df_grafico=df_grafico.sort_values("Promedio")

def color(v):
    if v<50:
        return "red"
    elif v<=80:
        return "gold"
    else:
        return "green"

proyectos=df_grafico["Proyecto"]

y=np.arange(len(proyectos))
h=0.35

fig,ax=plt.subplots(figsize=(12,max(6,len(proyectos)*0.6)))

b1=ax.barh(y-h/2,df_grafico["CumplimientoSemanal"],h,
color=[color(v) for v in df_grafico["CumplimientoSemanal"]],
label="Semanal")

b2=ax.barh(y+h/2,df_grafico["CumplimientoIntermedia"],h,
color=[color(v) for v in df_grafico["CumplimientoIntermedia"]],
label="Intermedia")

for bars in [b1,b2]:
    for bar in bars:
        w=bar.get_width()
        ax.text(w+1,bar.get_y()+bar.get_height()/2,f"{w:.0f}%",va="center")

ax.axvline(80,linestyle="--",label="Meta 80%")

ax.set_yticks(y)
ax.set_yticklabels(proyectos)

ax.set_xlim(0,100)
ax.set_xlabel("Cumplimiento (%)")
ax.set_title("Cumplimiento de Reuniones por Proyecto")

ax.legend()

plt.tight_layout()

plt.savefig("output/grafico_cumplimiento_proyectos.png")

plt.close()

print("Archivos generados correctamente")
