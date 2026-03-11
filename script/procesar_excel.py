import pandas as pd

# leer archivo
df = pd.read_excel("input/archivo.xlsx")

# ejemplo de transformación
df["fechaRegistro"] = pd.to_datetime(df["fechaRegistro"])
df["año"] = df["fechaRegistro"].dt.year
df["mes"] = df["fechaRegistro"].dt.month
df["dia"] = df["fechaRegistro"].dt.day

# guardar resultado
df.to_excel("output/resultado.xlsx", index=False)

print("Proceso terminado")
