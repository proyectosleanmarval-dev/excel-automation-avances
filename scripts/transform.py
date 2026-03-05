import pandas as pd
from pathlib import Path

# =====================================
# Definición de rutas
# =====================================

input_path = Path("data/original/estadoObra.xlsx")
output_excel = Path("data/transformed/estadoObra_filtrado.xlsx")
output_csv = Path("data/transformed/estadoObra_filtrado.csv")

# =====================================
# Columnas requeridas
# =====================================

columnas_requeridas = [
    "descSucursal",
    "descProyecto",
    "hc",
    "Actividad",
    "tipoRestriccion",
    "nomAcuerdoServ",
    "Responsable",
    "fechaRegistro",
    "FechaCompromisoInicial",
    "FecLegalizacion",
    "FecInicioFabricacion"
]

# =====================================
# Validación de existencia del archivo
# =====================================

if not input_path.exists():
    raise FileNotFoundError(f"No se encontró el archivo: {input_path}")

# =====================================
# Lectura del archivo (solo hoja Avances)
# =====================================

try:
    df = pd.read_excel(input_path, sheet_name="Avances")
except ValueError:
    raise ValueError("La hoja 'Avances' no existe en el archivo Excel.")

# =====================================
# Normalización de nombres de columnas
# =====================================

df.columns = df.columns.astype(str).str.strip()

# =====================================
# Validar columna de filtro
# =====================================

if "descSucursal" not in df.columns:
    raise ValueError(
        f"La columna 'descSucursal' no existe. "
        f"Columnas encontradas: {list(df.columns)}"
    )

# =====================================
# Filtro por sucursal exacta "BOGOTA "
# =====================================

df_filtrado = df[
    df["descSucursal"]
    .astype(str)
    .str.upper() == "BOGOTA "
]

# =====================================
# Validar que todas las columnas requeridas existan
# =====================================

columnas_faltantes = [
    col for col in columnas_requeridas if col not in df_filtrado.columns
]

if columnas_faltantes:
    raise ValueError(
        f"Faltan las siguientes columnas requeridas: {columnas_faltantes}"
    )

# =====================================
# Seleccionar solo las columnas requeridas
# =====================================

df_final = df_filtrado[columnas_requeridas]

# =====================================
# Crear carpeta destino si no existe
# =====================================

output_excel.parent.mkdir(parents=True, exist_ok=True)

# =====================================
# Guardar resultados
# =====================================

df_final.to_excel(output_excel, index=False)
df_final.to_csv(output_csv, index=False, encoding="utf-8")

# =====================================
# Logs informativos
# =====================================

print("Transformación completada correctamente.")
print(f"Registros originales: {len(df)}")
print(f"Registros después de filtro: {len(df_filtrado)}")
print(f"Columnas finales: {list(df_final.columns)}")
