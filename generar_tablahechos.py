
import pandas as pd

# Leer el archivo de origen
nombre_archivo_origen = "reporte_ventas_medicamentos-resultante_2.xlsx"
df_origen = pd.read_excel(nombre_archivo_origen)

# Crear diccionarios para almacenar IDs únicos por dimensión
dimensiones = {}

dimensiones_nombres = [
    "Departamento", "Sucursal", "Medicamento", "Laboratorio", "Categoría", "Presentación", 
    "Requiere Receta", "Vendedor", "Método de Pago", "Cliente"
]

# Asignar IDs únicos a cada valor de las dimensiones y crear DataFrames de dimensiones
for col in dimensiones_nombres:
    valores_unicos = df_origen[col].dropna().unique()
    dimensiones[col] = {valor: idx + 1 for idx, valor in enumerate(valores_unicos)}

# Crear listas para las tablas de dimensiones
tablas_dimensiones = {}
for col in dimensiones_nombres:
    tablas_dimensiones[col] = pd.DataFrame(
        list(dimensiones[col].items()), columns=[col,f"ID_{col}"]
    )

# Crear la tabla de hechos con los IDs
registros = []
for _, row in df_origen.iterrows():
    registros.append([
        row["ID"], row["Fecha Venta"], pd.to_datetime(row["Fecha Venta"]).year, pd.to_datetime(row["Fecha Venta"]).month, pd.to_datetime(row["Fecha Venta"]).day,
        dimensiones["Departamento"].get(row["Departamento"], None), row["Departamento"],
        dimensiones["Sucursal"].get(row["Sucursal"], None), row["Sucursal"],
        dimensiones["Medicamento"].get(row["Medicamento"], None), row["Medicamento"],
        dimensiones["Laboratorio"].get(row["Laboratorio"], None), row["Laboratorio"],
        row["Cantidad Vendida"], row["Precio Unitario"], row["Total Venta"],
        dimensiones["Categoría"].get(row["Categoría"], None), row["Categoría"],
        dimensiones["Presentación"].get(row["Presentación"], None), row["Presentación"],
        dimensiones["Requiere Receta"].get(row["Requiere Receta"], None), row["Requiere Receta"],
        dimensiones["Vendedor"].get(row["Vendedor"], None), row["Vendedor"],
        row["Hora Venta"],
        dimensiones["Método de Pago"].get(row["Método de Pago"], None), row["Método de Pago"],
        row["Descuento"],
        dimensiones["Cliente"].get(row["Cliente"], None), row["Cliente"]
    ])

# Crear DataFrame con la tabla de hechos
df_hechos = pd.DataFrame(registros, columns=[
    "ID_Venta", "Fecha_Venta", "Año", "Mes", "Día", 
    "ID_Departamento", "Departamento",
    "ID_Sucursal", "Sucursal",
    "ID_Medicamento", "Medicamento",
    "ID_Laboratorio", "Laboratorio",
    "Cantidad_Vendida", "Precio_Unitario", "Total_Venta",
    "ID_Categoría", "Categoría",
    "ID_Presentación", "Presentación",
    "ID_Requiere_Receta", "Requiere_Receta",
    "ID_Vendedor", "Vendedor",
    "Hora_Venta",
    "ID_Método_Pago", "Método_Pago",
    "Descuento",
    "ID_Cliente", "Cliente"
])

# Guardar en un archivo Excel con varias hojas
nombre_archivo_destino = "modelo_datos_ventas_2.xlsx"
with pd.ExcelWriter(nombre_archivo_destino, engine="openpyxl") as writer:
    df_hechos.to_excel(writer, sheet_name="Tabla_Hechos", index=False)
    for col, df_dim in tablas_dimensiones.items():
        df_dim.to_excel(writer, sheet_name=f"Dim_{col}", index=False)

print(f"Modelo de datos generado en: {nombre_archivo_destino}")
