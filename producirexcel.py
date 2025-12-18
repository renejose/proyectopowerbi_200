import pandas as pd
import tkinter as tk
import os
import re
from tkinter import filedialog, messagebox
from tkinter import ttk

def obtener_mes_desde_nombre(ruta_archivo):
    nombre = os.path.basename(ruta_archivo)
    match = re.search(r'_(\d{1,2})\.xlsx$', nombre)

    if match:
        mes_num = int(match.group(1))
        return mes_num
    return None

def procesar_excel(ruta_archivo):
    try:
        nombre_salida = ''
        df_origen = pd.read_excel(ruta_archivo)
        mes_num = obtener_mes_desde_nombre(ruta_archivo)

        MESES = {
            1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
            5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
            9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
        }

        if mes_num and mes_num in MESES:
            print(MESES[mes_num])
            nombre_salida = f"modelo_datos_ventas_{MESES[mes_num]}.xlsx"
            # 👉 alternativa numérica
            # nombre_salida = f"modelo_datos_ventas_{mes_num}.xlsx"
        else:
            nombre_salida = "modelo_datos_ventas.xlsx"

        dimensiones = {}
        dimensiones_nombres = [
            "Departamento", "Sucursal", "Medicamento", "Laboratorio", "Categoría",
            "Presentación", "Requiere Receta", "Vendedor", "Método de Pago", "Cliente"
        ]

        # Crear diccionarios de dimensiones
        for col in dimensiones_nombres:
            valores_unicos = df_origen[col].dropna().unique()
            dimensiones[col] = {valor: idx + 1 for idx, valor in enumerate(valores_unicos)}

        # Crear tablas de dimensiones
        tablas_dimensiones = {}
        for col in dimensiones_nombres:
            tablas_dimensiones[col] = pd.DataFrame(
                list(dimensiones[col].items()),
                columns=[col, f"ID_{col}"]
            )

        # Crear tabla de hechos
        registros = []
        for _, row in df_origen.iterrows():
            fecha = pd.to_datetime(row["Fecha Venta"])
            registros.append([
                row["ID"], fecha, fecha.year, fecha.month, fecha.day,
                dimensiones["Departamento"].get(row["Departamento"]),
                row["Departamento"],
                dimensiones["Sucursal"].get(row["Sucursal"]),
                row["Sucursal"],
                dimensiones["Medicamento"].get(row["Medicamento"]),
                row["Medicamento"],
                dimensiones["Laboratorio"].get(row["Laboratorio"]),
                row["Laboratorio"],
                row["Cantidad Vendida"],
                row["Precio Unitario"],
                row["Total Venta"],
                dimensiones["Categoría"].get(row["Categoría"]),
                row["Categoría"],
                dimensiones["Presentación"].get(row["Presentación"]),
                row["Presentación"],
                dimensiones["Requiere Receta"].get(row["Requiere Receta"]),
                row["Requiere Receta"],
                dimensiones["Vendedor"].get(row["Vendedor"]),
                row["Vendedor"],
                row["Hora Venta"],
                dimensiones["Método de Pago"].get(row["Método de Pago"]),
                row["Método de Pago"],
                row["Descuento"],
                dimensiones["Cliente"].get(row["Cliente"]),
                row["Cliente"]
            ])

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

        # Guardar Excel
        #salida = "modelo_datos_ventas.xlsx"
        with pd.ExcelWriter(nombre_salida, engine="openpyxl") as writer:
            df_hechos.to_excel(writer, sheet_name="Tabla_Hechos", index=False)
            for col, df_dim in tablas_dimensiones.items():
                df_dim.to_excel(writer, sheet_name=f"Dim_{col}", index=False)

        messagebox.showinfo("Éxito", f"Archivo generado:\n{nombre_salida}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


def seleccionar_archivo():
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if archivo:
        lbl_archivo.config(text=archivo)
        procesar_excel(archivo)


# ---------------- INTERFAZ ----------------
root = tk.Tk()
root.title("Generador de Tabla de Hechos")
root.geometry("600x250")
root.resizable(False, False)

frame = ttk.Frame(root, padding=20)
frame.pack(expand=True)

ttk.Label(frame, text="Modelo Dimensional - Ventas de Medicamentos",
          font=("Segoe UI", 14, "bold")).pack(pady=10)

ttk.Button(frame, text="Seleccionar archivo Excel",
           command=seleccionar_archivo).pack(pady=10)

lbl_archivo = ttk.Label(frame, text="Ningún archivo seleccionado",
                        wraplength=500)
lbl_archivo.pack(pady=5)

ttk.Label(frame, text="El archivo de salida se generará automáticamente",
          foreground="gray").pack(pady=10)

root.mainloop()

