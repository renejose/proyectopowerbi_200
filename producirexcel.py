import pandas as pd
import random
import datetime


# Función para generar fechas aleatorias en enero de 2024
def fecha_aleatoria(mes):
    inicio = datetime.date(2024, mes, 1)
    fin = datetime.date(2024, mes, 30)
    delta = (fin - inicio).days
    return inicio + datetime.timedelta(days=random.randint(0, delta))

# Listas de datos de ejemplo
def generar_datos(mes):
    departamentos = ["Lima", "Arequipa", "Cusco", "Trujillo", "Piura", "Iquitos", "Chiclayo", "Tacna", "Puno", "Ayacucho"]
    laboratorios = ["Lab Clínico A", "Lab Clínico B", "Lab Clínico C", "Lab Clínico D", "Lab Clínico E"]
    medicamentos = [
        "Paracetamol", "Ibuprofeno", "Amoxicilina", "Diclofenaco", "Aspirina", "Omeprazol", "Metformina",
        "Losartán", "Atorvastatina", "Loratadina", "Clorfenamina", "Naproxeno", "Salbutamol", "Dexametasona"
    ]
    sucursales = ["Sucursal1", "Sucursal2", "Sucursal3", "Sucursal4"]
    categorias = ["Analgésico", "Antibiótico", "Antihistamínico", "Antiácido", "Hipoglucemiante"]
    presentaciones = ["Tabletas", "Jarabe", "Inyectable", "Cápsulas", "Suspensión"]
    requiere_receta = ["Sí", "No"]
    vendedores = ["Juan Pérez", "María García", "Carlos Fernández", "Ana López"]
    metodos_pago = ["Efectivo", "Tarjeta", "Seguro"]
    clientes = ["Cliente Frecuente", "Nuevo Cliente", "Cliente VIP", "Sin Registro"]

    # Número de registros
    num_registros = 1000

    # Generar datos
    registros = []
    for i in range(num_registros):
        fecha = fecha_aleatoria(mes)
        departamento = random.choice(departamentos)
        sucursal = random.choice(sucursales)
        medicamento = random.choice(medicamentos)
        laboratorio = random.choice(laboratorios)
        cantidad = random.randint(1, 10)
        precio_unitario = round(random.uniform(5, 100), 2)
        total = round(cantidad * precio_unitario, 2)
        categoria = random.choice(categorias)
        presentacion = random.choice(presentaciones)
        receta = random.choice(requiere_receta)
        vendedor = random.choice(vendedores)
        hora_venta = f"{random.randint(8, 22)}:{random.randint(0, 59):02d}"
        metodo_pago = random.choice(metodos_pago)
        descuento = round(random.uniform(0, 10), 2) if random.random() < 0.3 else 0  # 30% chance de descuento
        cliente = random.choice(clientes)

        registros.append([
            i + 1, fecha.strftime("%Y-%m-%d"), departamento, sucursal, medicamento, laboratorio, cantidad, precio_unitario, total,
            categoria, presentacion, receta, vendedor, hora_venta, metodo_pago, descuento, cliente
        ])

    # Crear DataFrame con las nuevas columnas
    df = pd.DataFrame(registros, columns=[
        "ID", "Fecha Venta", "Departamento", "Sucursal", "Medicamento", "Laboratorio", "Cantidad Vendida", "Precio Unitario", "Total Venta",
        "Categoría", "Presentación", "Requiere Receta", "Vendedor", "Hora Venta", "Método de Pago", "Descuento", "Cliente"
    ])

    # Guardar en un archivo Excel
    nombre_archivo = "reporte_ventas_medicamentos-resultante_"+ str(mes)+ ".xlsx"
    df.to_excel(nombre_archivo, index=False)

    print(f"Reporte generado: {nombre_archivo}")


for i in range(3, 13):
    generar_datos(i)


