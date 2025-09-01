import pandas as pd
from openpyxl import load_workbook

# 📌 Ruta del archivo Excel (Modifícala según sea necesario)
ruta_excel = 

# 📌 Diccionario con proveedores y sus cuentas por pagar
cuentas_por_pagar = { 
    "NOMBRE DE PROVEEDOR":"2110-000-000",
   

    }  

# 📌 Diccionario con proveedores y sus cuentas de gasto
cuentas_gasto = {  
    "NOMBRE DE PROVEEDOR":"6100-000-000",
    

    }  

# 📌 Variables de cuentas
CUENTA_IVA = "1201-001-000"
CUENTA_ISH = "6100-037-000"


# 📌 Cargar el archivo de Excel
wb = load_workbook(ruta_excel)
hoja_original = wb.active  # Usa la hoja activa

# 📌 Convertir la hoja original a DataFrame, tomando la primera fila como encabezados
data = list(hoja_original.values)
encabezados = data[0]  # Primera fila como encabezado
df = pd.DataFrame(data[1:], columns=encabezados)  # Desde la segunda fila como datos
      

# 📌 Identificar las columnas clave
try:
    col_nombre_emisor = encabezados.index("Nombre Emisor")
    col_total = encabezados.index("Total")
    col_imp_trasladado = encabezados.index("IVA 16%")
    col_sub_total = encabezados.index("SubTotal")
    col_folio = encabezados.index("Folio")
    col_ISH = encabezados.index("ISH")
    col_num_poliza = encabezados.index("Poliza Dr")
    col_fecha = encabezados.index("Fecha")


except ValueError as e:
    print(f"❌ Error: No se encontraron los encabezados requeridos. {e}")
    exit()

# 📌 Crear la carga de datos
datos_carga = []

for i, fila in df.iterrows():
    # Obtener variables
    numero_poliza = fila[col_num_poliza]
    nombre_proveedor = str(fila[col_nombre_emisor]).strip()
    cuenta_proveedor = cuentas_por_pagar.get(nombre_proveedor, "SIN CUENTA")
    cuenta_gasto = cuentas_gasto.get(nombre_proveedor, "SIN CUENTA")
    fecha= fila[col_fecha]  # Usar fecha individual por fila

    folio = f"PROVISION DE GASTO F-{fila[col_folio]}//{nombre_proveedor}"
    iva = fila[col_imp_trasladado]
    ISH = fila[col_ISH]
    subtotal = fila[col_sub_total]
    total = fila[col_total]

    # Fila 1: Encabezado de la póliza
    datos_carga.append(["Dr", numero_poliza, folio, fecha])

    # Fila 2: GASTO
    datos_carga.append(["", cuenta_gasto, "0", folio, "1", subtotal, "", "0", "0"])

    # Fila 3: IVA por acreditar (si aplica)
    if pd.notna(iva) and iva != 0:
        datos_carga.append(["", CUENTA_IVA, "0", folio, "1", iva, "", "0", "0"])

    # Fila 4: ISH (si aplica)
    if pd.notna(ISH) and ISH != 0:
        datos_carga.append(["", CUENTA_ISH, "0", folio, "1", ISH, "", "0", "0"])

    # Fila 5: Cuenta por pagar
    datos_carga.append(["", cuenta_proveedor, "0", folio, "1", "", total, "0", "0"])

     # 📌 Fila 6: Fin de las partidas
    datos_carga.append(["", "FIN_PARTIDAS"])

# 📌 Crear un nuevo DataFrame con los datos reorganizados
df_carga = pd.DataFrame(datos_carga)

# 📌 Agregar la nueva hoja "carga"
with pd.ExcelWriter(ruta_excel, engine="openpyxl", mode="a") as writer:
    df_carga.to_excel(writer, sheet_name="carga", index=False, header=False)

print("✅ Proceso completado con éxito.")


