import pandas as pd
from openpyxl import load_workbook

# 📌 Ruta del archivo Excel (Modifícala según sea necesario)
ruta_excel = 

# 📌 Diccionario con clientes y sus cuentas (modifícalo si es necesario)
clientes_cuentas = {
   "NOMBRE DEL CLIENTE": "000-000-000",
    



}

# 📌 Variables de cuentas (Edita según sea necesario)
CUENTA_IVA = "000-000-000"
CUENTA_INGRESO = "000-000-000"
CUENTA_RET_IVA= "000-000-000"
CUENTA_RET_ISR= "000-000-000"

# 📌 Cargar el archivo de Excel
wb = load_workbook(ruta_excel)
hoja_original = wb.active  # Usa la hoja activa (cámbialo si es otra)

# 📌 Convertir la hoja original a DataFrame
df = pd.DataFrame(hoja_original.values)

# 📌 Buscar los encabezados
encabezados = list(df.iloc[0])  # Obtener la primera fila como lista

# 📌 Identificar las columnas clave (El archivo de excel tiene que tener estos encabezados tal cual, para funcionar)
try:
    col_nombre_receptor = encabezados.index("Nombre Receptor")
    col_total = encabezados.index("Total")  
    col_imp_trasladado = encabezados.index("IVA 16%")
    col_sub_total = encabezados.index("SubTotal")
    col_folio = encabezados.index("Folio")  # 📌 Se agregó la columna "Folio"
    col_ret_isr = encabezados.index("Retenido ISR")
    col_ret_IVA = encabezados.index("Retenido IVA")
    col_serie = encabezados.index("Serie")
except ValueError as e:
    print(f"Error: No se encontraron los encabezados en el archivo. {e}")
    exit()

# 📌 Crear una lista con los datos reorganizados en cinco filas por cada fila original
datos_carga = []
for i, fila in df.iterrows():
    if i == 0:
        continue  # Saltar la fila de encabezados

    # Obtener el nombre del cliente y asignarle la cuenta correspondiente
    nombre_cliente = str(fila[col_nombre_receptor]).strip()
    cuenta_cliente = clientes_cuentas.get(nombre_cliente, "SIN CUENTA")  # Si no existe, asigna "SIN CUENTA"
    folio = f"PROVISION INGRESOS F-{fila[col_folio]}//{nombre_cliente}" # 📌 Capturar el valor de "Folio"

    # 📌 Fila 1: TOTAL (Cuenta Cliente)
    datos_carga.append([
        "Cuenta cliente", cuenta_cliente, "", "Folio", folio, "", "Total", fila[col_total], "", "", ""
    ])

    # 📌 Fila 2: ISR retenido
    datos_carga.append([
        "Cuenta ISR ret", CUENTA_RET_ISR, "", "Folio", folio, "", "ISR RET", fila[col_ret_isr], "", "", ""
    ])

    # 📌 Fila 3: IVA retenido
    datos_carga.append([
        "Cuenta IVA ret", CUENTA_RET_IVA, "", "Folio", folio, "", "IVA RET", fila[col_ret_IVA], "", "", ""
    ])
    # 📌 Fila 4: SUBTOTAL (Cuenta Ingreso)
    datos_carga.append([
        "Cuenta Ingreso", CUENTA_INGRESO, "", "Folio", folio, "", "SubTotal", fila[col_sub_total], "", "", ""
    ])
    
    # 📌 Fila 5: IVA 16% (Cuenta IVA)
    datos_carga.append([
        "Cuenta IVA", CUENTA_IVA, "", "Folio", folio, "", "IVA 16%", fila[col_imp_trasladado], "", "", ""
    ])

# 📌 Crear un nuevo DataFrame con los datos reorganizados
df_carga = pd.DataFrame(datos_carga)

# 📌 Agregar la nueva hoja "carga"
with pd.ExcelWriter(ruta_excel, engine="openpyxl", mode="a") as writer:
    df_carga.to_excel(writer, sheet_name="carga", index=False, header=False)


print("✅ Proceso completado.")
