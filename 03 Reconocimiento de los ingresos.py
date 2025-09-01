import pandas as pd
from openpyxl import load_workbook

# 📌 Ruta del archivo Excel
ruta_excel = 
# 📌 Diccionarios para cuentas de cliente y cuentas de anticipo
clientes_cuentas = {
   "CUENTA DE CLIENTE": "000-000-000",
  

}
anticipo_clientes = {
    "CUENTA DE ANTICIPO DE CLIENTE" : "000-000-000",
    

}

# Otras cuentas de la operación
CUENTA_IVA_POR_COBRAR = "000-000-000"
CUENTA_IVA_COBRADO = "000-000-000"
CUENTA_RET_ISR_PEND = "000-000-000"
CUENTA_RET_ISR = "000-000-000"
CUENTA_RET_IVA_PEND = "000-000-000"
CUENTA_RET_IVA = "000-000-000"
CUENTA_ORDEN = "000-000-000"
CONTRACUENTA_ORDEN = "000-000-000"

wb = load_workbook(ruta_excel)
hoja_original = wb.active
df = pd.DataFrame(hoja_original.values)
encabezados = list(df.iloc[0])

# Identificación de columnas (El archivo de excel tiene que tener estos encabezados para funcionar
col_nombre_receptor = encabezados.index("Nombre Receptor")
col_total = encabezados.index("Total")
col_imp_trasladado = encabezados.index("IVA 16%")
col_sub_total = encabezados.index("SubTotal")
col_folio = encabezados.index("Folio")
col_ret_isr = encabezados.index("Retenido ISR")
col_ret_IVA = encabezados.index("Retenido IVA")
col_poliza_Ig_ = encabezados.index("POLIZA DE INGRESO")
col_fecha= encabezados.index("FECHA")

datos_carga = []
for i, fila in df.iterrows():
    if i == 0:
        continue

    nombre_cliente = str(fila[col_nombre_receptor]).strip()
    cuenta_cliente = clientes_cuentas.get(nombre_cliente, "SIN CUENTA")
    cuenta_anticipo = anticipo_clientes.get(nombre_cliente, "SIN CUENTA")
    folio = f"COBRO F-{fila[col_folio]}//{nombre_cliente}"
    numero_poliza = fila[col_poliza_Ig_]  # Capturar el número de póliza
    fecha = fila[col_fecha]
    ingreso = (fila[col_imp_trasladado]) / 0.16

    # Obtener valores de retenciones
    valor_ret_iva = fila[col_ret_IVA]
    valor_ret_isr = fila[col_ret_isr]

    # 📌 Fila 1: Encabezado de la póliza
    datos_carga.append(["Ig", numero_poliza, folio, fecha])

    # 📌 Fila 2: Cancelación del anticipo de cliente
    datos_carga.append(["", cuenta_anticipo, "0", folio, "1", fila[col_total], "", "0", "0"])

    # 📌 Fila 3: Cancelación del IVA por cobrar
    datos_carga.append(["", CUENTA_IVA_POR_COBRAR, "0", folio, "1", fila[col_imp_trasladado], "", "0", "0"])

    # 📌 Fila 4: IVA retenido (si aplica)
    if pd.notna(valor_ret_iva) and valor_ret_iva != 0:
        datos_carga.append(["", CUENTA_RET_IVA, "0", folio, "1", valor_ret_iva, "", "0", "0"])

    # 📌 Fila 5: ISR retenido (si aplica)
    if pd.notna(valor_ret_isr) and valor_ret_isr != 0:
        datos_carga.append(["", CUENTA_RET_ISR, "0", folio, "1", valor_ret_isr, "", "0", "0"])

    # 📌 Fila 6: Cancelación de cuenta cliente
    datos_carga.append(["", cuenta_cliente, "0", folio, "1", "", fila[col_total], "0", "0"])

    # 📌 Fila 7: IVA trasladado
    datos_carga.append(["", CUENTA_IVA_COBRADO, "0", folio, "1", "", fila[col_imp_trasladado], "0", "0"])

    # 📌 Fila 8: Cancelación del IVA por retener (si aplica)
    if pd.notna(valor_ret_iva) and valor_ret_iva != 0:
        datos_carga.append(["", CUENTA_RET_IVA_PEND, "0", folio, "1", "", valor_ret_iva, "0", "0"])

    # 📌 Fila 9: Cancelación del ISR por retener (si aplica)
    if pd.notna(valor_ret_isr) and valor_ret_isr != 0:
        datos_carga.append(["", CUENTA_RET_ISR_PEND, "0", folio, "1", "", valor_ret_isr, "0", "0"])

    # 📌 Fila 10: Cuenta de orden - ingreso cobrado
    datos_carga.append(["", CUENTA_ORDEN, "0", folio, "1", ingreso, "", "0", "0"])

    # 📌 Fila 11: Contracuenta de orden
    datos_carga.append(["", CONTRACUENTA_ORDEN, "0", folio, "1", "", ingreso, "0", "0"])

    # 📌 Fila 12: Fin de las partidas
    datos_carga.append(["", "FIN_PARTIDAS"])


    
    
    
    
    
    
    

df_carga = pd.DataFrame(datos_carga)
with pd.ExcelWriter(ruta_excel, engine="openpyxl", mode="a") as writer:
    df_carga.to_excel(writer, sheet_name="carga", index=False, header=False)

print("✅ Proceso completado.")

