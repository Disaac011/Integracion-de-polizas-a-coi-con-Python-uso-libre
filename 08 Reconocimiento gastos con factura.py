import pandas as pd
from openpyxl import load_workbook

# 📌 Ruta del archivo Excel
ruta_excel = 
# 📌 Diccionarios para cuentas de anticipo


cuentas_por_pagar = {
    "nombre de proveedor":"2110-000-000",
    

}

# Otras cuentas de la operación
CUENTA_IVA_ACREDITABLE = "1200-001-000"
CUENTA_IVA_POR_ACREDITAR = "1201-001-000"
CUENTA_BANCOS= "1120-007-000"


wb = load_workbook(ruta_excel)
hoja_original = wb.active
df = pd.DataFrame(hoja_original.values)
encabezados = list(df.iloc[0])

# Identificación de columnas
col_nombre_tercero = encabezados.index("Nombre Emisor")
col_total = encabezados.index("Total")
col_imp_trasladado = encabezados.index("IVA 16%")
col_sub_total = encabezados.index("SubTotal")
col_folio = encabezados.index("Folio")
col_poliza_Eg_ = encabezados.index("Poliza Eg")
col_fecha= encabezados.index("FECHA")

datos_carga = []
for i, fila in df.iterrows():
    if i == 0:
        continue

    nombre_tercero = str(fila[col_nombre_tercero]).strip()
    cuenta_proveedor = cuentas_por_pagar.get(nombre_tercero, "SIN CUENTA")
    banco= "INTERCAM" #CAMBIAR BANCO SEGUN CORRESPONDA
    folio = f"{str(fila[col_nombre_tercero]).strip()}//PAGO FACT--{str(fila[col_folio]).strip()}//--{banco}"
    numero_poliza = fila[col_poliza_Eg_]  # Capturar el número de póliza
    fecha = fila[col_fecha]
    iva= fila[col_imp_trasladado]
    

    # 📌 Fila 1: Encabezado de la póliza
    datos_carga.append(["Eg", numero_poliza, folio, fecha])

    # 📌 Fila 2: Cancelación de cuentas por pagar
    datos_carga.append(["", cuenta_proveedor, "0", folio, "1", fila[col_total], "", "0", "0"])

    # 📌 Fila 3: Salida de bancos
    datos_carga.append(["", CUENTA_BANCOS, "0", folio, "1", "", fila[col_total],"0", "0"])

    # 📌 Fila 4: IVA acreditable
    if pd.notna(iva) and iva != 0:
        datos_carga.append(["", CUENTA_IVA_ACREDITABLE, "0", folio, "1", iva, "", "0", "0"])

    # 📌 Fila 5: IVA pendiente de acreditar
    if pd.notna(iva) and iva != 0:
        datos_carga.append(["", CUENTA_IVA_POR_ACREDITAR, "0", folio, "1", "", iva, "0", "0"])
    # 📌 Fila 7: Fin de las partidas
    datos_carga.append(["", "FIN_PARTIDAS"])


df_carga = pd.DataFrame(datos_carga)
with pd.ExcelWriter(ruta_excel, engine="openpyxl", mode="a") as writer:
    df_carga.to_excel(writer, sheet_name="carga", index=False, header=False)

print("✅ Proceso completado.")

