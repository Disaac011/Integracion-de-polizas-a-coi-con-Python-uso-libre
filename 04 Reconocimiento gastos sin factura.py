import pandas as pd
from openpyxl import load_workbook

# 📌 Ruta del archivo Excel
ruta_excel = 
# 📌 Diccionarios para cuentas de anticipo


anticipo_clientes = {
    "CUENTA DE ANTICIPO DE CLIENTE" : "000-000-000",
    
    }

# Otras cuentas de la operación
CUENTA_BANCOS= "000-000-000"


wb = load_workbook(ruta_excel)
hoja_original = wb.active
df = pd.DataFrame(hoja_original.values)
encabezados = list(df.iloc[0])

# Identificación de columnas
col_nombre_tercero = encabezados.index("TERCERO")
col_total = encabezados.index("TOTAL")
col_folio = encabezados.index("CONCEPTO")
col_poliza_ET = encabezados.index("POLIZA DE GASTO")
col_fecha= encabezados.index("FECHA")
col_referencia= encabezados.index("REFERENCIA")
col_promotor= encabezados.index("PROMOTOR")

datos_carga = []
for i, fila in df.iterrows():
    if i == 0:
        continue

    nombre_tercero = str(fila[col_nombre_tercero]).strip()
    cuenta_anticipo = anticipo_clientes.get(nombre_tercero, "SIN CUENTA")
    banco= "SANTANDER" #CAMBIAR BANCO SEGUN CORRESPONDA
    folio = f"{str(fila[col_promotor]).strip()}//{str(fila[col_folio]).strip()}//CLIENTE {str(fila[col_nombre_tercero]).strip()}//{str(fila[col_referencia]).strip()}//{banco}"
    numero_poliza = fila[col_poliza_ET]  # Capturar el número de póliza
    fecha = fila[col_fecha]
    

    # 📌 Fila 1: Encabezado de la póliza
    datos_carga.append(["ET", numero_poliza, folio, fecha])

    # 📌 Fila 2: Cancelación del anticipo de cliente
    datos_carga.append(["", cuenta_anticipo, "0", folio, "1", fila[col_total], "", "0", "0"])

    # 📌 Fila 3: Salida de bancos
    datos_carga.append(["", CUENTA_BANCOS, "0", folio, "1", "", fila[col_total],"0", "0"])

    # 📌 Fila 4: Fin de las partidas
    datos_carga.append(["", "FIN_PARTIDAS"])


df_carga = pd.DataFrame(datos_carga)
with pd.ExcelWriter(ruta_excel, engine="openpyxl", mode="a") as writer:
    df_carga.to_excel(writer, sheet_name="carga", index=False, header=False)

print("✅ Proceso completado.")


