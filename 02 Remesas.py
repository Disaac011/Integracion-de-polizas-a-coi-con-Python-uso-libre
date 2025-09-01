import pandas as pd
from openpyxl import load_workbook

# 📌 Ruta del archivo Excel (Modifícala según sea necesario)
ruta_excel = 

# 📌 Diccionario de las cuentas de Remesas// Anticipo de clientes
tercero_cuentas = {
    "NOMBRE DE CUENTA ANTICIPO" : "000-000-000",
    
}

# 📌 Variables de cuentas (Edita según sea necesario)
CUENTA_BANCO= "000-000-000"

# 📌 Cargar el archivo de Excel
wb = load_workbook(ruta_excel)
hoja_original = wb.active  # Usa la hoja activa (cámbialo si es otra)

# 📌 Convertir la hoja original a DataFrame
df = pd.DataFrame(hoja_original.values)

# 📌 Buscar los encabezados
encabezados = list(df.iloc[0])  # Obtener la primera fila como lista

# 📌 Identificar las columnas clave (El archivo de excel tiene que tener estos encabezados para funcionar
try:
    col_poliza_IT_ = encabezados.index("POLIZA IT")
    col_fecha= encabezados.index("FECHA")
    col_terceros= encabezados.index("TERCERO")
    col_concepto= encabezados.index("CONCEPTO")
    col_monto= encabezados.index("INGRESOS")
    col_referencia= encabezados.index("REFERENCIA")
    
except ValueError as e:
    print(f"Error: No se encontraron los encabezados en el archivo. {e}")
    exit()

# 📌 Crear una lista con los datos reorganizados en cuatro filas por cada fila original
datos_carga = []
for i, fila in df.iterrows():
    if i == 0:
        continue  # Saltar la fila de encabezados

    # Obtener el nombre del tercero y asignarle la cuenta correspondiente
    nombre_tercero = str(fila[col_terceros]).strip()
    cuenta_tercero = tercero_cuentas.get(nombre_tercero, "SIN CUENTA")  # Si no existe, asigna "SIN CUENTA"
    concepto = f"REMESA-{nombre_tercero}//{fila[col_concepto]}--REF {fila[col_referencia]}--NOMBRE DEL BANCO" # 📌 Capturar el valor de "Folio" (Se puede cambiar para tener el concepto que se requiera)
    numero_poliza = fila[col_poliza_IT_] #Capturar el numero de fecha
    fecha = fila[col_fecha]

    # 📌 Fila 1: Títulos de la poliza

    datos_carga.append([
        "IT", numero_poliza, concepto,fecha
    ])

    # 📌 Fila 2: Entrada al banco
  
    datos_carga.append([
        "", CUENTA_BANCO, "0",concepto, "1",fila[col_monto], "", "0", "0"
    ])

    # 📌 Fila 3: Remesa

    datos_carga.append([
        "", cuenta_tercero, "0",concepto, "1", "", fila[col_monto], "0", "0"
    ])
    # 📌 Fila 4: Fin de la poliza

    datos_carga.append([
        "","FIN_PARTIDAS"
    ])
    

# 📌 Crear un nuevo DataFrame con los datos reorganizados
df_carga = pd.DataFrame(datos_carga)

# 📌 Agregar la nueva hoja "carga"
with pd.ExcelWriter(ruta_excel, engine="openpyxl", mode="a") as writer:
    df_carga.to_excel(writer, sheet_name="carga", index=False, header=False)


print("✅ Proceso completado.")
