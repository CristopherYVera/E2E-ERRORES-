import pandas as pd

def separar_datos_excel(archivo_excel, hoja_origen, hoja_destino):
    df = pd.read_excel(archivo_excel, sheet_name=hoja_origen, header=None)

    # Separar los datos en la columna 0 utilizando la coma después del paréntesis
    df_separado = df[0].str.extract(r'([^,]+)\(([^)]+)\),([^,]+),([^,]+),([^,]+),([^,]+),([^,]+),([^,]+)')

    # Renombrar las columnas
    columnas_nuevas = {i: f'Col{i + 1}' for i in range(len(df_separado.columns))}
    df_separado.rename(columns=columnas_nuevas, inplace=True)

    # Eliminar las columnas Col4 y Col8
    df_separado.drop(columns=['Col4', 'Col8'], inplace=True)

    # Reordenar las columnas según el orden deseado
    columnas_reordenadas = ['Col1', 'Col2', 'Col3', 'Col5', 'Col7', 'Col6']
    df_separado = df_separado[columnas_reordenadas]

    # Guardar los datos en una nueva hoja
    with pd.ExcelWriter(archivo_excel, engine='openpyxl', mode='a') as writer:
        # Cambiar el nombre de la hoja si ya existe
        if hoja_destino in writer.sheets:
            hoja_destino += 'nueva'  # nombre Hoja nueva
        df_separado.to_excel(writer, sheet_name=hoja_destino, index=False, header=False)

# Llamada a la función con los parámetros deseados
archivo_excel = 'C:\\Users\\krypto\\Desktop\\data.xlsx'
hoja_origen = 'Hoja1'
hoja_destino = '19-ASO-prueba6'

separar_datos_excel(archivo_excel, hoja_origen, hoja_destino)