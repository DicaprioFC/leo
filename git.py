import os
import pandas as pd
import streamlit as st
import re
import matplotlib.pyplot as plt
from io import BytesIO

def col2num(col_str):
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    return col_num - 1

def extraer_fecha_archivo(nombre_archivo):
    patron = r'AvanceVentasINTI\.(\d{4})\.(\d{2})\.(\d{2})'
    coincidencia = re.search(patron, nombre_archivo)
    if coincidencia:
        return coincidencia.group(1), coincidencia.group(2), coincidencia.group(3)
    return None

def cargar_archivos(archivos_subidos, columnas, fila_inicial):
    dataframes = []
    col_inicio, col_fin = [col2num(col) for col in columnas.split(':')]
    usecols = list(range(col_inicio, col_fin + 1))

    for archivo in archivos_subidos:
        try:
            xls = pd.ExcelFile(archivo)
            if 'ITEM_O' in xls.sheet_names:
                df = pd.read_excel(archivo, sheet_name='ITEM_O', usecols=usecols, skiprows=fila_inicial-1)
                fecha_info = extraer_fecha_archivo(archivo.name)
                if fecha_info:
                    año, mes, dia = fecha_info
                    df['Año'] = año
                    df['Mes'] = mes
                    df['Día'] = dia
                dataframes.append(df)
            else:
                st.warning(f"La hoja 'ITEM_O' no se encuentra en el archivo {archivo.name}.")
        except Exception as e:
            st.error(f"Error al leer el archivo {archivo.name}: {e}")

    if dataframes:
        df_final = pd.concat(dataframes, ignore_index=True)
        expected_columns = ['OFICINA', 'CODIGO', 'NOMBRE', 'LINEA', 'GRUPO', 'PNG', 'U', 'VALOR', 'U2', 'VALOR2', 'LV1', 'VALORC', 'LV2', 'Año', 'Mes', 'Día']
        df_final.columns = expected_columns[:len(df_final.columns)]
        return df_final
    else:
        return pd.DataFrame()

def generar_graficos(df, output_folder):
    fig1, ax1 = plt.subplots()
    df['LINEA'].value_counts().plot(kind='bar', ax=ax1)
    ax1.set_title('Distribución por LÍNEA')
    ax1.set_xlabel('LÍNEA')
    ax1.set_ylabel('Frecuencia')
    bar_chart_path = os.path.join(output_folder, 'Distribucion_por_LINEA.png')
    fig1.savefig(bar_chart_path)

    fig2, ax2 = plt.subplots()
    df['GRUPO'].value_counts().plot(kind='pie', ax=ax2, autopct='%1.1f%%')
    ax2.set_title('Distribución por GRUPO')
    ax2.set_ylabel('')
    pie_chart_path = os.path.join(output_folder, 'Distribucion_por_GRUPO.png')
    fig2.savefig(pie_chart_path)

    plt.close(fig1)
    plt.close(fig2)

    return bar_chart_path, pie_chart_path

def main():
    st.title("Proceso ETL - Archivos Excel")

    archivos_subidos = st.file_uploader("Suba los archivos Excel", accept_multiple_files=True, type=['xlsx'])
    columnas = st.text_input("Ingrese el rango de columnas (ej. A:C):")
    fila_inicial = st.number_input("Ingrese el número de fila inicial:", min_value=1, value=1)

    if st.button("Procesar Datos"):
        if archivos_subidos and columnas and fila_inicial:
            df_final = cargar_archivos(archivos_subidos, columnas, fila_inicial)
            if not df_final.empty:
                output_folder = "."
                output_path = os.path.join(output_folder, 'Out.xlsx')
                df_final.to_excel(output_path, index=False)
                bar_chart_path, pie_chart_path = generar_graficos(df_final, output_folder)
                st.success(f"Datos exportados a {output_path}")
                st.image(bar_chart_path, caption='Distribución por LÍNEA')
                st.image(pie_chart_path, caption='Distribución por GRUPO')
            else:
                st.info("No se encontraron datos para procesar.")
        else:
            st.error("Por favor, complete todos los campos.")

if __name__ == "__main__":
    main()
