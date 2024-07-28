import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import re
import matplotlib.pyplot as plt

def col2num(col_str):
    """ Convierte un nombre de columna de Excel a un número. """
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    return col_num - 1

def num2col(col_num):
    """ Convierte un número a un nombre de columna de Excel. """
    col_str = ''
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        col_str = chr(65 + remainder) + col_str
    return col_str

def extraer_fecha_archivo(nombre_archivo):
    """
    Extrae el año, mes y día del nombre del archivo.

    :param nombre_archivo: Nombre del archivo Excel.
    :return: Tupla con (año, mes, día) o None si no se puede extraer.
    """
    patron = r'AvanceVentasINTI\.(\d{4})\.(\d{2})\.(\d{2})'
    coincidencia = re.search(patron, nombre_archivo)
    if coincidencia:
        return coincidencia.group(1), coincidencia.group(2), coincidencia.group(3)
    return None

def cargar_archivos(ruta, columnas, fila_inicial, barra_progreso):
    """
    Carga los archivos Excel de la carpeta especificada y los consolida en un DataFrame.

    :param ruta: Ruta de la carpeta que contiene los archivos Excel.
    :param columnas: Rango de columnas a extraer (ej. 'A:O').
    :param fila_inicial: Número de fila desde donde se empezará a leer los datos.
    :param barra_progreso: Widget de barra de progreso para actualizar.
    :return: DataFrame consolidado con los datos de todos los archivos Excel.
    """
    dataframes = []

    # Convertir el rango de columnas a índices numéricos
    col_inicio, col_fin = [col2num(col) for col in columnas.split(':')]
    usecols = list(range(col_inicio, col_fin + 1))

    archivos_excel = [archivo for archivo in os.listdir(ruta) if archivo.endswith('.xlsx')]
    total_archivos = len(archivos_excel)

    max_columnas = 0

    for i, archivo in enumerate(archivos_excel, 1):
        ruta_completa = os.path.join(ruta, archivo)
        try:
            # Leer solo la hoja 'ITEM_O' con las columnas especificadas
            df = pd.read_excel(ruta_completa, sheet_name='ITEM_O', usecols=usecols, skiprows=fila_inicial-1)
            
            # Actualizar el número máximo de columnas
            max_columnas = max(max_columnas, len(df.columns))

            # Extraer año, mes y día del nombre del archivo
            fecha_info = extraer_fecha_archivo(archivo)
            if fecha_info:
                año, mes, dia = fecha_info
                df['Año'] = año
                df['Mes'] = mes
                df['Día'] = dia
            
            dataframes.append(df)

            # Actualizar la barra de progreso
            progreso = int((i / total_archivos) * 100)
            barra_progreso['value'] = progreso
            ventana.update_idletasks()

        except Exception as e:
            print(f"Error al leer el archivo {archivo}: {e}")

    # Asegurar que todos los DataFrames tengan el mismo número de columnas
    for i, df in enumerate(dataframes):
        if len(df.columns) < max_columnas:
            for j in range(len(df.columns), max_columnas):
                df[f'Col_{j}'] = pd.np.nan
        # Renombrar las columnas para asegurar consistencia
        df.columns = [f'Col_{i}' for i in range(max_columnas)] + ['Año', 'Mes', 'Día']

    # Consolidar todos los DataFrames en uno solo
    if dataframes:
        df_final = pd.concat(dataframes, ignore_index=True)
        # Renombrar las columnas con los nombres correctos
        df_final.columns = ['OFICINA', 'CODIGO', 'NOMBRE', 'LINEA', 'GRUPO', 'PNG', 'U', 'VALOR', 'U2', 'VALOR2', 'LV1', 'VALORC', 'LV2', 'Año', 'Mes', 'Día']
        return df_final
    else:
        return pd.DataFrame()

def seleccionar_carpeta():
    """
    Abre un diálogo para que el usuario seleccione la carpeta con los archivos Excel.

    :return: Ruta de la carpeta seleccionada.
    """
    ruta = filedialog.askdirectory(title="Seleccione la carpeta con los archivos Excel")
    if ruta:
        etiqueta_ruta.config(text=f"Carpeta seleccionada: {ruta}")
        return ruta
    return None

def procesar_datos():
    """
    Función principal que maneja el flujo del proceso ETL.
    """
    ruta = etiqueta_ruta.cget("text").split(": ")[-1]
    if not ruta or not os.path.isdir(ruta):
        messagebox.showerror("Error", "Por favor, seleccione una carpeta válida.")
        return

    # Solicitar el rango de columnas
    rango_columnas = simpledialog.askstring("Rango de Columnas", "Ingrese el rango de columnas (ej. A:C):")
    if not rango_columnas:
        return

    # Solicitar la fila inicial
    fila_inicial = simpledialog.askinteger("Fila Inicial", "Ingrese el número de fila inicial:", minvalue=1)
    if not fila_inicial:
        return

    # Mostrar y resetear la barra de progreso
    barra_progreso.pack(pady=10)
    barra_progreso['value'] = 0
    ventana.update_idletasks()

    # Procesar los datos
    df_final = cargar_archivos(ruta, rango_columnas, fila_inicial, barra_progreso)

    if df_final.empty:
        messagebox.showinfo("Información", "No se encontraron datos para procesar.")
    else:
        # Exportar el DataFrame a Excel
        try:
            output_path = os.path.join(ruta, 'Out.xlsx')
            df_final.to_excel(output_path, index=False)
            generar_graficos(df_final, ruta)
            messagebox.showinfo("Éxito", f"Datos exportados a {output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar datos: {e}")

        # Mostrar el DataFrame en una nueva ventana
        mostrar_dataframe(df_final)

    # Ocultar la barra de progreso
    barra_progreso.pack_forget()

def generar_graficos(df, output_folder):
    """
    Genera gráficos de barras y de tortas y los guarda como archivos PNG en la misma carpeta de los datos de entrada.

    :param df: DataFrame con los datos.
    :param output_folder: Ruta de la carpeta donde se guardarán los gráficos.
    """
    # Crear gráficos
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

def mostrar_dataframe(df):
    """
    Muestra el DataFrame en una nueva ventana.

    :param df: DataFrame a mostrar.
    """
    ventana_df = tk.Toplevel(ventana)
    ventana_df.title("Dataset Final")

    # Crear un widget Text para mostrar el DataFrame
    text_widget = tk.Text(ventana_df)
    text_widget.insert(tk.END, df.to_string())
    text_widget.pack(fill=tk.BOTH, expand=True)

# Configuración de la interfaz gráfica
ventana = tk.Tk()
ventana.title("Proceso ETL - Archivos Excel")
ventana.geometry("600x250")

etiqueta_ruta = tk.Label(ventana, text="Ninguna carpeta seleccionada")
etiqueta_ruta.pack(pady=10)

boton_seleccionar = tk.Button(ventana, text="Seleccionar Carpeta", command=seleccionar_carpeta)
boton_seleccionar.pack(pady=10)

boton_procesar = tk.Button(ventana, text="Procesar Datos", command=procesar_datos)
boton_procesar.pack(pady=10)

# Barra de progreso
barra_progreso = ttk.Progressbar(ventana, orient="horizontal", length=300, mode="determinate")
barra_progreso.pack_forget()  # Inicialmente oculta

ventana.mainloop()
