import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

class ETLGUI:
    def __init__(self, master):
        self.master = master
        master.title("Proceso ETL")
        master.geometry("600x400")

        # Widgets
        self.folder_path = tk.StringVar()
        self.folder_label = tk.Label(master, text="Carpeta de datos:")
        self.folder_entry = tk.Entry(master, textvariable=self.folder_path, width=50)
        self.folder_button = tk.Button(master, text="Seleccionar", command=self.select_folder)

        self.col_range_label = tk.Label(master, text="Rango de columnas (ej. A:D):")
        self.col_range_entry = tk.Entry(master)

        self.start_row_label = tk.Label(master, text="Fila inicial:")
        self.start_row_entry = tk.Entry(master)

        self.process_button = tk.Button(master, text="Procesar", command=self.process_data)

        self.progress_bar = ttk.Progressbar(master, orient="horizontal", length=300, mode="determinate")

        # Layout
        self.folder_label.pack(pady=5)
        self.folder_entry.pack(pady=5)
        self.folder_button.pack(pady=5)
        self.col_range_label.pack(pady=5)
        self.col_range_entry.pack(pady=5)
        self.start_row_label.pack(pady=5)
        self.start_row_entry.pack(pady=5)
        self.process_button.pack(pady=10)
        self.progress_bar.pack(pady=10)

    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        self.folder_path.set(folder_selected)

    def process_data(self):
        folder_path = self.folder_path.get()
        col_range = self.col_range_entry.get()
        start_row = int(self.start_row_entry.get()) - 1  # Ajustar para índice basado en 0

        if not folder_path or not col_range or not start_row:
            messagebox.showerror("Error", "Por favor, complete todos los campos.")
            return

        try:
            # Obtener la lista de archivos Excel en la carpeta
            excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and f.startswith('AvanceVentasINTI')]

            if not excel_files:
                messagebox.showerror("Error", "No se encontraron archivos Excel en la carpeta seleccionada.")
                return

            all_data = []
            self.progress_bar["maximum"] = len(excel_files)

            for i, file in enumerate(excel_files):
                file_path = os.path.join(folder_path, file)
                
                # Extraer año, mes y día del nombre del archivo
                date_str = file.split('.')[1:4]
                year, month, day = date_str

                # Leer el archivo Excel
                df = pd.read_excel(file_path, sheet_name='ITEM_O', usecols=col_range, skiprows=start_row)

                # Añadir columnas de año, mes y día
                df['Year'] = year
                df['Month'] = month
                df['Day'] = day

                all_data.append(df)

                # Actualizar la barra de progreso
                self.progress_bar["value"] = i + 1
                self.master.update_idletasks()

            # Consolidar todos los DataFrames
            final_df = pd.concat(all_data, ignore_index=True)

            # Exportar a Excel
            output_path = os.path.join(folder_path, 'Out.xlsx')
            final_df.to_excel(output_path, index=False)

            messagebox.showinfo("Éxito", f"Proceso completado. Archivo guardado en {output_path}")

            # Mostrar el DataFrame final
            self.show_dataframe(final_df)

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")

    def show_dataframe(self, df):
        # Crear una nueva ventana para mostrar el DataFrame
        data_window = tk.Toplevel(self.master)
        data_window.title("Dataset Final")
        data_window.geometry("800x600")

        # Crear un widget Text para mostrar el DataFrame
        text_widget = tk.Text(data_window)
        text_widget.pack(expand=True, fill='both')

        # Insertar el DataFrame en el widget Text
        text_widget.insert(tk.END, df.to_string())

if __name__ == "__main__":
    root = tk.Tk()
    etl_gui = ETLGUI(root)
    root.mainloop()
