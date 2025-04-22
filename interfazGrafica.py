import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def seleccionar_carpeta():
    carpeta = filedialog.askdirectory(title="Seleccioná la carpeta con archivos Excel")
    if carpeta:
        entrada_carpeta.delete(0, tk.END)
        entrada_carpeta.insert(0, carpeta)

def seleccionar_guardado():
    archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        entrada_guardado.delete(0, tk.END)
        entrada_guardado.insert(0, archivo)

def procesar_archivos():
    try:
        ruta = entrada_carpeta.get()
        archivo_salida = entrada_guardado.get()

        if not ruta or not archivo_salida:
            messagebox.showwarning("Falta información", "Seleccioná una carpeta de origen y un nombre para guardar.")
            return

        archivos = [f for f in os.listdir(ruta) if f.lower().endswith(('.xlsx', '.xls', '.xlsm'))]

        lista_df = []
        for file in archivos:
            path_completo = os.path.join(ruta, file)
            hojas = pd.read_excel(path_completo, sheet_name=None, usecols="A,B,C,D,E")
            for nombre_hoja, df in hojas.items():
                df['Archivo'] = file
                df['Hoja'] = nombre_hoja
                df.columns = ['Nombre', 'Codigo', 'Cantidad', 'Stock o Cliente', 'Observacion', 'Archivo', 'Hoja']
                df = df[df['Cantidad'].notna()]
                lista_df.append(df)

        df_consolidado = pd.concat(lista_df, ignore_index=True)
        df_consolidado.to_excel(archivo_salida, index=False)

        messagebox.showinfo("Éxito", "Archivo consolidado exportado correctamente.")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

# Crear ventana
ventana = tk.Tk()
ventana.title("Consolidador de Excel")

# Widgets
tk.Label(ventana, text="Carpeta de archivos:").grid(row=0, column=0, sticky="e")
entrada_carpeta = tk.Entry(ventana, width=50)
entrada_carpeta.grid(row=0, column=1)
tk.Button(ventana, text="Seleccionar", command=seleccionar_carpeta).grid(row=0, column=2)

tk.Label(ventana, text="Guardar como:").grid(row=1, column=0, sticky="e")
entrada_guardado = tk.Entry(ventana, width=50)
entrada_guardado.grid(row=1, column=1)
tk.Button(ventana, text="Seleccionar", command=seleccionar_guardado).grid(row=1, column=2)

tk.Button(ventana, text="Procesar Archivos", command=procesar_archivos, bg="lightgreen").grid(row=2, column=1, pady=10)

ventana.mainloop()
