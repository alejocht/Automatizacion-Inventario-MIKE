import utilidades
from tkinter import messagebox
import tkinter as tk

util = utilidades.Utilidad()


def ingresar_Carpeta_Fuente(txtObject):
    util.seleccionar_directorio()

    txtObject.config(state="normal") #HABILITAR ESCRITURA
    txtObject.delete("1.0", tk.END) #BORRAR CONTENIDO
    txtObject.insert(tk.END, util.ruta) #AGREGAR RUTA
    txtObject.config(state="disabled") #DESHABILITAR ESCRITURA

    cantidad = util.contar_Libros_Excel()
    messagebox.showinfo("Directorios cargados", f"Se encontraron {cantidad} archivos Excel")
    
    
def ingresar_Stock_X_Deposito():
    messagebox.showinfo("En desarrollo", "Todavia no lo terminé")

def guardar_Como(txtObject):
    util.destinoExportado = util.guardar_Como()

    txtObject.config(state="normal") #HABILITAR ESCRITURA
    txtObject.delete("1.0", tk.END) #BORRAR CONTENIDO
    txtObject.insert(tk.END, util.destinoExportado) #AGREGAR RUTA
    txtObject.config(state="disabled") #DESHABILITAR ESCRITURA


def procesar_Datos():
    if not util.destinoExportado:
        messagebox.showwarning("Error", "No se definió ruta de guardado")
        return
    if not util.ruta:
        messagebox.showwarning("Error","No se definió Ruta de Archivos Fuente")
        return
    util.dataframes = util.devolver_DataFrame_De_Los_Archivos_En_Este_Directorio()
    util.df_consolidado = util.consolidar_Dataframes(util.dataframes)
    agrupado = util.agrupar_datos_DataFrame(util.df_consolidado)
    util.exportar_excel(agrupado, util.destinoExportado)
    messagebox.showinfo("Reporte creado Exitosamente", f"Guardado en: {util.destinoExportado}")