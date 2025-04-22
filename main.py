import pandas
import os
from tkinter import Tk, filedialog, messagebox

def seleccionar_directorio():
    root = Tk()
    root.withdraw()
    ruta = filedialog.askdirectory(title="Selecciona una carpeta con archivos Excel")
    return ruta

def seleccionar_archivo_guardado(defecto="reporte.xlsx"):
    archivo = filedialog.asksaveasfilename(title="Guardar archivo como...", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile=defecto)
    return archivo

def contarLibrosDeDirectorio(ruta):
    #Definir Tipos de Archivo admisibles
    tipos = ('.xlsx', '.xls', '.xlsm')
    #Guardar en una lista todos los archivos de este directorio que cumplan con la extension de archivo
    archivos = [f for f in os.listdir(ruta) if os.path.isfile(os.path.join(ruta, f)) and f.lower().endswith(tipos)]
    #Devolver tamaño
    return len(archivos)

def devolverDataFrameDeLosArchivosEnEsteDirectorio(ruta):
    #Definir Tipos de Archivo admisibles
    tipos = ('.xlsx', '.xls', '.xlsm')
    #Guardar en una lista todos los archivos de este directorio que cumplan con la extension de archivo
    archivos = [f for f in os.listdir(ruta) if os.path.isfile(os.path.join(ruta, f)) and f.lower().endswith(tipos)]

    listaDataframes = []

    for file in archivos:
        path_completo = os.path.join(ruta, file)

        try:
            # Leer todas las hojas del archivo
            hojas = pandas.read_excel(path_completo, sheet_name=None, usecols="A,B,C,D,E")

            for nombre_hoja, df in hojas.items():
                df['Archivo'] = file
                df['Hoja'] = nombre_hoja
                #Agregar a la lista de los dataframes
                listaDataframes.append(df)

        except Exception as e:
            print(f"Error leyendo '{file}': {e}")

    return listaDataframes


def consolidarDataframes(listaDataFrames):
    dataFramesLimpios = []
    for dataframe in listaDataFrames:
        #Unificar Nombres para las columnas
        dataframe.columns = ['Nombre', 'Codigo', 'Cantidad', 'Stock o Cliente', 'Observacion', 'Archivo', 'Hoja']
        #Quitar los registros con Cantidad Vacia
        dataframe = dataframe[dataframe['Cantidad'].notna()]
        #Agregar dataframe a la lista de dataframes limpios
        dataFramesLimpios.append(dataframe)
    #concatenar lista de dataframes limpios
    consolidado = pandas.concat(dataFramesLimpios, ignore_index=True)
    #retornar dataframe consolidado
    return consolidado
        
def agruparDatosDataFrame(dataFrame):
    #Agrupar por codigo y luego por Stock o cliente, la cantidad se totaliza
    dataFrame = dataFrame.groupby(['Codigo','Stock o Cliente'])['Cantidad'].sum().reset_index()
    #Ordenar por Stock o Cliente
    dataFrame = dataFrame.sort_values(by='Stock o Cliente', ascending=True)
    return dataFrame

def exportarExcel(dataframe, nombreNuevoArchivo):
    try:
        dataframe.to_excel(nombreNuevoArchivo, index=False)
    except Exception as e:
        print(f"Error al exportar archivo "+nombreNuevoArchivo)
        

if __name__ == "__main__":
    try:
        ruta = seleccionar_directorio()
        if not ruta:
            print("No se seleccionó ninguna carpeta. Finalizando")
            exit()
        listaDF = devolverDataFrameDeLosArchivosEnEsteDirectorio(ruta)

        consolidado = consolidarDataframes(listaDF)
        ruta_destino_consolidado = seleccionar_archivo_guardado("reporte.xlsx")
        if ruta_destino_consolidado:
            exportarExcel(consolidado, ruta_destino_consolidado)

        agrupado = agruparDatosDataFrame(consolidado)
        ruta_destino_agrupado = seleccionar_archivo_guardado("reporte.xlsx")
        if ruta_destino_agrupado:
            exportarExcel(agrupado, ruta_destino_agrupado)
    except Exception as e:
        print(f"El programa falló, finalizando : {e}")
        exit()
    
