import os
import pandas
from tkinter import Tk, filedialog

class Utilidad:
    def __init__(self):
        self.ruta = ""
        self.destinoExportado = ""
        self.dataframes = []
        self.df_consolidado = None
    
    def seleccionar_directorio(self):
        root = Tk()
        root.withdraw()
        rutaArchivo = filedialog.askdirectory(title="Selecciona una carpeta con archivos Excel")
        self.ruta = rutaArchivo
   
    @staticmethod
    def guardar_Como(defecto="reporte.xlsx"):
        archivo = filedialog.asksaveasfilename(title="Guardar archivo como...", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile=defecto)
        return archivo
    
    def contar_Libros_Excel(self):
        #Definir Tipos de Archivo admisibles
        tipos = ('.xlsx', '.xls', '.xlsm')
        #Guardar en una lista todos los archivos de este directorio que cumplan con la extension de archivo
        archivos = [f for f in os.listdir(self.ruta) if os.path.isfile(os.path.join(self.ruta, f)) and f.lower().endswith(tipos)]
        #Devolver tamaño
        return len(archivos)
    
    def devolver_DataFrame_De_Los_Archivos_En_Este_Directorio(self):
        #Definir Tipos de Archivo admisibles
        tipos = ('.xlsx', '.xls', '.xlsm')
        #Guardar en una lista todos los archivos de este directorio que cumplan con la extension de archivo
        if not self.ruta:
            print("error en devolver_dataframe_de_los_archivos_en_este_directorio")
            exit()
        archivos = [f for f in os.listdir(self.ruta) if os.path.isfile(os.path.join(self.ruta, f)) and f.lower().endswith(tipos)]

        listaDataframes = []

        for file in archivos:
            path_completo = os.path.join(self.ruta, file)

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
    
    @staticmethod
    def consolidar_Dataframes(listaDataFrames):
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
    
    @staticmethod
    def agrupar_datos_DataFrame(dataFrame):
        #Agrupar por codigo y luego por Stock o cliente, la cantidad se totaliza
        dataFrame = dataFrame.groupby(['Codigo','Stock o Cliente'])['Cantidad'].sum().reset_index()
        #Ordenar por Stock o Cliente
        dataFrame = dataFrame.sort_values(by='Stock o Cliente', ascending=True)
        return dataFrame
    
    @staticmethod
    def exportar_excel(dataframe, nombreNuevoArchivo):
        try:
            dataframe.to_excel(nombreNuevoArchivo, index=False)
        except Exception as e:
            print(f"Error al exportar archivo {nombreNuevoArchivo} : {e}")
        


















        


