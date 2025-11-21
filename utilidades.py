import os
import pandas
from tkinter import Tk, filedialog, messagebox

class Utilidad:
    def __init__(self):
        self.ruta = ""
        self.rutaStockXDepo = ""
        self.destinoExportado = ""
        self.dataframes = []
        self.df_consolidado = None
        self.df_agrupado = None
        self.df_stockXDepo = None
        self.df_comparacion = None
        
    def seleccionar_archivo(self):
        try:
            root = Tk()
            root.withdraw()
            rutaArchivo = filedialog.askopenfilename(title="Selecciona el Stock por Deposito")
            self.rutaStockXDepo = rutaArchivo
        except Exception as e:
            messagebox.showerror("Error al seleccionar Stock por Deposito", f"{e}")

    def seleccionar_directorio(self):
        try:
            root = Tk()
            root.withdraw()
            rutaArchivo = filedialog.askdirectory(title="Selecciona una carpeta con archivos Excel")
            self.ruta = rutaArchivo
        except Exception as e:
            messagebox.showerror("Error al seleccionar Datos Fuente", f"{e}")
   
    @staticmethod
    def guardar_Como(defecto="reportelocales.xlsx"):
        try:
            #Abre el explorador de windows para elegir nombre y destino del archivo
            archivo = filedialog.asksaveasfilename(title="Guardar archivo como...", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile=defecto)
            #devuelve la ruta elegida
            return archivo
        except Exception as e:
            messagebox.showerror("Error al elegir nombre y destino del archivo", f"{e}")
    
    def contar_Libros_Excel(self):
        try:
            #Definir Tipos de Archivo admisibles
            tipos = ('.xlsx', '.xls', '.xlsm')
            #Guardar en una lista todos los archivos de este directorio que cumplan con la extension de archivo
            archivos = [f for f in os.listdir(self.ruta) if os.path.isfile(os.path.join(self.ruta, f)) and f.lower().endswith(tipos)]
            #Devolver tamaño
            return len(archivos)
        except Exception as e:
            messagebox.showerror("Error al contar Libros Excel", f"{e}")
    
    def devolver_DataFrame_De_Los_Archivos_En_Este_Directorio(self):
        try:
            #Definir Tipos de Archivo admisibles
            tipos = ('.xlsx', '.xls', '.xlsm')
            #Guardar en una lista todos los archivos de este directorio que cumplan con la extension de archivo
            if not self.ruta:
                print("error en devolver_dataframe_de_los_archivos_en_este_directorio")
                exit()
            #Recorrer todos los archivos del directorio en busqueda de los .xlsx
            archivos = [f for f in os.listdir(self.ruta) if os.path.isfile(os.path.join(self.ruta, f)) and f.lower().endswith(tipos)]

            listaDataframes = []
            #Por cada Archivo del Directorio
            for file in archivos:
                path_completo = os.path.join(self.ruta, file)

                try:
                    # Leer todas las hojas del archivo
                    hojas = pandas.read_excel(path_completo, sheet_name=None, usecols="A,B,C,D,E")
                    #Por cada DataFrame(hojas) se agrega responsable del archivo y nombre de hoja
                    for nombre_hoja, df in hojas.items():
                        df['Archivo'] = file
                        df['Hoja'] = nombre_hoja
                        #Agregar a la lista de los dataframes
                        listaDataframes.append(df)
                except Exception as e:
                    messagebox.showerror("Error leyendo", f"Error leyendo '{file}': {e}")
            return listaDataframes
        except Exception as e:
            messagebox.showerror("Error en devolver DataFrames del directorio", f"{e}")
        
    @staticmethod
    def consolidar_Dataframes(listaDataFrames):
        try:
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
        except Exception as e:
            messagebox.showerror("Error en Consolidar DataFrames", f"{e}")
    
    @staticmethod
    def agrupar_datos_DataFrame(dataFrame):
        try:
            #Agrupar por codigo y luego por Stock o cliente, la cantidad se totaliza
            dataFrame = dataFrame.groupby(['Codigo','Stock o Cliente'])['Cantidad'].sum().reset_index()
            #Ordenar por Stock o Cliente
            dataFrame = dataFrame.sort_values(by='Stock o Cliente', ascending=True)
            return dataFrame
        except Exception as e:
           messagebox.showerror("Error al Agrupar Archivos", f"{e}") 
    
    @staticmethod
    def exportar_excel(dataframe, nombreNuevoArchivo):
        try:
            dataframe.to_excel(nombreNuevoArchivo, index=False)
        except Exception as e:
            messagebox.showerror("Error al exportar", f"Error al exportar archivo {nombreNuevoArchivo} : {e}")

    @staticmethod
    def generar_reporte(dataframeConsolidado, dataFrameAgrupado, dataframeComparacion, nombreNuevoArchivo):
        try:
            with pandas.ExcelWriter(nombreNuevoArchivo, engine='openpyxl') as writer:
                dataframeConsolidado.to_excel(writer, sheet_name='Consolidacion', index=False)
                dataFrameAgrupado.to_excel(writer, sheet_name='Resumen agrupado', index=False)
                dataframeComparacion.to_excel(writer, sheet_name='Comparacion', index=False)

        except Exception as e:
            messagebox.showerror("Error al Exportar Archivo", f"{e}")
    
    def leerStockXDeposito(self):
        try:
            self.df_stockXDepo = pandas.read_excel(self.rutaStockXDepo, sheet_name='hoja1', usecols="A,B,C,D,E,F,G" )
            self.df_stockXDepo.columns = ['Producto', 'Codigo', 'Deposito', 'Cantidad', 'Unidad','Familia','Activo']
        except Exception as e:
            messagebox.showerror("Error al leer Stock por Deposito", f"{e}" )
    
    def comparar_inventarios(self):
        try:
            REEMPLAZOS = {'CLIENTE':'ENTREGA', 'CLIENTE ':'ENTREGA', 'STOCK':'ENTREGA INM.', 'STOCK ':'ENTREGA INM.'}
            df_agrupado_copia = self.df_agrupado

            df_agrupado_copia = df_agrupado_copia.replace(REEMPLAZOS)

            self.df_stockXDepo = self.df_stockXDepo.rename(columns={"Cantidad" : "Cantidad_Sistema"})
            df_agrupado_copia = df_agrupado_copia.rename(columns={"Cantidad":"Cantidad_Inventario"})
            df_agrupado_copia = df_agrupado_copia.rename(columns={"Stock o Cliente":"Deposito"})
            
            self.df_comparacion = pandas.merge(self.df_stockXDepo, df_agrupado_copia, on=['Codigo', 'Deposito'], how='outer')
            self.df_comparacion['Cantidad_Sistema'] = self.df_comparacion['Cantidad_Sistema'].fillna(0)
            self.df_comparacion['Cantidad_Inventario'] = self.df_comparacion['Cantidad_Inventario'].fillna(0)
            self.df_comparacion['Diferencia'] = self.df_comparacion['Cantidad_Inventario'] - self.df_comparacion['Cantidad_Sistema']

            self.df_comparacion = self.df_comparacion[['Producto', 'Codigo', 'Deposito', 'Cantidad_Sistema', 'Cantidad_Inventario', 'Diferencia', 'Unidad', 'Familia', 'Activo']]
        except Exception as e:
            messagebox.showerror("Error al comparar inventarios", f"{e}")

