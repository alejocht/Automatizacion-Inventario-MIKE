import tkinter as tk
from tkinter import ttk
import negocio as neg

class VentanaPrincipal(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("900x300")
        self.title("Consolidador Excel")

        self.frame_principal = ttk.Frame(self, padding=10)
        self.frame_principal.pack(fill="both", expand=True)
        
        self.txtCarpetaFuente = tk.Text(self.frame_principal, width=50, height=1, font="Arial 11", state="disabled")
        self.txtCarpetaFuente.grid(row=0, column=1, pady=10, padx=5)

        self.txtStockXDepo = tk.Text(self.frame_principal, width=50, height=1, font="Arial 11", state="disabled")
        self.txtStockXDepo.grid(row=2, column=1, pady=10, padx=5)
        
        self.txtArchivoDestino = tk.Text(self.frame_principal, width=50, height=1, font="Arial 11", state="disabled")
        self.txtArchivoDestino.grid(row=4, column=1, pady=10, padx=5)

        self.btnElegirCarpetaFuente = ttk.Button(self.frame_principal, text="Carpeta Fuente ...", command=lambda:neg.ingresar_Carpeta_Fuente(self.txtCarpetaFuente))
        self.btnElegirCarpetaFuente.grid(row=0, column=6)

        self.btnElegirStockXDepo = ttk.Button(self.frame_principal, text="Stock x Deposito ...", command=lambda:neg.ingresar_Stock_X_Deposito(self.txtStockXDepo))
        self.btnElegirStockXDepo.grid(row=2, column=6)

        self.btnElegirDestino = ttk.Button(self.frame_principal, text="Guardar Como ...", command=lambda:neg.guardar_Como(self.txtArchivoDestino))
        self.btnElegirDestino.grid(row=4, column=6)

        self.btnProcesar = ttk.Button(self.frame_principal, text="Procesar Datos", command=neg.procesar_Datos)
        self.btnProcesar.grid(row=5, column=0, columnspan=2, pady=10)

        self.btnAyuda = ttk.Button(self.frame_principal, text="Ayuda", command=neg.ayuda)
        self.btnAyuda.grid(row=5, column=2, columnspan=2, pady=10)