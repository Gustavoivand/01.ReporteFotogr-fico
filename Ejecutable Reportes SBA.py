import tkinter as tk
from tkinter import messagebox
#UTILIZA ESTA PARTE DEL CODIGO PARA INSTALAR TODOS LOS COMPONENTES NECESARIOS:
#py -m pip install pillow
#py -m pip install openpyxl
#py -m pip install natsort
#py -m pip install tk        
#py -m pip install translate
#Esta linea inicia la rutina y todos los comandos necesarios para la ejecución del código
    
from PIL import JpegImagePlugin
from decimal import Decimal
import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker 
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU, cm_to_EMU
JpegImagePlugin._getmp = lambda: None
import glob
import os
from natsort import natsorted


#Esta linea define los bordes de las fotos a ingresar
thick_border = Border(left=Side(style='thick'), 
                     right=Side(style='thick'), 
                     top=Side(style='thick'), 
                     bottom=Side(style='thick'))

half_thick = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thick'), 
                     bottom=Side(style='thick'))

import translate
from translate import Translator

def translate_phrase(phrase, target_language='en'):
    translator= Translator(to_lang=target_language,from_lang='es',provider='mymemory',email='gustavoivand@gmail.com')
    translation = translator.translate(phrase)
    
    return translation

class MyGUI:
    def __init__(self):
        self.root =tk.Tk(className="-REPORTES DD SBA-")
        self.label0=tk.Label(self.root,text="Reportes DD SBA",font=('Arial',18))
        self.label0.pack(padx=10,pady=10)
        #BUSCAR DOCUMENTO BASE DE REPORTE
        self.button =tk.Button(self.root, text="Buscar documento Base de reporte", font=('Arial',12))
        self.button.pack(padx=10,pady=10)
        self.label=tk.Label(self.root,text="Dirección:",font=('Arial',8))
        self.label.pack(padx=5,pady=5)
        #BUSCAR DOCUMENTO DE BASE DE DATOS
        self.button2 =tk.Button(self.root, text="Buscar documento Base de datos (opcional)", font=('Arial',12))
        self.button2.pack(padx=10,pady=10)
        self.button21 =tk.Button(self.root, text="Introduce el código SBA del sitio (opcional)", font=('Arial',12))
        self.button21.pack(padx=10,pady=10)
        self.label2=tk.Label(self.root,text="Dirección:",font=('Arial',8))
        self.label2.pack(padx=10,pady=10)
        self.label21=tk.Label(self.root,text="Código:",font=('Arial',8))
        self.label21.pack(padx=10,pady=10)
        #BUSCAR CARPETA DE REPORTE
        self.button3 =tk.Button(self.root, text="Buscar carpeta con las imágenes ordenadas", font=('Arial',12),command=self.buscar_carpeta)
        self.button3.pack(padx=10,pady=10)
        self.label3=tk.Label(self.root,text="Dirección:",font=('Arial',8))
        self.label3.pack(padx=10,pady=10)
        #EJECUTAR REPORTE
        self.button4 =tk.Button(self.root, text="Ejecutar Reporte", font=('Arial',12))
        self.button4.pack(padx=10,pady=10)
        #BUSCAR DOCUMENTO DE SALIDA 1_ESPAÑOL
        self.button5 =tk.Button(self.root, text="Buscar documento editado", font=('Arial',12))
        self.button5.pack(padx=10,pady=10)
        self.label5=tk.Label(self.root,text="Dirección:",font=('Arial',8))
        self.label5.pack(padx=10,pady=10)
        #TRADUCIR DOCUMENTO
        self.button6 =tk.Button(self.root, text="Traducir documento", font=('Arial',12))
        self.button6.pack(padx=10,pady=10)
        self.root.mainloop()
    def buscar_carpeta(self):
        pass
        import tkinter
        import tkinter.filedialog as tkFileDialog
        root = tkinter.Tk()
        root.lift()
        root.attributes('-topmost',True)
        root.after_idle(root.attributes,'-topmost',False)
        root.withdraw()
        folder_path = tkFileDialog.askdirectory()
        self.label3.config(text=folder_path)
        print(folder_path)
        root.destroy()
    def buscar_doc_Breporte(self):
        pass
        #SELECCIONAR EXCEL BASE
        import tkinter
        import tkinter.filedialog as tkFileDialog

        root = tkinter.Tk()

        root.lift()
        root.attributes('-topmost',True)
        root.after_idle(root.attributes,'-topmost',False)
        root.withdraw()
        file_path = tkFileDialog.askopenfile()
        print(file_path.name)
        root.destroy()
    def buscar_doc_BDatos(self):
        pass
        #SELECCIONAR BASE DE DATOS
        import tkinter
        import tkinter.filedialog as tkFileDialog

        root = tkinter.Tk()

        root.lift()
        root.attributes('-topmost',True)
        root.after_idle(root.attributes,'-topmost',False)
        root.withdraw()
        file_BD = tkFileDialog.askopenfile()
        #print(file_BD.name)
        root.destroy()
    def buscar_doc_Reportelisto(self):
        pass



MyGUI()