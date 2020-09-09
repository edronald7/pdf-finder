import os
import PyPDF2
import pandas as pd
from tkinter import filedialog
from tkinter import *
from tkinter import scrolledtext

class ProcesarPDF():
    def __init__(self):
        pass

    def procesar(self, ruta, lista_frases):
        # leer el contenido de todos los archivo de una carpeta dada
        archivos_en_dir = os.listdir(ruta)

        data = []

        for archivo in archivos_en_dir: #recore los archivos
            if archivo.lower().endswith('.pdf'):
                ruta_archivo = ruta + archivo
                abre_pdf = open(ruta_archivo, 'rb') #abre el pdf
                lector_pdf = PyPDF2.PdfFileReader(abre_pdf) #lee el pdf
                num_paginas = lector_pdf.numPages
                print("El archivo {} tiene {} páginas.".format(archivo, num_paginas))
                for indice_pagina in range(num_paginas): #lee las paginas del pdf
                    pagina = lector_pdf.getPage(indice_pagina)
                    contenido = pagina.extractText() #extrae el texto de la pagina

                    resultado_existe, frases_encontradas = self.comparar(contenido, lista_frases)
                    if resultado_existe:
                        for frase in frases_encontradas:
                            #registro = '{},{},{}'.format(archivo, indice_pagina, frase)
                            data.append({
                                'ARCHIVO': archivo,
                                'PAGINA': indice_pagina,
                                'FRASE': frase
                            })
        df_data = pd.DataFrame(data)
        df_data.to_excel('resultado_busqueda.xlsx', index=False)

    def comparar(self, contenido, lista_frases):
        #print(contenido)
        existe_frase = False
        frases_encontradas = []
        for frase in lista_frases:
            if frase.lower() in contenido.lower():
                existe_frase = True
                separados_por_punto = contenido.split('.')
                for parrafo in separados_por_punto:
                    if frase.lower() in parrafo.lower():
                        frases_encontradas.append(" ".join([p.strip() for p in parrafo.split('\n') if len(p.strip())>0 ]))
                break
        #print(existe_frase, frases_encontradas)
        return existe_frase, frases_encontradas


if __name__=='__main__':

    window = Tk()
    window.title("Procesar PDF")
    window.geometry('370x300')

    lbl = Label(window, text="Seleccione Carpeta de PDF")
    lbl.grid(column=0, row=0, padx=(10, 10), pady=(5, 5))

    def seleccionar():
        folder_selected = filedialog.askdirectory()
        ruta_pdf = folder_selected if folder_selected.endswith('/') else folder_selected+'/'
        txt1.insert(0, ruta_pdf)

    def procesar():
        lbl.configure(text="Procesando...")
        ruta_pdf = txt1.get()
        texto = txt2.get('1.0', END)
        lista_temp = texto.split('\n')
        lista_frases = [frase.strip() for frase in lista_temp if len(frase)>0]
        procesar = ProcesarPDF()
        procesar.procesar(ruta_pdf, lista_frases)
        lbl.configure(text="Listo !")

    btn1 = Button(window, text="Seleccionar Carpeta", command=seleccionar)
    btn1.grid(column=0, row=1, padx=(10, 10), pady=(5, 5))

    txt1 = Entry(window, width=50)
    txt1.grid(column=0, row=2, padx=(10, 10), pady=(5, 5))

    txt2 = scrolledtext.ScrolledText(window, width=40, height=8)
    txt2.grid(column=0, row=3, padx=(10, 10), pady=(5, 5))

    btn2 = Button(window, text="Procesar", command=procesar)
    btn2.grid(column=0, row=4, padx=(10, 10), pady=(5, 5))

    window.mainloop()

