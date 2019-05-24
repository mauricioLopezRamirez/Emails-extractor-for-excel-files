import tkinter as tk
from tkinter import filedialog
import tkMessageBox
import pandas as pd
import re

# extrac emails
def extraerCorreos(txt):
    emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", str(txt)) 
    #almacenamos los correos en un archivo txt
    
    if len(emails) != 0:
        with open('emails.txt', 'w') as f:
            for item in emails:
                f.write("%s\n" % item)
        tkMessageBox.showinfo("Proceso finalizado", "Se ha generado un archivo txt con "+str(len(emails))+" correos extraidos.")
    else:
        tkMessageBox.showinfo("Proceso finalizado", "No se han encontrado correos en el documento seleccionado")
root= tk.Tk()
canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')
canvas1.pack()

def getExcel ():
    global xl
    #ventana para seleccionar el archivo
    import_file_path = filedialog.askopenfilename()
    # cargar el archivo
    xl = pd.ExcelFile(import_file_path)
    # numero de hojas que tiene el documento
    num_hojas = len(xl.sheet_names)
    contenido_hoja = ""
    
    for sheet in xl.sheet_names:
        #establecemos los limites
        pd.options.display.max_rows = xl.parse(sheet).shape[0]
        pd.options.display.max_columns = xl.parse(sheet).shape[1]
        # sacamos la informacion de la hoja en curso
        contenido_hoja = contenido_hoja + str(xl.parse(sheet))
    
    extraerCorreos(contenido_hoja)


browseButton_Excel = tk.Button(text='Import Excel File', command=getExcel, bg='blue', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(100, 20, window=browseButton_Excel)
root.mainloop()


