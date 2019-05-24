import pandas as pd
import re

# extrac emails
def extraerCorreos(txt):
    emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", str(txt)) 
    #almacenamos los correos en un archivo txt
    with open('emails.txt', 'w') as f:
        for item in emails:
            f.write("%s\n" % item)

# nombre o ruta del archivo a analizar
file = 'programacion.xlsx'

# cargar el archivo
xl = pd.ExcelFile(file)

# numero de hojas que tiene el documento
num_hojas = len(xl.sheet_names)
contenido_hoja = ""

for sheet in xl.sheet_names:
    # sacamos la informacion de la hoja en curso
    contenido_hoja = contenido_hoja + str(xl.parse(sheet).values)


extraerCorreos(contenido_hoja)

