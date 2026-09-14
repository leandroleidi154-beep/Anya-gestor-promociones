"""
Algunos sistemas de farmacia "exportan a Excel" pero en realidad generan
una página HTML con una tabla adentro, y le ponen la extensión .xls
por costumbre. Este archivo se encarga de detectar esos casos y leerlos
igual, además de los archivos de Excel y LibreOffice normales.
"""

import pandas as pd


def es_html_disfrazado(archivo):
    """
    Mira los primeros bytes del archivo para saber si en realidad es
    una página HTML (aunque el nombre diga .xls o .xlsx), en vez de
    un archivo de Excel/LibreOffice de verdad.
    """

    with open(archivo, "rb") as f:
        inicio = f.read(1000)

    texto = inicio.decode("latin-1", errors="ignore").lower()

    return "<html" in texto or "<table" in texto


def leer_tabla_bruta(archivo):
    """
    Lee cualquiera de estos formatos y devuelve siempre un DataFrame
    "crudo" (sin asignar encabezados todavía), tal como si hubiéramos
    hecho pd.read_excel(archivo, header=None):

    - Excel real (.xlsx, .xls)
    - LibreOffice (.ods)
    - HTML disfrazado de .xls (algunos sistemas de farmacia exportan así)
    """

    if es_html_disfrazado(archivo):
        tablas = pd.read_html(archivo, encoding="iso-8859-1", header=None)
        return tablas[0]

    return pd.read_excel(archivo, header=None)
