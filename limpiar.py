import re
from datetime import datetime

import pandas as pd


MESES_ESPAÑOL = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4,
    "may": 5, "jun": 6, "jul": 7, "ago": 8,
    "sep": 9, "set": 9, "oct": 10, "nov": 11, "dic": 12
}


def limpiar_codigo(codigo):
    if pd.isna(codigo):
        return ""

    codigo = str(codigo).strip()

    if codigo.endswith(".0"):
        codigo = codigo[:-2]

    # A veces Excel guarda el mismo código de barras con un 0 adelante
    # en un archivo, y sin ese 0 en el otro (lo "come" al tratarlo como
    # número). Sacamos los ceros a la izquierda para que ambos lados
    # queden iguales y se puedan cruzar bien.
    codigo_sin_ceros = codigo.lstrip("0")

    return codigo_sin_ceros if codigo_sin_ceros else codigo

def limpiar_precio(precio):
    """
    Convierte el precio a un número, sin importar si ya viene como
    número (Excel normal) o como texto con formato argentino
    (por ejemplo "32.711,20", con punto de miles y coma decimal,
    como pasa en el archivo que exportan algunos sistemas de farmacia).
    """
    if pd.isna(precio):
        return None

    if isinstance(precio, (int, float)):
        return float(precio)

    texto = str(precio).strip()
    texto = texto.replace(".", "").replace(",", ".")

    try:
        return float(texto)
    except ValueError:
        return precio


def limpiar_cantidad(cantidad):
    """
    Convierte la cantidad de stock a un número entero, sin importar
    si viene como número (Excel normal) o como texto (formato HTML).
    """
    if pd.isna(cantidad):
        return 0

    if isinstance(cantidad, (int, float)):
        return int(cantidad)

    try:
        return int(str(cantidad).strip())
    except ValueError:
        return cantidad


def limpiar_dto(dto):
    if pd.isna(dto):
        return ""

    if isinstance(dto, str):
        dto = dto.strip()

        if dto.endswith("%"):
            return dto

        return dto

    if isinstance(dto, (int, float)):
        if 0 < dto < 1:
            return f"{dto * 100:g}%"

        return str(dto)

    return str(dto)

def limpiar_fecha(fecha):
    """
    Convierte la fecha a texto dd/mm/aaaa. Además del caso normal
    (una fecha real de Excel), maneja un caso especial: algunas hojas
    traen la fecha como texto SIN año, por ejemplo "1-sep" o "30-sep".
    Si se lo dejamos a pandas, lo interpreta mal (confunde el "1" con
    el año 1), así que ese formato lo resolvemos a mano primero,
    usando el año actual.
    """
    if pd.isna(fecha):
        return ""

    # Si ya es una fecha real (Excel la reconoció bien), la usamos directo.
    if isinstance(fecha, (pd.Timestamp, datetime)):
        return fecha.strftime("%d/%m/%Y")

    texto = str(fecha).strip()

    # Caso especial: texto tipo "1-sep", "30-sep", "1/sep", sin año.
    coincidencia = re.match(r"^(\d{1,2})[-/]([A-Za-zÁÉÍÓÚáéíóú]{3,})$", texto)

    if coincidencia:
        dia = int(coincidencia.group(1))
        mes_texto = coincidencia.group(2)[:3].lower()
        mes_texto = (
            mes_texto.replace("á", "a").replace("é", "e")
            .replace("í", "i").replace("ó", "o").replace("ú", "u")
        )
        mes = MESES_ESPAÑOL.get(mes_texto)

        if mes:
            anio_actual = datetime.now().year
            try:
                return datetime(anio_actual, mes, dia).strftime("%d/%m/%Y")
            except ValueError:
                pass  # fecha inválida (ej. 31 de febrero); seguimos abajo

    try:
        fecha_convertida = pd.to_datetime(texto, dayfirst=True)
        return fecha_convertida.strftime("%d/%m/%Y")
    except Exception:
        return texto

def limpiar_stock(stock):
    stock["Código de Barras"] = stock["Código de Barras"].apply(limpiar_codigo)

    return stock


def limpiar_promociones(promociones, columna_codigo):
    promociones.iloc[:, columna_codigo] = promociones.iloc[:, columna_codigo].apply(limpiar_codigo)

    return promociones

def buscar_encabezados(tabla):
    for numero_fila, fila in tabla.iterrows():

        valores = []

        for valor in fila:
            if pd.notna(valor):
                valores.append(str(valor).strip().upper())

        tiene_codigo = "EAN" in valores or "COD" in valores
        tiene_dto = "DTO" in valores
        tiene_inicio = "INICIO" in valores
        tiene_fin = "FIN" in valores

        if tiene_codigo and tiene_dto and tiene_inicio and tiene_fin:
            return numero_fila

    return None

def buscar_columnas(tabla, fila_encabezados):
    fila = tabla.iloc[fila_encabezados]

    columnas = {}

    for numero_columna, valor in enumerate(fila):
        if pd.isna(valor):
            continue

        nombre = str(valor).strip().upper()

        if nombre in ("EAN", "COD", "CODIGO", "CÓDIGO", "CODIGO DE BARRAS", "CÓDIGO DE BARRAS"):
            columnas["codigo"] = numero_columna

        elif nombre in ("DESCRIPCION", "DESCRIPCIÓN"):
            columnas["descripcion"] = numero_columna

        elif nombre == "DTO":
            columnas["dto"] = numero_columna

        elif nombre == "INICIO":
            columnas["inicio"] = numero_columna

        elif nombre == "FIN":
            columnas["fin"] = numero_columna

        elif nombre == "STOCK":
            columnas["stock"] = numero_columna

        elif nombre == "PRECIO DE VENTA":
            columnas["precio"] = numero_columna

        elif nombre == "PROVEEDOR":
            columnas["proveedor"] = numero_columna

        elif nombre in ("LINEA", "LÍNEA"):
            columnas["linea"] = numero_columna

    return columnas


def buscar_encabezados_stock(tabla):
    for numero_fila, fila in tabla.iterrows():

        valores = []

        for valor in fila:
            if pd.notna(valor):
                valores.append(str(valor).strip().upper())

        tiene_codigo = (
            "CODIGO DE BARRAS" in valores
            or "CÓDIGO DE BARRAS" in valores
        )

        tiene_stock = "STOCK" in valores
        tiene_precio = "PRECIO DE VENTA" in valores

        if tiene_codigo and tiene_stock and tiene_precio:
            return numero_fila

    return None