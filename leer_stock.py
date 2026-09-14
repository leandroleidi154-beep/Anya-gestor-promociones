import pandas as pd

from limpiar import (
    buscar_encabezados_stock,
    buscar_columnas,
    limpiar_codigo,
    limpiar_precio,
    limpiar_cantidad
)
from formatos import leer_tabla_bruta


def leer_stock(archivo):

    stock = leer_tabla_bruta(archivo)

    fila_stock = buscar_encabezados_stock(stock)

    if fila_stock is None:
        raise ValueError("No se encontraron los encabezados de Stock.")

    columnas_stock = buscar_columnas(
        stock,
        fila_stock
    )

    columnas_necesarias = [
        "codigo",
        "descripcion",
        "stock",
        "precio"
    ]

    for columna in columnas_necesarias:
        if columna not in columnas_stock:
            raise ValueError(
                f"No se encontró la columna '{columna}' en el archivo de Stock."
            )

    resultados = []

    for i in range(fila_stock + 1, len(stock)):

        codigo = stock.iloc[
            i, columnas_stock["codigo"]
        ]

        if pd.isna(codigo):
            continue

        codigo = limpiar_codigo(codigo)

        if codigo == "":
            continue

        descripcion = stock.iloc[
            i, columnas_stock["descripcion"]
        ]

        cantidad = stock.iloc[
            i, columnas_stock["stock"]
        ]

        cantidad = limpiar_cantidad(cantidad)

        precio = stock.iloc[
            i, columnas_stock["precio"]
        ]

        precio = limpiar_precio(precio)

        resultados.append({
            "Código de Barras": codigo,
            "Descripción": descripcion,
            "Stock": cantidad,
            "Precio de Venta": precio
        })

    return pd.DataFrame(resultados)
