import pandas as pd

from limpiar import (
    buscar_encabezados,
    buscar_columnas,
    limpiar_codigo,
    limpiar_dto,
    limpiar_fecha
)


def leer_promociones(archivo):

    # sheet_name=None le dice a pandas "traeme TODAS las hojas".
    # El resultado es un diccionario: {"nombre de la hoja": tabla}
    todas_las_hojas = pd.read_excel(
        archivo,
        header=None,
        sheet_name=None
    )

    columnas_necesarias = [
        "codigo",
        "descripcion",
        "dto",
        "inicio",
        "fin"
    ]

    resultados = []

    for nombre_hoja, promociones in todas_las_hojas.items():

        fila_promociones = buscar_encabezados(promociones)

        if fila_promociones is None:
            print(f"  [Aviso] La hoja '{nombre_hoja}' no tiene encabezados reconocibles, se saltea.")
            continue

        columnas_promociones = buscar_columnas(
            promociones,
            fila_promociones
        )

        faltantes = [
            columna for columna in columnas_necesarias
            if columna not in columnas_promociones
        ]

        if faltantes:
            print(f"  [Aviso] A la hoja '{nombre_hoja}' le faltan columnas {faltantes}, se saltea.")
            continue

        # Extraer el proveedor desde la celda A1 si existe, o usar el nombre de la hoja
        proveedor_encabezado = ""
        try:
            val_a1 = str(promociones.iloc[0, 0]).strip()
            if val_a1 and val_a1.upper() != "NAN":
                proveedor_encabezado = val_a1
            else:
                proveedor_encabezado = nombre_hoja
        except Exception:
            proveedor_encabezado = nombre_hoja

        for i in range(fila_promociones + 1, len(promociones)):

            codigo = promociones.iloc[
                i, columnas_promociones["codigo"]
            ]

            if pd.isna(codigo):
                continue

            codigo = limpiar_codigo(codigo)

            if codigo == "":
                continue

            descripcion = promociones.iloc[
                i, columnas_promociones["descripcion"]
            ]

            dto = promociones.iloc[
                i, columnas_promociones["dto"]
            ]

            dto = limpiar_dto(dto)

            inicio = promociones.iloc[
                i, columnas_promociones["inicio"]
            ]

            inicio = limpiar_fecha(inicio)

            fin = promociones.iloc[
                i, columnas_promociones["fin"]
            ]

            fin = limpiar_fecha(fin)

            # Si la columna "proveedor" existe en la tabla (Septiembre), la lee.
            # Si NO existe (Octubre), usa el valor sacado del título / nombre de hoja.
            if "proveedor" in columnas_promociones:
                val_prov = promociones.iloc[i, columnas_promociones["proveedor"]]
                proveedor = str(val_prov).strip() if pd.notna(val_prov) else proveedor_encabezado
            else:
                proveedor = proveedor_encabezado

            if "linea" in columnas_promociones:
                linea = promociones.iloc[
                    i, columnas_promociones["linea"]
                ]
            else:
                linea = ""

            resultados.append({
                "Código de Barras": codigo,
                "Descripción": descripcion,
                "DTO": dto,
                "Inicio": inicio,
                "Fin": fin,
                "Proveedor": proveedor,
                "Linea": linea,
                "Hoja": nombre_hoja
            })

    return pd.DataFrame(resultados)
