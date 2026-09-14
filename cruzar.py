import pandas as pd
from datetime import datetime, timedelta


def calcular_estado(fecha_inicio_texto, fecha_fin_texto):
    """
    Mira las fechas de Inicio y Fin de la promoción y devuelve una
    palabra que resume la situación:

    - "PRÓXIMAMENTE"  -> todavía no arrancó (el Inicio es una fecha futura)
    - "VENCIDA"       -> la fecha de fin ya pasó
    - "VENCE MAÑANA"  -> falta exactamente 1 día para el Fin
    - "VIGENTE"       -> está corriendo ahora
    - "SIN FECHA"     -> la celda de Fin está vacía

    Si alguna fecha no tiene formato dd/mm/aaaa (por ejemplo, dice
    "hasta agotar stock" o cualquier otra frase), no se puede evaluar
    si es futura, así que seguimos con la lógica normal usando el Fin.
    """

    hoy = datetime.now().date()

    # Primero chequeamos si todavía no empezó (esto tiene prioridad:
    # una promo que arranca la semana que viene no está "vigente" hoy,
    # aunque su fecha de fin sea lejana).
    if fecha_inicio_texto:
        texto_inicio = str(fecha_inicio_texto).strip()
        try:
            fecha_inicio = datetime.strptime(texto_inicio, "%d/%m/%Y")
            if fecha_inicio.date() > hoy:
                return "PRÓXIMAMENTE"
        except ValueError:
            pass  # Inicio no es una fecha reconocible; seguimos abajo.

    if not fecha_fin_texto:
        return "SIN FECHA"

    texto_fin = str(fecha_fin_texto).strip()

    try:
        fecha_fin = datetime.strptime(texto_fin, "%d/%m/%Y")
    except ValueError:
        # No es una fecha reconocible: reproducimos el texto original.
        return texto_fin

    manana = hoy + timedelta(days=1)

    if fecha_fin.date() < hoy:
        return "VENCIDA"

    if fecha_fin.date() == manana:
        return "VENCE MAÑANA"

    return "VIGENTE"


def cruzar_promociones_stock(promociones, stock):

    resultados = []

    for _, promocion in promociones.iterrows():

        codigo_promocion = promocion["Código de Barras"]

        coincidencias = stock[
            stock["Código de Barras"] == codigo_promocion
        ]

        for _, producto in coincidencias.iterrows():

            resultados.append({
                "Código de Barras": codigo_promocion,
                "Descripción": producto["Descripción"],
                "Linea": promocion.get("Linea", ""),
                "Proveedor": promocion.get("Proveedor", ""),
                "DTO": promocion["DTO"],
                "Precio de Venta": producto["Precio de Venta"],
                "Inicio": promocion["Inicio"],
                "Fin": promocion["Fin"],
                "Estado": calcular_estado(promocion["Inicio"], promocion["Fin"]),
                "Stock": producto["Stock"],
            })

    resultado = pd.DataFrame(resultados)

    if not resultado.empty:

        # Si el mismo código de barras aparece más de una vez
        # (por ejemplo, promocionado en dos hojas distintas),
        # lo marcamos para que se revise a mano.
        # Esta columna solo se agrega si realmente hay algún duplicado.
        conteo_por_codigo = resultado["Código de Barras"].value_counts()

        hay_duplicados = (conteo_por_codigo > 1).any()

        if hay_duplicados:
            resultado["Revisar duplicado"] = resultado["Código de Barras"].apply(
                lambda codigo: "SI" if conteo_por_codigo[codigo] > 1 else ""
            )

    return resultado
