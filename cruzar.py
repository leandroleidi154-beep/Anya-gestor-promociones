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


def _parsear_fecha_dd_mm_aaaa(texto):
    """
    Intenta convertir un texto "dd/mm/aaaa" a fecha real. Si el texto
    no tiene ese formato (por ejemplo "hasta agotar stock"), devuelve
    None en vez de romper, para que esas promociones no se pierdan
    al filtrar por período.
    """
    if not texto:
        return None

    try:
        return datetime.strptime(str(texto).strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def filtrar_por_periodo(promociones, fecha_desde=None, fecha_hasta=None):
    """
    Se queda solo con las promociones cuyo período (Inicio-Fin) se
    superpone con el rango [fecha_desde, fecha_hasta] elegido por el
    usuario. Ambos límites son opcionales (se puede filtrar solo con
    "Desde", solo con "Hasta", o con los dos).

    Las promociones cuya fecha de Inicio o Fin no se puede interpretar
    como fecha real (texto libre tipo "hasta agotar stock") se dejan
    pasar siempre, porque no hay forma de saber si entran o no en el
    rango.
    """

    if fecha_desde is None and fecha_hasta is None:
        return promociones

    def cumple(fila):
        fecha_inicio = _parsear_fecha_dd_mm_aaaa(fila.get("Inicio"))
        fecha_fin = _parsear_fecha_dd_mm_aaaa(fila.get("Fin"))

        if fecha_hasta is not None and fecha_inicio is not None and fecha_inicio > fecha_hasta:
            return False

        if fecha_desde is not None and fecha_fin is not None and fecha_fin < fecha_desde:
            return False

        return True

    return promociones[promociones.apply(cumple, axis=1)]


def cruzar_promociones_stock(promociones, stock):
    if promociones.empty or stock.empty:
        return pd.DataFrame()

    # Cruce directo vectorizado O(N + M)
    resultado = pd.merge(
        promociones,
        stock,
        on="Código de Barras",
        suffixes=("_promo", "_stock")
    )

    if resultado.empty:
        return pd.DataFrame()

    # Asignación de descripción y cálculo de estado
    resultado["Descripción"] = resultado["Descripción_stock"]
    resultado["Estado"] = resultado.apply(
        lambda row: calcular_estado(row["Inicio"], row["Fin"]), axis=1
    )

    # Identificación de duplicados
    conteo_por_codigo = resultado["Código de Barras"].value_counts()
    if (conteo_por_codigo > 1).any():
        resultado["Revisar duplicado"] = resultado["Código de Barras"].apply(
            lambda c: "SI" if conteo_por_codigo[c] > 1 else ""
        )

    columnas_finales = [
        "Código de Barras", "Descripción", "Linea", "Proveedor",
        "Hoja", "DTO", "Precio de Venta", "Inicio", "Fin", "Estado", "Stock"
    ]
    if "Revisar duplicado" in resultado.columns:
        columnas_finales.append("Revisar duplicado")

    return resultado[columnas_finales]