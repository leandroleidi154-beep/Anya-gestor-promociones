import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# Colores usados en el Excel
COLOR_ENCABEZADO = "2F5496"
COLOR_DUPLICADO = "FFF2CC"     # amarillo suave
COLOR_VENCIDA = "F8CBAD"       # rojo/salmón suave
COLOR_VENCE_MANANA = "FCE4D6"  # naranja suave
COLOR_PROXIMAMENTE = "D9E1F2"  # celeste suave


def generar_excel(resultado, archivo_salida, dividir_control=False):
    """
    Recibe el DataFrame ya cruzado (resultado) y genera el Excel final.
    """
    libro = openpyxl.Workbook()

    _armar_hoja_principal(libro, resultado)

    if dividir_control:
        _armar_hojas_control_divididas(libro, resultado)
    else:
        _armar_hoja_control(libro, "Control", resultado, titulo_personalizado="CONTROL GENERAL DE PROMOCIONES")

    libro.save(archivo_salida)


# ============================================================
# HOJA 1: PROMOCIONES EN STOCK (listado completo)
# ============================================================

def _armar_hoja_principal(libro, resultado):
    hoja = libro.active
    hoja.title = "Promociones en Stock"

    columnas = [
        ("Linea", "Línea"),
        ("Descripción", "Descripción"),
        ("Código de Barras", "Código de Barras"),
        ("Stock", "Stock"),
        ("DTO", "Promoción"),
        ("Precio de Venta", "Precio de Venta"),
        ("Inicio", "Inicio"),
        ("Fin", "Fin"),
        ("Estado", "Estado"),
    ]

    tiene_duplicados = "Revisar duplicado" in resultado.columns

    if tiene_duplicados:
        columnas.append(("Revisar duplicado", "Control"))

    encabezados = [nombre_mostrado for _, nombre_mostrado in columnas]
    hoja.append(encabezados)

    for _, fila in resultado.iterrows():
        hoja.append([fila[columna_original] for columna_original, _ in columnas])

    _dar_estilo_encabezado(hoja)

    posiciones = {nombre: i + 1 for i, (_, nombre) in enumerate(columnas)}

    _resaltar_por_estado(hoja, posiciones.get("Estado"))

    if tiene_duplicados:
        _resaltar_duplicados(hoja, posiciones.get("Control"))

    if "Precio de Venta" in posiciones:
        _formato_moneda(hoja, posiciones["Precio de Venta"])

    anchos = {
        "A": 18, "B": 40, "C": 16, "D": 8,
        "E": 12, "F": 15, "G": 13, "H": 13,
        "I": 14, "J": 12
    }
    for letra, ancho in anchos.items():
        hoja.column_dimensions[letra].width = ancho

    hoja.freeze_panes = "A2"
    _configurar_impresion(hoja)


# ============================================================
# HOJA 2: CONTROL (para imprimir y tildar en el local)
# ============================================================

def _armar_hoja_control(libro, nombre_hoja, datos, titulo_personalizado=None):
    hoja = libro.create_sheet(nombre_hoja)

    # 1. Título institucional en la Fila 1 (Combinado de A1 hasta J1)
    hoja.merge_cells("A1:J1")
    celda_titulo = hoja.cell(1, 1)
    
    if not titulo_personalizado:
        titulo_personalizado = f"CONTROL DE PROMOCIONES - {nombre_hoja.replace('Control - ', '').upper()}"
        
    celda_titulo.value = titulo_personalizado
    celda_titulo.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    
    relleno_titulo = PatternFill(start_color="39476A", end_color="39476A", fill_type="solid")
    celda_titulo.fill = relleno_titulo
    celda_titulo.alignment = Alignment(horizontal="center", vertical="center")
    hoja.row_dimensions[1].height = 25

    # 2. Fila 2 de separación (vacía y finita)
    hoja.row_dimensions[2].height = 10

    # 3. Encabezados de la tabla exactamente en la Fila 3
    encabezados = [
        "Línea", "Descripción", "Código de Barras", "Stock",
        "Promoción", "Precio de Venta", "Inicio", "Fin", "Vigencia", "Control"
    ]
    for col_idx, enc in enumerate(encabezados, start=1):
        hoja.cell(row=3, column=col_idx, value=enc)
    
    hoja.row_dimensions[3].height = 22
    _dar_estilo_encabezado_fila(hoja, fila_num=3)

    # 4. Inserción de datos a partir de la Fila 4
    datos_ordenados = datos.sort_values(["Linea", "Descripción"])
    fila_actual = 4

    for _, fila in datos_ordenados.iterrows():
        hoja.cell(row=fila_actual, column=1, value=fila["Linea"])
        hoja.cell(row=fila_actual, column=2, value=fila["Descripción"])
        hoja.cell(row=fila_actual, column=3, value=fila["Código de Barras"])
        hoja.cell(row=fila_actual, column=4, value=fila["Stock"])
        hoja.cell(row=fila_actual, column=5, value=fila["DTO"])
        hoja.cell(row=fila_actual, column=6, value=fila["Precio de Venta"])
        hoja.cell(row=fila_actual, column=7, value=fila["Inicio"])
        hoja.cell(row=fila_actual, column=8, value=fila["Fin"])
        hoja.cell(row=fila_actual, column=9, value=fila["Estado"])
        hoja.cell(row=fila_actual, column=10, value="")
        fila_actual += 1

    columna_precio = 6
    columna_vigencia = 9
    columna_control = 10

    _resaltar_por_estado_desde_fila(hoja, columna_vigencia, fila_inicio=4)
    _formato_moneda_desde_fila(hoja, columna_precio, fila_inicio=4)

    # Cuadrícula completa desde la fila 3 en adelante
    borde = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    for fila in range(3, hoja.max_row + 1):
        for columna in range(1, hoja.max_column + 1):
            hoja.cell(fila, columna).border = borde

    for fila in range(4, hoja.max_row + 1):
        celda_control = hoja.cell(fila, columna_control)
        celda_control.alignment = Alignment(horizontal="center")
        hoja.row_dimensions[fila].height = 20

    anchos = {
        "A": 18, "B": 42, "C": 16, "D": 8,
        "E": 12, "F": 15, "G": 13, "H": 13, "I": 14, "J": 12
    }
    for letra, ancho in anchos.items():
        hoja.column_dimensions[letra].width = ancho

    # Inmovilizar paneles debajo de los encabezados (fila 3)
    hoja.freeze_panes = "A4"

    # Configuración de impresión repitiendo la fila 3 como cabecera
    hoja.page_setup.orientation = "landscape"
    hoja.page_setup.fitToWidth = 1
    hoja.page_setup.fitToHeight = 0
    hoja.sheet_properties.pageSetUpPr.fitToPage = True
    hoja.print_title_rows = "3:3"

    hoja.page_margins.left = 0.4
    hoja.page_margins.right = 0.4
    hoja.page_margins.top = 0.5
    hoja.page_margins.bottom = 0.5
    hoja.page_margins.header = 0.2
    hoja.page_margins.footer = 0.2


# Caracteres que Excel no permite en un nombre de hoja.
_CARACTERES_INVALIDOS_NOMBRE_HOJA = ["\\", "/", "*", "[", "]", ":", "?"]


def _nombre_hoja_control(nombre_hoja_marketing, nombres_ya_usados):
    nombre_base = str(nombre_hoja_marketing).strip() or "Sin nombre"

    for caracter in _CARACTERES_INVALIDOS_NOMBRE_HOJA:
        nombre_base = nombre_base.replace(caracter, "")

    nombre = f"Control - {nombre_base}"[:31]

    nombre_final = nombre
    contador = 2
    while nombre_final in nombres_ya_usados:
        sufijo = f" ({contador})"
        nombre_final = nombre[:31 - len(sufijo)] + sufijo
        contador += 1

    nombres_ya_usados.add(nombre_final)
    return nombre_final


def _armar_hojas_control_divididas(libro, resultado):
    nombres_ya_usados = set()

    hojas_marketing = sorted(
        resultado["Hoja"].dropna().unique(),
        key=lambda valor: str(valor).lower()
    )

    for hoja_marketing in hojas_marketing:
        datos_de_esta_hoja = resultado[resultado["Hoja"] == hoja_marketing]

        if datos_de_esta_hoja.empty:
            continue

        nombre_hoja_excel = _nombre_hoja_control(hoja_marketing, nombres_ya_usados)
        _armar_hoja_control(
            libro, 
            nombre_hoja_excel, 
            datos_de_esta_hoja, 
            titulo_personalizado=f"CONTROL DE PROMOCIONES - SECCIÓN: {str(hoja_marketing).upper()}"
        )


# ============================================================
# FUNCIONES DE ESTILO REUTILIZABLES
# ============================================================

def _configurar_impresion(hoja):
    hoja.page_setup.orientation = "landscape"
    hoja.page_setup.fitToWidth = 1
    hoja.page_setup.fitToHeight = 0
    hoja.sheet_properties.pageSetUpPr.fitToPage = True
    hoja.print_title_rows = "1:1"

    hoja.page_margins.left = 0.4
    hoja.page_margins.right = 0.4
    hoja.page_margins.top = 0.5
    hoja.page_margins.bottom = 0.5
    hoja.page_margins.header = 0.2
    hoja.page_margins.footer = 0.2


def _dar_estilo_encabezado(hoja):
    relleno = PatternFill(start_color=COLOR_ENCABEZADO, end_color=COLOR_ENCABEZADO, fill_type="solid")
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = relleno
        celda.alignment = Alignment(horizontal="center", vertical="center")


def _dar_estilo_encabezado_fila(hoja, fila_num):
    relleno = PatternFill(start_color=COLOR_ENCABEZADO, end_color=COLOR_ENCABEZADO, fill_type="solid")
    for celda in hoja[fila_num]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = relleno
        celda.alignment = Alignment(horizontal="center", vertical="center")


def _resaltar_por_estado(hoja, columna_estado):
    _resaltar_por_estado_desde_fila(hoja, columna_estado, fila_inicio=2)


def _resaltar_por_estado_desde_fila(hoja, columna_estado, fila_inicio):
    if columna_estado is None:
        return

    relleno_vencida = PatternFill(start_color=COLOR_VENCIDA, end_color=COLOR_VENCIDA, fill_type="solid")
    relleno_manana = PatternFill(start_color=COLOR_VENCE_MANANA, end_color=COLOR_VENCE_MANANA, fill_type="solid")
    relleno_proximamente = PatternFill(start_color=COLOR_PROXIMAMENTE, end_color=COLOR_PROXIMAMENTE, fill_type="solid")

    for fila in range(fila_inicio, hoja.max_row + 1):
        valor = hoja.cell(fila, columna_estado).value

        if valor == "VENCIDA":
            relleno = relleno_vencida
        elif valor == "VENCE MAÑANA":
            relleno = relleno_manana
        elif valor == "PRÓXIMAMENTE":
            relleno = relleno_proximamente
        else:
            continue

        for columna in range(1, hoja.max_column + 1):
            hoja.cell(fila, columna).fill = relleno


def _resaltar_duplicados(hoja, columna_duplicado):
    if columna_duplicado is None:
        return

    relleno = PatternFill(start_color=COLOR_DUPLICADO, end_color=COLOR_DUPLICADO, fill_type="solid")
    for fila in range(2, hoja.max_row + 1):
        if hoja.cell(fila, columna_duplicado).value == "SI":
            hoja.cell(fila, columna_duplicado).fill = relleno


def _formato_moneda(hoja, columna_precio):
    _formato_moneda_desde_fila(hoja, columna_precio, fila_inicio=2)


def _formato_moneda_desde_fila(hoja, columna_precio, fila_inicio):
    for fila in range(fila_inicio, hoja.max_row + 1):
        hoja.cell(fila, columna_precio).number_format = '$#,##0.00'
