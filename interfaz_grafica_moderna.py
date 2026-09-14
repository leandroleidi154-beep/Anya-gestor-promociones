"""
Interfaz gráfica del Gestor de Promociones - Farmacia del Pueblo

Versión con CustomTkinter para un aspecto más moderno, con el logo
y los colores institucionales de la farmacia.

IMPORTANTE: esta versión necesita dos librerías extra que la versión
anterior no necesitaba:

    pip install customtkinter pillow

(pillow puede que ya la tengas instalada, la usa mucha gente).
"""

import os
import sys
import traceback
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image

from leer_promociones import leer_promociones
from leer_stock import leer_stock
from cruzar import cruzar_promociones_stock, filtrar_por_periodo
from formato import generar_excel


# ============================================================
# COLORES INSTITUCIONALES - Farmacia del Pueblo
# (sacados del logo: amarillo del círculo y azul del texto)
# ============================================================

AMARILLO = "#FDD732"
AMARILLO_HOVER = "#E8C22A"
AZUL = "#39476A"
AZUL_CLARO = "#5B6598"
GRIS_TEXTO = "#666666"
BLANCO = "#FFFFFF"


def ruta_recurso(nombre_archivo):
    """
    Devuelve la ruta a un archivo dentro de la carpeta 'recursos',
    tanto si el programa corre normal (python interfaz_grafica.py)
    como si ya está empaquetado como .exe con PyInstaller.
    """
    if hasattr(sys, "_MEIPASS"):
        # Estamos corriendo como .exe empaquetado
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base, "recursos", nombre_archivo)


class VentanaAviso(ctk.CTkToplevel):
    """
    Ventana de aviso simple, con una imagen, un título, un mensaje
    y un botón "Aceptar". Se usa en vez de los cartelitos comunes de
    Windows para poder mostrar una foto junto al mensaje.
    """

    def __init__(self, padre, titulo, mensaje, nombre_imagen):
        super().__init__(padre)

        self.title(titulo)
        self.resizable(False, False)
        self.configure(fg_color=BLANCO)

        self.transient(padre)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        contenedor = ctk.CTkFrame(self, fg_color=BLANCO)
        contenedor.pack(fill="both", expand=True, padx=25, pady=25)

        try:
            imagen = Image.open(ruta_recurso(nombre_imagen))
            imagen_ctk = ctk.CTkImage(light_image=imagen, size=(130, 130))
            etiqueta_imagen = ctk.CTkLabel(contenedor, image=imagen_ctk, text="")
            etiqueta_imagen.pack(pady=(0, 15))
        except Exception:
            pass

        etiqueta_titulo = ctk.CTkLabel(
            contenedor,
            text=titulo,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=AZUL
        )
        etiqueta_titulo.pack(pady=(0, 8))

        etiqueta_mensaje = ctk.CTkLabel(
            contenedor,
            text=mensaje,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#444444",
            wraplength=340,
            justify="center"
        )
        etiqueta_mensaje.pack(pady=(0, 15))

        boton_aceptar = ctk.CTkButton(
            contenedor,
            text="Aceptar",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=AZUL,
            hover_color=AZUL_CLARO,
            height=36,
            corner_radius=10,
            command=self.destroy
        )
        boton_aceptar.pack(fill="x")

        # Ajustamos el tamaño de la ventana al contenido real, en vez de
        # un tamaño fijo, para que nunca corte texto ni el botón por más
        # largo que sea el mensaje.
        self.update_idletasks()
        ancho = max(420, contenedor.winfo_reqwidth() + 50)
        alto = contenedor.winfo_reqheight() + 50
        self.geometry(f"{ancho}x{alto}")


class VentanaFiltro(ctk.CTkToplevel):
    """
    Ventana emergente que aparece después de cruzar los datos.
    Le permite al usuario elegir si quiere ver TODAS las promociones,
    o filtrar por Proveedor o por Línea (marca), tildando una o
    varias opciones.

    El resultado final queda guardado en self.resultado_filtrado:
    - Un DataFrame con las filas elegidas, si el usuario confirmó.
    - None, si cerró la ventana sin elegir nada (se interpreta como
      "cancelar todo el proceso").
    """

    def __init__(self, padre, resultado):
        super().__init__(padre)

        self.resultado_original = resultado
        self.resultado_filtrado = None
        self.casillas = {}

        # Filtro de período: es independiente de las 4 opciones de
        # arriba (Todas/Proveedor/Línea/Hoja) y se puede combinar con
        # cualquiera de ellas. Se guarda acá (y no solo en un widget)
        # para que no se pierda si el usuario pasa al paso 2 y vuelve.
        self.usar_periodo = ctk.BooleanVar(value=False)
        self.texto_desde = ctk.StringVar(value="")
        self.texto_hasta = ctk.StringVar(value="")

        self.title("Filtrar promociones")
        self.geometry("640x520")
        self.resizable(False, False)
        self.configure(fg_color=BLANCO)

        # La hacemos modal: bloquea la ventana principal hasta que se cierre.
        self.transient(padre)
        self.grab_set()

        self.protocol("WM_DELETE_WINDOW", self.cancelar)

        self.armar_paso_1()

    def limpiar_ventana(self):
        for widget in self.winfo_children():
            widget.destroy()

    # ============================================================
    # PASO 1: elegir cómo filtrar
    # ============================================================

    def armar_paso_1(self):

        self.limpiar_ventana()

        contenedor = ctk.CTkFrame(self, fg_color=BLANCO)
        contenedor.pack(fill="both", expand=True, padx=25, pady=25)

        # Columna izquierda: imagen decorativa
        columna_imagen = ctk.CTkFrame(contenedor, fg_color=BLANCO)
        columna_imagen.pack(side="left", fill="y", padx=(0, 20))

        try:
            imagen_filtro = Image.open(ruta_recurso("filtro.jpg"))
            imagen_ctk = ctk.CTkImage(light_image=imagen_filtro, size=(220, 220))
            etiqueta_imagen = ctk.CTkLabel(columna_imagen, image=imagen_ctk, text="")
            etiqueta_imagen.pack(pady=(10, 0))
        except Exception:
            pass

        # Columna derecha: título y botones
        columna_botones = ctk.CTkFrame(contenedor, fg_color=BLANCO)
        columna_botones.pack(side="left", fill="both", expand=True)

        titulo = ctk.CTkLabel(
            columna_botones,
            text="¿Qué promociones\nquerés incluir?",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color=AZUL,
            justify="left",
            anchor="w"
        )
        titulo.pack(fill="x", pady=(0, 20))

        boton_todas = ctk.CTkButton(
            columna_botones,
            text="Mostrar todas las promociones",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=AZUL,
            hover_color=AZUL_CLARO,
            height=42,
            corner_radius=10,
            command=self.elegir_todas
        )
        boton_todas.pack(fill="x", pady=8)

        boton_proveedor = ctk.CTkButton(
            columna_botones,
            text="Filtrar por Proveedor...",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=AMARILLO,
            hover_color=AMARILLO_HOVER,
            text_color=AZUL,
            height=42,
            corner_radius=10,
            command=lambda: self.armar_paso_2("Proveedor")
        )
        boton_proveedor.pack(fill="x", pady=8)

        boton_linea = ctk.CTkButton(
            columna_botones,
            text="Filtrar por Línea...",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=AMARILLO,
            hover_color=AMARILLO_HOVER,
            text_color=AZUL,
            height=42,
            corner_radius=10,
            command=lambda: self.armar_paso_2("Linea")
        )
        boton_linea.pack(fill="x", pady=8)

        boton_hoja = ctk.CTkButton(
            columna_botones,
            text="Filtrar por Hoja del Excel...",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=AMARILLO,
            hover_color=AMARILLO_HOVER,
            text_color=AZUL,
            height=42,
            corner_radius=10,
            command=lambda: self.armar_paso_2("Hoja")
        )
        boton_hoja.pack(fill="x", pady=8)

        self.armar_bloque_periodo(columna_botones)

    def armar_bloque_periodo(self, padre):
        """
        Casillero "Filtrar por período", destildado por defecto, con
        dos campos opcionales (Desde/Hasta). Es independiente de las
        4 opciones de arriba: se puede combinar con "Mostrar todas",
        con "Proveedor", con "Línea" o con "Hoja".
        """

        marco = ctk.CTkFrame(padre, fg_color="#F5F5F5", corner_radius=10)
        marco.pack(fill="x", pady=(14, 0))

        casillero = ctk.CTkCheckBox(
            marco,
            text="Filtrar por período (fecha de la promoción)",
            variable=self.usar_periodo,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=AZUL,
            hover_color=AZUL_CLARO,
            text_color=AZUL
        )
        casillero.pack(anchor="w", padx=12, pady=(10, 6))

        fila_fechas = ctk.CTkFrame(marco, fg_color="#F5F5F5")
        fila_fechas.pack(fill="x", padx=12, pady=(0, 12))

        etiqueta_desde = ctk.CTkLabel(
            fila_fechas, text="Desde:",
            font=ctk.CTkFont(family="Segoe UI", size=12), text_color=GRIS_TEXTO
        )
        etiqueta_desde.pack(side="left")

        campo_desde = ctk.CTkEntry(
            fila_fechas, textvariable=self.texto_desde,
            placeholder_text="dd/mm/aaaa", width=100, height=30
        )
        campo_desde.pack(side="left", padx=(6, 16))

        etiqueta_hasta = ctk.CTkLabel(
            fila_fechas, text="Hasta:",
            font=ctk.CTkFont(family="Segoe UI", size=12), text_color=GRIS_TEXTO
        )
        etiqueta_hasta.pack(side="left")

        campo_hasta = ctk.CTkEntry(
            fila_fechas, textvariable=self.texto_hasta,
            placeholder_text="dd/mm/aaaa", width=100, height=30
        )
        campo_hasta.pack(side="left", padx=(6, 0))

    # ============================================================
    # FILTRO DE PERÍODO (combinable con cualquiera de las 4 opciones)
    # ============================================================

    def _validar_y_obtener_periodo(self):
        """
        Si el casillero de período está tildado, valida las fechas
        escritas (si hay alguna) y devuelve una tupla (desde, hasta)
        con objetos date o None. Si algo está mal escrito, avisa con
        un cartel y devuelve None para que no se siga adelante.
        Si el casillero no está tildado, devuelve (None, None) directo
        (sin filtrar nada por fecha).
        """

        if not self.usar_periodo.get():
            return (None, None)

        texto_desde = self.texto_desde.get().strip()
        texto_hasta = self.texto_hasta.get().strip()

        fecha_desde = None
        fecha_hasta = None

        if texto_desde:
            try:
                fecha_desde = datetime.strptime(texto_desde, "%d/%m/%Y").date()
            except ValueError:
                messagebox.showwarning(
                    "Fecha inválida",
                    "La fecha 'Desde' tiene que tener el formato dd/mm/aaaa."
                )
                return None

        if texto_hasta:
            try:
                fecha_hasta = datetime.strptime(texto_hasta, "%d/%m/%Y").date()
            except ValueError:
                messagebox.showwarning(
                    "Fecha inválida",
                    "La fecha 'Hasta' tiene que tener el formato dd/mm/aaaa."
                )
                return None

        return (fecha_desde, fecha_hasta)

    # ============================================================
    # "MOSTRAR TODAS": no filtra nada (salvo el período, si está tildado)
    # ============================================================

    def elegir_todas(self):

        periodo = self._validar_y_obtener_periodo()
        if periodo is None:
            return

        fecha_desde, fecha_hasta = periodo
        self.resultado_filtrado = filtrar_por_periodo(
            self.resultado_original, fecha_desde, fecha_hasta
        )
        self.destroy()

    # ============================================================
    # PASO 2: elegir valores puntuales (con casilleros)
    # ============================================================

    def armar_paso_2(self, columna):

        self.limpiar_ventana()
        self.casillas = {}
        self.casillas_widgets = {}

        etiquetas_columna = {
            "Proveedor": "Proveedor",
            "Linea": "Línea",
            "Hoja": "Hoja del Excel"
        }
        etiqueta_columna = etiquetas_columna.get(columna, columna)

        valores = sorted(
            set(
                str(valor).strip()
                for valor in self.resultado_original[columna].dropna().unique()
                if str(valor).strip() != ""
            ),
            key=lambda v: v.lower()
        )

        contenedor = ctk.CTkFrame(self, fg_color=BLANCO)
        contenedor.pack(fill="both", expand=True, padx=20, pady=20)

        titulo = ctk.CTkLabel(
            contenedor,
            text=f"Elegí uno o varios de {etiqueta_columna}",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=AZUL,
            wraplength=380
        )
        titulo.pack(pady=(0, 10))

        self.texto_busqueda = ctk.StringVar()
        self.texto_busqueda.trace_add("write", lambda *args: self.filtrar_lista())

        campo_busqueda = ctk.CTkEntry(
            contenedor,
            textvariable=self.texto_busqueda,
            placeholder_text=f"Buscar {etiqueta_columna.lower()}...",
            height=34,
            fg_color=BLANCO,
            border_color="#DDDDDD"
        )
        campo_busqueda.pack(fill="x", pady=(0, 10))
        self.campo_busqueda = campo_busqueda

        fila_botones_rapidos = ctk.CTkFrame(contenedor, fg_color=BLANCO)
        fila_botones_rapidos.pack(fill="x", pady=(0, 10))

        boton_todos = ctk.CTkButton(
            fila_botones_rapidos,
            text="Seleccionar todos",
            width=150,
            height=30,
            fg_color="#E8E8E8",
            text_color=AZUL,
            hover_color="#D5D5D5",
            command=lambda: self.marcar_todas(True)
        )
        boton_todos.pack(side="left", padx=(0, 8))

        boton_ninguno = ctk.CTkButton(
            fila_botones_rapidos,
            text="Ninguno",
            width=100,
            height=30,
            fg_color="#E8E8E8",
            text_color=AZUL,
            hover_color="#D5D5D5",
            command=lambda: self.marcar_todas(False)
        )
        boton_ninguno.pack(side="left")

        self.aviso_sin_resultados = ctk.CTkLabel(
            contenedor,
            text="No hay coincidencias con esa búsqueda.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#999999"
        )
        # Se muestra solo cuando el filtro de búsqueda no encuentra nada.

        self.lista_scroll = ctk.CTkScrollableFrame(
            contenedor,
            fg_color="#F5F5F5",
            height=250
        )
        self.lista_scroll.pack(fill="both", expand=True, pady=(0, 15))

        for i, valor in enumerate(valores):
            variable = ctk.BooleanVar(value=False)
            casilla = ctk.CTkCheckBox(
                self.lista_scroll,
                text=str(valor),
                variable=variable,
                font=ctk.CTkFont(family="Segoe UI", size=12),
                fg_color=AZUL,
                hover_color=AZUL_CLARO
            )
            casilla.grid(row=i, column=0, sticky="w", pady=4, padx=5)
            self.casillas[valor] = variable
            self.casillas_widgets[valor] = casilla

        fila_final = ctk.CTkFrame(contenedor, fg_color=BLANCO)
        fila_final.pack(fill="x")

        boton_volver = ctk.CTkButton(
            fila_final,
            text="← Volver",
            width=100,
            height=36,
            fg_color="#E8E8E8",
            text_color=AZUL,
            hover_color="#D5D5D5",
            command=self.armar_paso_1
        )
        boton_volver.pack(side="left")

        boton_continuar = ctk.CTkButton(
            fila_final,
            text="Continuar",
            height=36,
            fg_color=AZUL,
            hover_color=AZUL_CLARO,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            command=lambda: self.confirmar_seleccion(columna)
        )
        boton_continuar.pack(side="right", fill="x", expand=True, padx=(10, 0))

    def filtrar_lista(self):
        """
        Se ejecuta cada vez que el usuario escribe en el buscador.
        Muestra solo los casilleros cuyo texto contiene lo escrito
        (sin importar mayúsculas/minúsculas), y oculta el resto sin
        perder si estaban tildados o no.
        """
        texto = self.texto_busqueda.get().strip().lower()

        hay_alguno_visible = False

        for valor, casilla in self.casillas_widgets.items():
            if texto in valor.lower():
                casilla.grid()
                hay_alguno_visible = True
            else:
                casilla.grid_remove()

        if hay_alguno_visible:
            self.aviso_sin_resultados.pack_forget()
        else:
            self.aviso_sin_resultados.pack(pady=10, after=self.campo_busqueda)

    def marcar_todas(self, valor):
        """
        Tilda o destilda, pero solo las opciones que están visibles
        en este momento (si hay un texto de búsqueda escrito, respeta
        ese filtro en vez de tocar las que están ocultas).
        """
        for nombre, variable in self.casillas.items():
            casilla = self.casillas_widgets[nombre]
            if casilla.winfo_ismapped():
                variable.set(valor)

    def confirmar_seleccion(self, columna):

        seleccionados = [
            valor for valor, variable in self.casillas.items()
            if variable.get()
        ]

        if not seleccionados:
            messagebox.showwarning(
                "Nada seleccionado",
                "Tildá al menos una opción, o volvé atrás y elegí 'Mostrar todas'."
            )
            return

        periodo = self._validar_y_obtener_periodo()
        if periodo is None:
            return

        filtrado = self.resultado_original[
            self.resultado_original[columna]
            .apply(lambda valor: str(valor).strip())
            .isin(seleccionados)
        ]

        fecha_desde, fecha_hasta = periodo
        self.resultado_filtrado = filtrar_por_periodo(filtrado, fecha_desde, fecha_hasta)
        self.destroy()

    def cancelar(self):
        self.resultado_filtrado = None
        self.destroy()


class VentanaResumenControl(ctk.CTkToplevel):
    """
    Ventana que aparece después de cruzar los datos y ANTES de generar
    el archivo final. Muestra el total de coincidencias encontradas
    (vencidas, vencen mañana, duplicadas) y le pregunta al usuario si
    quiere una única hoja "Control" con todo junto, o una hoja
    "Control" separada por cada hoja del Excel de marketing.

    El resultado queda en self.continuar (True/False) y, si continuó,
    en self.dividir_control (True = varias hojas, False = una sola).
    """

    def __init__(self, padre, resultado):
        super().__init__(padre)

        self.continuar = False
        self.dividir_control = False

        self.title("Coincidencias encontradas")
        self.resizable(False, False)
        self.configure(fg_color=BLANCO)

        self.transient(padre)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.cancelar)

        self.opcion_control = ctk.StringVar(value="unica")

        contenedor = ctk.CTkFrame(self, fg_color=BLANCO)
        contenedor.pack(fill="both", expand=True, padx=25, pady=25)

        titulo = ctk.CTkLabel(
            contenedor,
            text="Coincidencias encontradas",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color=AZUL
        )
        titulo.pack(pady=(0, 15))

        cantidad_vencidas = (resultado["Estado"] == "VENCIDA").sum() if "Estado" in resultado.columns else 0
        cantidad_vence_manana = (resultado["Estado"] == "VENCE MAÑANA").sum() if "Estado" in resultado.columns else 0
        cantidad_duplicados = (resultado["Revisar duplicado"] == "SI").sum() if "Revisar duplicado" in resultado.columns else 0

        texto_resumen = (
            f"Se encontraron {len(resultado)} promociones para productos en stock.\n\n"
            f"Vencidas: {cantidad_vencidas}\n"
            f"Vencen mañana: {cantidad_vence_manana}\n"
            f"Para revisar (duplicadas): {cantidad_duplicados}"
        )

        etiqueta_resumen = ctk.CTkLabel(
            contenedor,
            text=texto_resumen,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#444444",
            justify="left"
        )
        etiqueta_resumen.pack(fill="x", pady=(0, 20))

        subtitulo = ctk.CTkLabel(
            contenedor,
            text="¿Cómo querés armar la hoja de Control (la que se imprime)?",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=AZUL,
            justify="left",
            wraplength=380
        )
        subtitulo.pack(fill="x", pady=(0, 10))

        opcion_unica = ctk.CTkRadioButton(
            contenedor,
            text="Una sola hoja de Control, con todas las promociones juntas",
            variable=self.opcion_control,
            value="unica",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=AZUL,
            hover_color=AZUL_CLARO
        )
        opcion_unica.pack(anchor="w", pady=4)

        opcion_dividida = ctk.CTkRadioButton(
            contenedor,
            text="Varias hojas de Control, una por cada hoja del Excel de marketing",
            variable=self.opcion_control,
            value="dividida",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=AZUL,
            hover_color=AZUL_CLARO
        )
        opcion_dividida.pack(anchor="w", pady=4)

        fila_botones = ctk.CTkFrame(contenedor, fg_color=BLANCO)
        fila_botones.pack(fill="x", pady=(20, 0))

        boton_cancelar = ctk.CTkButton(
            fila_botones,
            text="Cancelar",
            width=100,
            height=36,
            fg_color="#E8E8E8",
            text_color=AZUL,
            hover_color="#D5D5D5",
            command=self.cancelar
        )
        boton_cancelar.pack(side="left")

        boton_continuar = ctk.CTkButton(
            fila_botones,
            text="Continuar",
            height=36,
            fg_color=AZUL,
            hover_color=AZUL_CLARO,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            command=self.confirmar
        )
        boton_continuar.pack(side="right", fill="x", expand=True, padx=(10, 0))

        self.update_idletasks()
        ancho = max(440, contenedor.winfo_reqwidth() + 50)
        alto = contenedor.winfo_reqheight() + 50
        self.geometry(f"{ancho}x{alto}")

    def confirmar(self):
        self.continuar = True
        self.dividir_control = (self.opcion_control.get() == "dividida")
        self.destroy()

    def cancelar(self):
        self.continuar = False
        self.destroy()


class Aplicacion(ctk.CTk):

    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("light")

        self.title("Gestor de Promociones - Farmacia del Pueblo")
        self.geometry("600x480")
        self.resizable(False, False)
        self.configure(fg_color=BLANCO)

        self.aplicar_icono()

        self.ruta_promociones = ctk.StringVar()
        self.ruta_stock = ctk.StringVar()
        self.texto_estado = ctk.StringVar(value="Elegí los dos archivos para empezar.")

        self.armar_ventana()

    def aplicar_icono(self):
        """
        Le pone el logo de la farmacia como ícono de la ventana
        (se ve en la barra de título y en la barra de tareas de Windows
        mientras el programa está abierto).
        """
        try:
            self.iconbitmap(ruta_recurso("logo.ico"))
        except Exception:
            # Si por algún motivo no encuentra el ícono, seguimos
            # sin romper la app (por ejemplo, en Mac/Linux .ico no aplica).
            pass

    # ============================================================
    # ARMADO DE LA VENTANA
    # ============================================================

    def armar_ventana(self):

        contenedor = ctk.CTkFrame(self, fg_color=BLANCO)
        contenedor.pack(fill="both", expand=True, padx=25, pady=20)

        self.armar_encabezado(contenedor)

        self.armar_selector(
            contenedor,
            etiqueta="1.  Archivo de Promociones (Excel de marketing)",
            variable=self.ruta_promociones,
            comando=self.elegir_promociones
        )

        self.armar_selector(
            contenedor,
            etiqueta="2.  Archivo de Stock (Excel o LibreOffice .ods)",
            variable=self.ruta_stock,
            comando=self.elegir_stock
        )

        boton_generar = ctk.CTkButton(
            contenedor,
            text="Generar reporte",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color=AZUL,
            hover_color=AZUL_CLARO,
            text_color=BLANCO,
            height=42,
            corner_radius=10,
            command=self.generar
        )
        boton_generar.pack(pady=(25, 10), fill="x")

        etiqueta_estado = ctk.CTkLabel(
            contenedor,
            textvariable=self.texto_estado,
            wraplength=520,
            justify="left",
            text_color=GRIS_TEXTO,
            font=ctk.CTkFont(family="Segoe UI", size=12)
        )
        etiqueta_estado.pack(pady=(5, 0), fill="x")

        self.armar_pie_de_pagina(contenedor)

    def armar_encabezado(self, padre):

        marco = ctk.CTkFrame(padre, fg_color=BLANCO)
        marco.pack(fill="x", pady=(0, 20))

        # Logo (si no se encuentra el archivo, seguimos sin romper la app)
        try:
            imagen_logo = Image.open(ruta_recurso("logo.jpg"))
            logo_ctk = ctk.CTkImage(light_image=imagen_logo, size=(60, 60))

            etiqueta_logo = ctk.CTkLabel(marco, image=logo_ctk, text="")
            etiqueta_logo.pack(side="left", padx=(0, 15))
        except Exception:
            pass

        marco_titulos = ctk.CTkFrame(marco, fg_color=BLANCO)
        marco_titulos.pack(side="left", fill="both", expand=True)

        titulo = ctk.CTkLabel(
            marco_titulos,
            text="Farmacia del Pueblo",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color=AZUL,
            anchor="w"
        )
        titulo.pack(fill="x")

        subtitulo = ctk.CTkLabel(
            marco_titulos,
            text="Gestor de Promociones",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color=GRIS_TEXTO,
            anchor="w"
        )
        subtitulo.pack(fill="x")

    def armar_selector(self, padre, etiqueta, variable, comando):

        marco = ctk.CTkFrame(padre, fg_color="#F5F5F5", corner_radius=10)
        marco.pack(fill="x", pady=8)

        etiqueta_texto = ctk.CTkLabel(
            marco,
            text=etiqueta,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=AZUL,
            anchor="w"
        )
        etiqueta_texto.pack(fill="x", padx=15, pady=(10, 5))

        fila = ctk.CTkFrame(marco, fg_color="#F5F5F5")
        fila.pack(fill="x", padx=15, pady=(0, 12))

        campo = ctk.CTkEntry(
            fila,
            textvariable=variable,
            state="readonly",
            height=34,
            fg_color=BLANCO,
            border_color="#DDDDDD"
        )
        campo.pack(side="left", fill="x", expand=True, padx=(0, 10))

        boton = ctk.CTkButton(
            fila,
            text="Elegir archivo...",
            width=140,
            height=34,
            fg_color=AMARILLO,
            hover_color=AMARILLO_HOVER,
            text_color=AZUL,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=comando
        )
        boton.pack(side="right")

    def armar_pie_de_pagina(self, padre):

        pie = ctk.CTkLabel(
            padre,
            text="Creado por: Farm. Leandro Leidi",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color="#BBBBBB"
        )
        pie.pack(side="right", pady=(15, 0))

    # ============================================================
    # ELEGIR ARCHIVOS
    # ============================================================

    def elegir_promociones(self):

        ruta = filedialog.askopenfilename(
            title="Seleccioná el archivo de Promociones",
            filetypes=[("Excel y LibreOffice", "*.xlsx *.xls *.ods")]
        )

        if ruta:
            self.ruta_promociones.set(ruta)

    def elegir_stock(self):

        ruta = filedialog.askopenfilename(
            title="Seleccioná el archivo de Stock",
            filetypes=[("Excel y LibreOffice", "*.xlsx *.xls *.ods")]
        )

        if ruta:
            self.ruta_stock.set(ruta)

    # ============================================================
    # GENERAR EL REPORTE
    # ============================================================

    def generar(self):

        ruta_promociones = self.ruta_promociones.get()
        ruta_stock = self.ruta_stock.get()

        if not ruta_promociones or not ruta_stock:
            messagebox.showwarning(
                "Faltan archivos",
                "Elegí primero los dos archivos (Promociones y Stock)."
            )
            return

        try:
            self.actualizar_estado("Leyendo promociones... (puede tardar unos segundos)")
            promociones = leer_promociones(ruta_promociones)

            self.actualizar_estado("Leyendo stock...")
            stock = leer_stock(ruta_stock)

            # Le preguntamos al usuario si quiere ver todo, o filtrar
            # por Proveedor o por Línea. Esto se hace ANTES de cruzar con
            # el stock, para que la lista muestre TODOS los proveedores y
            # líneas del archivo de marketing, tengan o no stock cargado.
            ventana_filtro = VentanaFiltro(self, promociones)
            self.wait_window(ventana_filtro)

            promociones_filtradas = ventana_filtro.resultado_filtrado

            if promociones_filtradas is None:
                self.actualizar_estado("Operación cancelada.")
                return

            self.actualizar_estado("Cruzando promociones con el stock...")
            resultado = cruzar_promociones_stock(promociones_filtradas, stock)

            if resultado.empty:
                self.actualizar_estado("No hay productos en stock para esa selección.")
                aviso = VentanaAviso(
                    self,
                    "Sin coincidencias",
                    "No se encontró ningún producto en stock para lo que elegiste.\n\n"
                    "Puede que esa línea o proveedor no tenga productos cargados en el stock."
                    "\n\nSi tildaste el filtro por período, puede que no haya promociones "
                    "en ese rango de fechas.",
                    "error.jpg"
                )
                self.wait_window(aviso)
                return

            # Antes de generar el archivo, mostramos el total de
            # coincidencias y preguntamos cómo armar la hoja de Control.
            ventana_resumen = VentanaResumenControl(self, resultado)
            self.wait_window(ventana_resumen)

            if not ventana_resumen.continuar:
                self.actualizar_estado("Operación cancelada.")
                return

            dividir_control = ventana_resumen.dividir_control

            archivo_salida = filedialog.asksaveasfilename(
                title="Guardar reporte de promociones como...",
                defaultextension=".xlsx",
                initialfile="Promociones_en_Stock.xlsx",
                filetypes=[("Excel", "*.xlsx")]
            )

            if not archivo_salida:
                self.actualizar_estado("Operación cancelada: no se eligió dónde guardar.")
                return

            generar_excel(resultado, archivo_salida, dividir_control=dividir_control)

            self.actualizar_estado(f"Listo. Archivo guardado en {archivo_salida}")
            aviso = VentanaAviso(
                self,
                "Reporte generado",
                f"El archivo se generó correctamente.\n\nGuardado en:\n{archivo_salida}",
                "exito.jpg"
            )
            self.wait_window(aviso)

            self.abrir_archivo(archivo_salida)

        except PermissionError:
            self.actualizar_estado("El archivo está abierto en otro programa.")
            messagebox.showerror(
                "El archivo está abierto",
                "No se pudo guardar el archivo porque ya está abierto en Excel, "
                "LibreOffice, o algún otro programa.\n\n"
                "Cerralo e intentá generar el reporte de nuevo."
            )

        except Exception as error:
            self.actualizar_estado("Ocurrió un error. Ver detalle en el cartel.")
            messagebox.showerror(
                "Ocurrió un error",
                f"No se pudo generar el reporte:\n\n{error}"
            )
            print(traceback.format_exc())

    def actualizar_estado(self, texto):
        self.texto_estado.set(texto)
        self.update_idletasks()

    def abrir_archivo(self, archivo):
        """
        Abre directamente el Excel generado con el programa que
        Windows tenga asociado a .xlsx (Excel o LibreOffice Calc,
        según lo que tenga instalado cada computadora).
        """
        try:
            os.startfile(archivo)
        except AttributeError:
            # No estamos en Windows (ej. probando en Mac/Linux); no pasa nada.
            pass
        except Exception:
            # Si por algún motivo no se puede abrir (archivo bloqueado, etc.)
            # no rompemos el programa, el usuario igual sabe dónde quedó.
            pass


if __name__ == "__main__":
    app = Aplicacion()
    app.mainloop()
