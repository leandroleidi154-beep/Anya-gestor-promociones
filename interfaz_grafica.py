"""
Interfaz gráfica del Gestor de Promociones.

Esta ventana le permite a cualquier persona de la farmacia:
1. Elegir el archivo de promociones (el que manda marketing, con varias hojas)
2. Elegir el archivo de stock (Excel o LibreOffice .ods)
3. Apretar un botón y generar el Excel final, sin tocar nada de código.
"""

import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from leer_promociones import leer_promociones
from leer_stock import leer_stock
from cruzar import cruzar_promociones_stock
from formato import generar_excel


class Aplicacion:

    def __init__(self, raiz):

        self.raiz = raiz
        self.raiz.title("Gestor de Promociones - Farmacia")
        self.raiz.geometry("560x320")
        self.raiz.resizable(False, False)

        self.ruta_promociones = tk.StringVar()
        self.ruta_stock = tk.StringVar()
        self.texto_estado = tk.StringVar(value="Elegí los dos archivos para empezar.")

        self.armar_ventana()

    # ============================================================
    # ARMADO DE LA VENTANA
    # ============================================================

    def armar_ventana(self):

        contenedor = tk.Frame(self.raiz, padx=20, pady=20)
        contenedor.pack(fill="both", expand=True)

        titulo = tk.Label(
            contenedor,
            text="Cruce de Promociones con Stock",
            font=("Segoe UI", 14, "bold")
        )
        titulo.pack(pady=(0, 15))

        # --- Archivo de promociones ---
        self.armar_selector(
            contenedor,
            etiqueta="1. Archivo de Promociones (Excel de marketing)",
            variable=self.ruta_promociones,
            comando=self.elegir_promociones
        )

        # --- Archivo de stock ---
        self.armar_selector(
            contenedor,
            etiqueta="2. Archivo de Stock (Excel o LibreOffice .ods)",
            variable=self.ruta_stock,
            comando=self.elegir_stock
        )

        # --- Botón generar ---
        boton_generar = tk.Button(
            contenedor,
            text="Generar reporte",
            font=("Segoe UI", 11, "bold"),
            bg="#2F5496",
            fg="white",
            padx=10,
            pady=8,
            command=self.generar
        )
        boton_generar.pack(pady=(20, 10), fill="x")

        # --- Barra de estado ---
        etiqueta_estado = tk.Label(
            contenedor,
            textvariable=self.texto_estado,
            wraplength=500,
            justify="left",
            fg="#444444"
        )
        etiqueta_estado.pack(pady=(5, 0), fill="x")

    def armar_selector(self, padre, etiqueta, variable, comando):

        marco = tk.LabelFrame(padre, text=etiqueta, padx=10, pady=10)
        marco.pack(fill="x", pady=6)

        fila = tk.Frame(marco)
        fila.pack(fill="x")

        campo = tk.Entry(fila, textvariable=variable, state="readonly")
        campo.pack(side="left", fill="x", expand=True, padx=(0, 8))

        boton = tk.Button(fila, text="Elegir archivo...", command=comando)
        boton.pack(side="right")

    # ============================================================
    # ELEGIR ARCHIVOS
    # ============================================================

    def elegir_promociones(self):

        ruta = filedialog.askopenfilename(
            title="Seleccioná el archivo de Promociones",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
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

            self.actualizar_estado("Cruzando promociones con el stock...")
            resultado = cruzar_promociones_stock(promociones, stock)

            if resultado.empty:
                self.actualizar_estado("No se encontraron coincidencias.")
                messagebox.showinfo(
                    "Sin coincidencias",
                    "No se encontró ninguna promoción para los productos de ese stock."
                )
                return

            # Le preguntamos al usuario dónde quiere guardar el archivo final
            archivo_salida = filedialog.asksaveasfilename(
                title="Guardar reporte de promociones como...",
                defaultextension=".xlsx",
                initialfile="Promociones_en_Stock.xlsx",
                filetypes=[("Excel", "*.xlsx")]
            )

            if not archivo_salida:
                self.actualizar_estado("Operación cancelada: no se eligió dónde guardar.")
                return

            generar_excel(resultado, archivo_salida)

            carpeta_destino = os.path.dirname(archivo_salida) or "."

            cantidad_vencidas = 0
            if "Estado" in resultado.columns:
                cantidad_vencidas = (resultado["Estado"] == "VENCIDA").sum()

            cantidad_vence_manana = 0
            if "Estado" in resultado.columns:
                cantidad_vence_manana = (resultado["Estado"] == "VENCE MAÑANA").sum()

            cantidad_duplicados = 0
            if "Revisar duplicado" in resultado.columns:
                cantidad_duplicados = (resultado["Revisar duplicado"] == "SI").sum()

            mensaje = (
                f"Se generaron {len(resultado)} promociones para productos en stock.\n\n"
                f"Vencidas: {cantidad_vencidas}\n"
                f"Vencen mañana: {cantidad_vence_manana}\n"
                f"Para revisar (duplicadas): {cantidad_duplicados}\n\n"
                f"Archivo guardado en:\n{archivo_salida}"
            )

            self.actualizar_estado("Listo. " + mensaje.splitlines()[0])
            messagebox.showinfo("Reporte generado", mensaje)

            self.abrir_carpeta(carpeta_destino)

        except Exception as error:
            self.actualizar_estado("Ocurrió un error. Ver detalle en el cartel.")
            messagebox.showerror(
                "Ocurrió un error",
                f"No se pudo generar el reporte:\n\n{error}"
            )
            print(traceback.format_exc())

    def actualizar_estado(self, texto):
        self.texto_estado.set(texto)
        self.raiz.update_idletasks()

    def abrir_carpeta(self, carpeta):
        """Abre la carpeta donde quedó el resultado (solo en Windows)."""
        try:
            os.startfile(carpeta)
        except AttributeError:
            # No estamos en Windows (ej. estamos probando en Mac/Linux); no pasa nada.
            pass


if __name__ == "__main__":
    raiz = tk.Tk()
    aplicacion = Aplicacion(raiz)
    raiz.mainloop()
