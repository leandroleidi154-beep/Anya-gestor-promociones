import streamlit as st
import pandas as pd
import requests
import os

# Importar funciones originales del proyecto
from leer_stock import leer_stock
from leer_promociones import leer_promociones

# ---------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# ---------------------------------------------------------
st.set_page_config(
    page_title="Generador de Reportes de Promociones",
    page_icon="🐾",
    layout="wide"
)

# Enlace directo de exportación para el archivo de Octubre
URL_GOOGLE_DRIVE = "https://docs.google.com/spreadsheets/d/1jbd2kIlZ9MJaatJt2xPygj8MQxpXdELK/export?format=xlsx"

def descargar_promociones_drive(url):
    """
    Descarga el archivo desde Google Drive a un archivo temporal local.
    """
    ruta_temp = "temp_promociones_octubre.xlsx"
    try:
        session = requests.Session()
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        }
        response = session.get(url, headers=headers, allow_redirects=True)
        
        with open(ruta_temp, "wb") as f:
            f.write(response.content)
            
        return ruta_temp
    except Exception as e:
        st.error(f"Error al descargar promociones: {e}")
        return None

# ---------------------------------------------------------
# INTERFAZ
# ---------------------------------------------------------
st.title("🐱 Generar reporte para Sucursal")

tab1, tab2 = st.tabs(["🏢 Uso en Sucursal", "📢 Carga de Marketing"])

with tab1:
    st.markdown("### 1. Cargar archivo de Stock (.xlsx, .xls, .ods)")
    
    stock_file = st.file_uploader(
        "Arrastrá o seleccioná el reporte de stock exportado del sistema",
        type=["xlsx", "xls", "ods"],
        key="stock_uploader"
    )

    stock_df = None
    if stock_file is not None:
        try:
            # Solución al error 'UploadedFile': Guardar en un archivo temporal en disco
            temp_stock_path = f"temp_{stock_file.name}"
            with open(temp_stock_path, "wb") as f:
                f.write(stock_file.getbuffer())

            # Pasar la ruta física del archivo a leer_stock
            stock_df = leer_stock(temp_stock_path)
            
            # Limpiar archivo temporal de stock
            if os.path.exists(temp_stock_path):
                os.remove(temp_stock_path)

            if stock_df is not None and not stock_df.empty:
                st.success(f"✅ Archivo de stock cargado correctamente ({len(stock_df)} filas).")
            else:
                st.warning("⚠️ El archivo de stock se leyó pero no contiene datos válidos.")

        except Exception as e:
            st.error(f"❌ Error al procesar el archivo de stock: {e}")

    st.markdown("---")
    st.markdown("### 2. Selección de Promociones")

    usar_personalizado = st.checkbox("⚠️ Usar un archivo de promociones personalizado (solo para este cruce)")
    
    promociones_df = None

    if usar_personalizado:
        promo_file = st.file_uploader(
            "Cargar archivo de promociones propio (.xlsx)",
            type=["xlsx"],
            key="promo_uploader"
        )
        if promo_file is not None:
            try:
                temp_promo_custom = f"temp_custom_{promo_file.name}"
                with open(temp_promo_custom, "wb") as f:
                    f.write(promo_file.getbuffer())

                promociones_df = leer_promociones(temp_promo_custom)

                if os.path.exists(temp_promo_custom):
                    os.remove(temp_promo_custom)

                if promociones_df is not None and not promociones_df.empty:
                    st.success(f"✅ Promociones personalizadas cargadas ({len(promociones_df)} registros).")
            except Exception as e:
                st.error(f"❌ Error al leer las promociones personalizadas: {e}")
    else:
        # Descarga desde Google Drive usando archivo temporal
        ruta_promo = descargar_promociones_drive(URL_GOOGLE_DRIVE)
        
        if ruta_promo and os.path.exists(ruta_promo):
            try:
                promociones_df = leer_promociones(ruta_promo)
                if promociones_df is not None and not promociones_df.empty:
                    st.info("ℹ️ Utilizando el archivo de promociones oficial de Octubre desde Google Drive.")
                else:
                    st.warning("⚠️ El archivo de promociones bajó, pero no se reconoció su estructura interna o nombres de hojas. Verificá los permisos del archivo en Google Drive.")
            except Exception as e:
                st.warning(f"⚠️ No se pudo procesar el archivo de promociones de Google Drive: {e}")

    # ---------------------------------------------------------
    # 3. FILTROS Y PROCESAMIENTO FINAL
    # ---------------------------------------------------------
    if stock_df is not None and promociones_df is not None and not promociones_df.empty:
        st.markdown("---")
        st.markdown("### 3. Opciones de Filtro y Generación")

        # Restaurar imágenes si existen en la raíz de tu repositorio
        if os.path.exists("gata.png"):
            st.image("gata.png", width=150)

        # Si en tu código original tenías filtros o cruce de datos, continúa aquí...
        st.success("🎉 Datos listos para procesar.")

with tab2:
    st.subheader("Carga y administración para el equipo de Marketing")
    st.write("Esta sección está reservada para actualizar la base de datos principal de promociones.")
