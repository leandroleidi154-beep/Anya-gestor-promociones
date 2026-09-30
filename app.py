import streamlit as st
import pandas as pd
import requests
import os

# Importar funciones de los scripts auxiliares
from leer_stock import leer_stock
from leer_promociones import leer_promociones

# ---------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA E IMÁGENES
# ---------------------------------------------------------
st.set_page_config(
    page_title="Generador de Reportes de Promociones",
    page_icon="🐾",
    layout="wide"
)

# Enlace de Google Drive / Sheets para las promociones de Octubre
URL_GOOGLE_DRIVE = "https://docs.google.com/spreadsheets/d/1jbd2kIlZ9MJaatJt2xPygj8MQxpXdELK/export?format=xlsx"

# ---------------------------------------------------------
# FUNCIÓN PARA DESCARGAR EL ARCHIVO DE GOOGLE DRIVE
# ---------------------------------------------------------
def descargar_excel_promociones(url):
    """
    Descarga el archivo Excel desde Google Drive/Sheets evitando que devuelva
    páginas HTML de aviso/confirmación.
    """
    ruta_destino = "temp_promociones_octubre.xlsx"
    try:
        session = requests.Session()
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
        }
        
        response = session.get(url, headers=headers, allow_redirects=True, stream=True)
        
        # Guardar contenido en disco
        with open(ruta_destino, "wb") as f:
            for chunk in response.iter_content(chunk_size=32768):
                if chunk:
                    f.write(chunk)
                    
        return ruta_destino
    except Exception as e:
        st.error(f"Error al descargar promociones desde Google Drive: {e}")
        return None


# ---------------------------------------------------------
# INTERFAZ PRINCIPAL DE STREAMLIT
# ---------------------------------------------------------
st.title("🐱 Generar reporte para Sucursal")

# Pestañas de la aplicación
tab1, tab2 = st.tabs(["🏢 Uso en Sucursal", "📢 Carga de Marketing"])

with tab1:
    st.markdown("### 1. Cargar archivo de Stock (.xlsx, .xls, .ods)")
    
    # Cargar archivo de Stock de la sucursal
    stock_file = st.file_uploader(
        "Arrastrá o seleccioná el reporte de stock exportado del sistema",
        type=["xlsx", "xls", "ods"],
        key="stock_uploader"
    )

    stock_df = None
    if stock_file is not None:
        try:
            stock_df = leer_stock(stock_file)
            st.success(f"✅ Archivo de stock cargado correctamente ({len(stock_df)} filas).")
        except Exception as e:
            st.error(f"❌ Error al procesar el archivo de stock: {e}")

    st.markdown("---")
    st.markdown("### 2. Selección de Promociones")

    # Opción para usar promociones personalizadas o el archivo por defecto de Drive
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
                promociones_df = leer_promociones(promo_file)
                st.success(f"✅ Promociones personalizadas cargadas ({len(promociones_df)} registros).")
            except Exception as e:
                st.error(f"❌ Error al leer las promociones personalizadas: {e}")
    else:
        # Descarga automática desde la URL oficial de Google Drive
        with st.spinner("Descargando archivo de promociones oficial de Octubre..."):
            ruta_temp_promo = descargar_excel_promociones(URL_GOOGLE_DRIVE)
            
        if ruta_temp_promo and os.path.exists(ruta_temp_promo):
            try:
                promociones_df = leer_promociones(ruta_temp_promo)
                if promociones_df is not None and not promociones_df.empty:
                    st.info("ℹ️ Utilizando el archivo de promociones oficial de Octubre desde Google Drive.")
                else:
                    st.warning("⚠️ El archivo de promociones bajó, pero no se reconoció su estructura interna o nombres de hojas. Verificá los permisos del archivo en Google Drive.")
            except Exception as e:
                st.warning("⚠️ No se pudo procesar el archivo de promociones de Google Drive. Asegurate de que esté público para 'Cualquier persona con el enlace'.")

    # ---------------------------------------------------------
    # 3. FILTROS Y PROCESAMIENTO FINAL (SE MUESTRA SI HAY PROMOS Y STOCK)
    # ---------------------------------------------------------
    if stock_df is not None and promociones_df is not None and not promociones_df.empty:
        st.markdown("---")
        st.markdown("### 3. Opciones de Filtro y Generación")

        # Imagen de la gata si existe localmente o enlace directo
        if os.path.exists("gata.png"):
            st.image("gata.png", width=150)

        # Controles de filtro (Laboratorio, Marca, etc. según la estructura de tu archivo)
        col1, col2 = st.columns(2)
        
        with col1:
            laboratorios_disponibles = sorted(promociones_df["Laboratorio"].dropna().unique()) if "Laboratorio" in promociones_df.columns else []
            lab_seleccionados = st.multiselect("Filtrar por Laboratorio (opcional):", laboratorios_disponibles)

        with col2:
            marcas_disponibles = sorted(promociones_df["Marca"].dropna().unique()) if "Marca" in promociones_df.columns else []
            marcas_seleccionadas = st.multiselect("Filtrar por Marca (opcional):", marcas_disponibles)

        # Aplicar filtros si se seleccionaron
        promos_filtradas = promociones_df.copy()
        if lab_seleccionados:
            promos_filtradas = promos_filtradas[promos_filtradas["Laboratorio"].isin(lab_seleccionados)]
        if marcas_seleccionadas:
            promos_filtradas = promos_filtradas[promos_filtradas["Marca"].isin(marcas_seleccionadas)]

        # Botón para generar el reporte
        if st.button("🚀 Procesar y Generar Reporte Final", type="primary"):
            try:
                # Cruce de datos entre Stock y Promociones
                # Se asume que ambas tablas tienen una columna común (ej. Código o EAN)
                col_cruce = "EAN" if "EAN" in stock_df.columns else stock_df.columns[0]
                
                reporte_final = pd.merge(
                    stock_df,
                    promos_filtradas,
                    on=col_cruce,
                    how="inner"
                )
                
                st.success(f"🎉 ¡Reporte generado con éxito! Se encontraron {len(reporte_final)} coincidencias.")
                
                # Vista previa
                st.dataframe(reporte_final.head(20))

                # Exportar a Excel
                output_path = "Reporte_Promociones_Sucursal.xlsx"
                with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                    reporte_final.to_excel(writer, index=False, sheet_name="Promociones")

                # Botón de descarga para el usuario
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 Descargar Reporte en Excel",
                        data=f,
                        file_name="Reporte_Promociones_Sucursal.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"❌ Ocurrió un error al realizar el cruce de información: {e}")

with tab2:
    st.subheader("Carga y administración para el equipo de Marketing")
    st.write("Esta sección está reservada para actualizar la base de datos principal de promociones.")
