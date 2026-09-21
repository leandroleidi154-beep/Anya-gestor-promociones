import os
import io
import pandas as pd
import streamlit as st

from leer_promociones import leer_promociones
from leer_stock import leer_stock
from cruzar import cruzar_promociones_stock
from formato import generar_excel

# ============================================================
# CONFIGURACIÓN DE RUTAS Y ENLACES DE GOOGLE DRIVE
# ============================================================
URL_GOOGLE_DRIVE = "https://drive.google.com/uc?export=download&id=1v_cUpBdva_MXc_CfXBThxnJvY6uweMVZ"

DIRECTORIO_BASE = os.path.dirname(os.path.abspath(__file__))
CARPETA_RECURSOS = os.path.join(DIRECTORIO_BASE, "recursos")

def obtener_ruta_imagen(nombre_buscado):
    if not os.path.exists(CARPETA_RECURSOS):
        return None
    for archivo in os.listdir(CARPETA_RECURSOS):
        nombre_sin_ext, _ = os.path.splitext(archivo)
        if nombre_sin_ext.lower().startswith(nombre_buscado.lower()):
            return os.path.join(CARPETA_RECURSOS, archivo)
    return None

RUTA_LOGO = obtener_ruta_imagen("gata_logo")
RUTA_PENSANDO = obtener_ruta_imagen("gata_pensando")
RUTA_EXITO = obtener_ruta_imagen("gata_exito")
RUTA_TRISTE = obtener_ruta_imagen("gata_triste")

# Configuración de página
st.set_page_config(
    page_title="Gestor de Promociones - Farmacia del Pueblo",
    page_icon=RUTA_LOGO if RUTA_LOGO else "💊",
    layout="centered"
)

# Estilos CSS avanzados para ocultar la interfaz por defecto de Streamlit
st.markdown("""
    <style>
    /* Ocultar barra superior (Header), botones de GitHub, Fork y menú de 3 puntos */
    [data-testid="stHeader"],
    header,
    .stAppHeader {
        display: none !important;
    }
    
    /* Ocultar pie de página por defecto */
    footer {
        display: none !important;
    }
    
    /* Ocultar el botón rojo/blanco inferior de Streamlit Cloud y estado */
    [data-testid="stStatusWidget"],
    [data-testid="stViewerBadge"],
    .stStatusWidget,
    #stDecoration,
    [data-testid="stDecoration"],
    div[class*="viewerBadge"] {
        display: none !important;
        visibility: hidden !important;
    }
    
    .stAppToolbar {
        display: none !important;
    }

    /* Estilos del encabezado principal */
    .main-header { font-size: 28px; font-weight: bold; color: #39476A; }
    .sub-header { font-size: 15px; color: #888888; margin-bottom: 20px; }
    div.stButton > button:first-child {
        background-color: #39476A;
        color: white;
        border-radius: 8px;
        font-weight: bold;
    }
    div.stButton > button:first-child:hover {
        background-color: #5B6598;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

# Encabezado principal
col_logo, col_titulo = st.columns([1.2, 3.8])
with col_logo:
    if RUTA_LOGO:
        st.image(RUTA_LOGO, width=120)
    else:
        st.write("💊")

with col_titulo:
    st.markdown('<div class="main-header">Farmacia del Pueblo</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Gestor de Promociones | Cruce de listas con Stock local</div>', unsafe_allow_html=True)
    st.markdown('<div style="font-size: 13px; font-weight: 500; color: #5B6598; margin-top: -15px; margin-bottom: 20px;">Creado por: Farm. Leandro Leidi</div>', unsafe_allow_html=True)

# Pestañas
tab_sucursal, tab_marketing = st.tabs(["🏬 Uso en Sucursal", "📢 Carga de Marketing"])

# ============================================================
# PESTAÑA 1: SUCURSALES
# ============================================================
with tab_sucursal:
    st.subheader("Generar reporte para Sucursal")
    
    archivo_stock_subido = st.file_uploader(
        "1. Cargar archivo de Stock (.xlsx, .xls, .ods)",
        type=["xlsx", "xls", "ods"],
        key="uploader_stock"
    )
    
    if archivo_stock_subido is not None:
        ruta_stock_temp = "temp_stock_sucursal.xlsx"
        with open(ruta_stock_temp, "wb") as f:
            f.write(archivo_stock_subido.getbuffer())
        
# --- BUSCAR ESTA SECCIÓN DENTRO DE APP.PY Y REEMPLAZARLA ---

st.divider()
st.write("**2. Filtros opcionales**")

col_img_filtro, col_txt = st.columns([1, 2.5])
with col_img_filtro:
    if RUTA_PENSANDO:
        st.image(RUTA_PENSANDO, caption="Filtrando...", width=160)

with col_txt:
    # 1. Agregamos "Filtrar por Hoja" al selector de tipo de filtro:
    tipo_filtro = st.radio(
        "¿Cómo querés filtrar las promociones?",
        ["Mostrar todas", "Filtrar por Proveedor", "Filtrar por Línea", "Filtrar por Hoja"],
        horizontal=False
    )

promociones_filtradas = promociones_df.copy()

if tipo_filtro == "Filtrar por Proveedor":
    opciones = sorted([str(x) for x in promociones_df["Proveedor"].dropna().unique() if str(x).strip() != ""])
    sel = st.multiselect("Seleccioná los proveedores:", opciones)
    if sel:
        promociones_filtradas = promociones_df[promociones_df["Proveedor"].astype(str).isin(sel)]

elif tipo_filtro == "Filtrar por Línea":
    opciones = sorted([str(x) for x in promociones_df["Linea"].dropna().unique() if str(x).strip() != ""])
    sel = st.multiselect("Seleccioná las líneas:", opciones)
    if sel:
        promociones_filtradas = promociones_df[promociones_df["Linea"].astype(str).isin(sel)]

# 2. Nueva opción para separar/filtrar exactamente por Hoja de Marketing:
elif tipo_filtro == "Filtrar por Hoja":
    opciones = sorted([str(x) for x in promociones_df["Hoja"].dropna().unique() if str(x).strip() != ""])
    sel = st.multiselect("Seleccioná la o las hojas de Marketing a procesar:", opciones)
    if sel:
        promociones_filtradas = promociones_df[promociones_df["Hoja"].astype(str).isin(sel)]

        st.divider()
            st.write("**3. Filtros opcionales**")
            
            col_img_filtro, col_txt = st.columns([1, 2.5])
            with col_img_filtro:
                if RUTA_PENSANDO:
                    st.image(RUTA_PENSANDO, caption="Filtrando...", width=160)
            
            with col_txt:
                tipo_filtro = st.radio(
                    "¿Cómo querés filtrar las promociones?",
                    ["Mostrar todas", "Filtrar por Proveedor", "Filtrar por Línea", "Filtrar por Hoja"],
                    horizontal=False
                )
            
            promociones_filtradas = promociones_df.copy()
            
            if tipo_filtro == "Filtrar por Proveedor":
                opciones = sorted([str(x) for x in promociones_df["Proveedor"].dropna().unique() if str(x).strip() != ""])
                sel = st.multiselect("Seleccioná los proveedores:", opciones)
                if sel:
                    promociones_filtradas = promociones_df[promociones_df["Proveedor"].astype(str).isin(sel)]
            
            elif tipo_filtro == "Filtrar por Línea":
                opciones = sorted([str(x) for x in promociones_df["Linea"].dropna().unique() if str(x).strip() != ""])
                sel = st.multiselect("Seleccioná las líneas:", opciones)
                if sel:
                    promociones_filtradas = promociones_df[promociones_df["Linea"].astype(str).isin(sel)]
                    
            elif tipo_filtro == "Filtrar por Hoja":
                opciones = sorted([str(x) for x in promociones_df["Hoja"].dropna().unique() if str(x).strip() != ""])
                sel = st.multiselect("Seleccioná las hojas:", opciones)
                if sel:
                    promociones_filtradas = promociones_df[promociones_df["Hoja"].astype(str).isin(sel)]

            st.divider()
            
            if st.button("🚀 Procesar y Generar Reporte", use_container_width=True):
                with st.spinner("Cruzando promociones con el stock..."):
                    stock_df = leer_stock(ruta_stock_temp)
                    resultado = cruzar_promociones_stock(promociones_filtradas, stock_df)
                
                if resultado.empty:
                    col_triste_txt, col_triste_img = st.columns([2.5, 1])
                    with col_triste_txt:
                        st.error("No se encontraron coincidencias para los productos de tu stock con la selección realizada.")
                    with col_triste_img:
                        if RUTA_TRISTE:
                            st.image(RUTA_TRISTE, width=160)
                else:
                    col_exito_txt, col_exito_img = st.columns([2.5, 1])
                    with col_exito_txt:
                        st.success("¡Cruce realizado con éxito!")
                        
                        col1, col2, col3, col4 = st.columns(4)
                        col1.metric("Promociones", len(resultado))
                        
                        vencidas = (resultado["Estado"] == "VENCIDA").sum() if "Estado" in resultado.columns else 0
                        vence_manana = (resultado["Estado"] == "VENCE MAÑANA").sum() if "Estado" in resultado.columns else 0
                        duplicados = (resultado["Revisar duplicado"] == "SI").sum() if "Revisar duplicado" in resultado.columns else 0
                        
                        col2.metric("Vencidas", vencidas)
                        col3.metric("Vencen Mañana", vence_manana)
                        col4.metric("Duplicadas", duplicados)
                        
                    with col_exito_img:
                        if RUTA_EXITO:
                            st.image(RUTA_EXITO, caption="¡Promos encontradas!", width=160)
                    
                    output = io.BytesIO()
                    generar_excel(resultado, output)
                    bytes_excel = output.getvalue()
                    
                    st.download_button(
                        label="📥 Descargar Reporte en Excel (.xlsx)",
                        data=bytes_excel,
                        file_name="Promociones_en_Stock.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        if os.path.exists(ruta_stock_temp):
            os.remove(ruta_stock_temp)

# ============================================================
# PESTAÑA 2: INFORMACIÓN DE MARKETING
# ============================================================
with tab_marketing:
    st.subheader("Origen de Datos Central")
    st.info("ℹ️ La aplicación lee las promociones por defecto directamente desde la carpeta de Google Drive.")
    st.write("Cada vez que el equipo de Marketing reemplace o edite este archivo en Google Drive, todas las sucursales verán la información actualizada sin necesidad de reiniciar la web.")
