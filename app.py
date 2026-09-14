import io
import os
import streamlit as st

# Importaciones de tus módulos del proyecto
# (Asegúrate de ajustar los nombres si tus funciones de lectura/cruce tienen otros nombres)
from formato import generar_excel

# Configuración de la página
st.set_page_config(
    page_title="Gestor de Promociones",
    page_icon="🏷️",
    layout="wide"
)

# Estilos / Recursos de imagen opcionales
RUTA_TRISTE = "assets/gato_triste.png" if os.path.exists("assets/gato_triste.png") else None
RUTA_EXITO = "assets/gato_exito.png" if os.path.exists("assets/gato_exito.png") else None

st.title("🏷️ Gestor de Promociones")

# Pestañas principales
tab_gestor, tab_marketing = st.tabs(["📊 Gestor de Promociones", "📢 Información de Marketing"])

# ============================================================
# PESTAÑA 1: GESTOR DE PROMOCIONES
# ============================================================
with tab_gestor:
    st.header("1. Carga de Archivos y Filtros")
    
    # Ruta temporal de stock
    ruta_stock_temp = "temp_stock.xlsx"

    # --- AQUÍ VA TU LÓGICA DE CARGA DE ARCHIVOS Y FILTROS ---
    # (Asegúrate de definir 'promociones_filtradas', 'ruta_stock_temp', 'opcion_control', etc.)
    
    st.markdown("---")
    
    # --- PROCESAMIENTO Y GENERACIÓN DEL REPORTE ---
    if st.button("🚀 Procesar y Generar Reporte", use_container_width=True):
        with st.spinner("Cruzando promociones con el stock..."):
            # 1. Leer archivo de stock
            stock_df = leer_stock(ruta_stock_temp)
            
            # 2. Realizar cruce de información
            resultado = cruzar_promociones_stock(promociones_filtradas, stock_df)
        
        # Si no hay coincidencias
        if resultado.empty:
            col_triste_txt, col_triste_img = st.columns([2.5, 1])
            with col_triste_txt:
                st.error("No se encontraron coincidencias para los productos de tu stock con la selección realizada.")
            with col_triste_img:
                if RUTA_TRISTE:
                    st.image(RUTA_TRISTE, width=160)
        
        # Si se encontraron promociones
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

            # --- GENERACIÓN DEL EXCEL CON FORMATO (DENTRO DEL ELSE) ---
            output = io.BytesIO()
            
            # Determina si el usuario eligió separar hojas de control
            dividir_flag = (opcion_control == "dividida") if 'opcion_control' in locals() else False
            
            # Genera el Excel aplicando todas las reglas de formato.py
            generar_excel(resultado, output, dividir_control=dividir_flag)
            
            # Obtiene los bytes listos para descarga
            bytes_excel = output.getvalue()

            # Botón para descargar el reporte final
            st.download_button(
                label="📥 Descargar Reporte en Excel (.xlsx)",
                data=bytes_excel,
                file_name="Promociones_en_Stock.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    # Limpieza de archivos temporales
    if os.path.exists(ruta_stock_temp):
        try:
            os.remove(ruta_stock_temp)
        except Exception:
            pass

# ============================================================
# PESTAÑA 2: INFORMACIÓN DE MARKETING
# ============================================================
with tab_marketing:
    st.subheader("Origen de Datos Central")
    st.info("ℹ️ La aplicación lee las promociones por defecto directamente desde la carpeta de Google Drive.")
    st.write("Cada vez que el equipo de Marketing reemplace o edite este archivo en Google Drive, todas las sucursales verán la información actualizada sin necesidad de reiniciar la web.")
