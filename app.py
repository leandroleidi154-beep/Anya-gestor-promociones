import os
import io
import tempfile
import streamlit as st
import pandas as pd

from leer_promociones import leer_promociones
from leer_stock import leer_stock
from cruzar import cruzar_promociones_stock
from formato import generar_excel

# Configuración de página
st.set_page_config(page_title="Anya Gestor Promociones", page_icon="🐾", layout="wide")

# Rutas de imágenes locales (si existen en el repo)
RUTA_ESPERANDO = "imagenes/esperando.jpg" if os.path.exists("imagenes/esperando.jpg") else None
RUTA_PENSANDO = "imagenes/pensando.jpg" if os.path.exists("imagenes/pensando.jpg") else None
RUTA_EXITO = "imagenes/exito.jpg" if os.path.exists("imagenes/exito.jpg") else None
RUTA_TRISTE = "imagenes/triste.jpg" if os.path.exists("imagenes/triste.jpg") else None

# Enlace de Google Drive con las promociones centralizadas
URL_GOOGLE_DRIVE = "https://docs.google.com/spreadsheets/d/1v_cUpBdva_MXc_CfXBThxnJvY6uweMVZ/export?format=xlsx"

st.title("🐾 Anya - Gestor de Promociones")
st.write("Sube el stock de tu sucursal para cruzarlo con el archivo de promociones central de Marketing.")

st.divider()
st.write("**1. Cargar Stock de la Sucursal**")

archivo_stock = st.file_uploader("Selecciona el archivo de Stock (Excel / TXT / CSV)", type=["xlsx", "xls", "txt", "csv", "ods"])

if archivo_stock is not None:
    # Guardar temporalmente el archivo subido
    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(archivo_stock.name)[1]) as tmp_file:
        tmp_file.write(archivo_stock.getvalue())
        ruta_stock_temp = tmp_file.name

    try:
        with st.spinner("Descargando base de promociones desde Google Drive..."):
            promociones_df = leer_promociones(URL_GOOGLE_DRIVE)

        st.divider()
        st.write("**2. Filtros opcionales**")

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
            sel = st.multiselect("Seleccioná las hojas de Marketing:", opciones)
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

    except Exception as e:
        st.error(f"⚠️ No se pudo obtener la lista de promociones desde Google Drive. Asegúrese de que el archivo tenga permisos de lectura abiertos. Detalles: {e}")

    finally:
        # Limpieza del archivo temporal
        if os.path.exists(ruta_stock_temp):
            os.remove(ruta_stock_temp)
