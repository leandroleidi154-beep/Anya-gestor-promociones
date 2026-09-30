import pandas as pd

def leer_stock(filepath_or_buffer):
    """
    Lee el archivo de stock en cualquier formato:
    - Excel moderno (.xlsx)
    - Excel clásico (.xls)
    - LibreOffice Calc (.ods)
    - HTML exportado desde ERP con extensión .xls
    """
    df = None
    
    # 1. Intentar leerlo como Excel tradicional o LibreOffice (.xlsx, .xls, .ods)
    try:
        df = pd.read_excel(filepath_or_buffer)
    except Exception:
        pass

    # 2. Si falló, intentar leerlo como HTML (típico de exportaciones de sistemas de farmacia)
    if df is None:
        try:
            dfs = pd.read_html(filepath_or_buffer)
            if dfs:
                df = dfs[0]
                # Si los nombres de columnas quedaron en la primera fila de datos
                if df.iloc[0].astype(str).str.contains("Descripción|Troquel|Código de Barras|Stock", case=False).any():
                    df.columns = df.iloc[0]
                    df = df[1:].reset_index(drop=True)
        except Exception as e:
            raise ValueError(f"No se pudo interpretar el formato del archivo de stock: {e}")

    if df is None or df.empty:
        raise ValueError("El archivo de stock está vacío o no se pudo procesar.")

    # Limpiar nombres de columnas (quitar espacios extra)
    df.columns = [str(col).strip() for col in df.columns]

    return df
