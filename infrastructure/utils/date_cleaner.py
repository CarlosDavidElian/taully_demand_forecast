import pandas as pd

def clean_sales_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Limpia el DataFrame de ventas:
    - Elimina filas de resumen o anuladas en la columna PROD.
    - Convierte CANT y TOTAL a números.
    - Elimina filas con CANT nula o cero.
    - Resetea el índice.
    """
    if 'PROD' not in df.columns:
        raise ValueError("El DataFrame no contiene la columna 'PROD'")

    # 1. Filtrar únicamente etiquetas completas de resumen. No se utiliza una
    # búsqueda parcial: nombres válidos como "CUIDADO TOTAL" no son totales.
    summary_labels = r"(?:ANULADO|TOTAL|SUBTOTAL|GRUPOS)\s*:?"
    is_summary = df['PROD'].astype(str).str.strip().str.fullmatch(summary_labels, case=False, na=False)
    df = df[~is_summary].copy()

    # 2. Convertir columnas numéricas
    df['CANT'] = pd.to_numeric(df['CANT'], errors='coerce')
    df['TOTAL'] = pd.to_numeric(df['TOTAL'], errors='coerce')

    # 3. Eliminar filas con CANT nula o cero
    df = df.dropna(subset=['CANT'])
    df = df[df['CANT'] > 0]

    # 4. Resetear índice
    df = df.reset_index(drop=True)

    return df
