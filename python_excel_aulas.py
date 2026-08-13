import pandas as pd
import json
import numpy as np

# Ruta del archivo Excel
excel_path = r"C:\Users\jonathan.lopez\Documents\Python\Master Web\PA horarios\Reporte_Aulas_CUGDL_2026B_13_Agosto_13_00.xlsx"

# Archivo JSON de salida
json_path = r"C:\Users\jonathan.lopez\Documents\Python\Master Web\PA horarios\aulas_2026B.json"

try:
    # Leer primera hoja
    df = pd.read_excel(excel_path, sheet_name=0, engine='openpyxl')

    # Limpiar nombres de columnas
    df.columns = [col.strip().replace(" ", "_").lower() for col in df.columns]

    # Eliminar filas completamente vacías
    df = df.dropna(how='all').reset_index(drop=True)

    # Detectar tipos de columnas
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            # Números → rellenar con 0
            df[col] = df[col].fillna(0)
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            # Fechas → dejar como None (null en JSON)
            df[col] = df[col].where(pd.notnull(df[col]), None)
        else:
            # Texto → string vacío
            df[col] = df[col].fillna("").astype(str).str.strip()

    # Convertir fechas a string ISO (muy importante para web)
    for col in df.select_dtypes(include=['datetime64[ns]']).columns:
        df[col] = df[col].apply(lambda x: x.isoformat() if x else None)

    # Agregar ID único
    # df.insert(0, "id", range(1, len(df) + 1))

    # Convertir a lista de diccionarios
    data = df.to_dict(orient='records')

    # Guardar JSON optimizado
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))

    print(f"✅ JSON generado correctamente:\n{json_path}")

except Exception as e:
    print(f"❌ Error: {e}")