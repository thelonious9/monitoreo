import pandas as pd
import json
import os
import win32com.client # Librería para controlar Excel

def refresh_power_query(file_path):
    print("Actualizando consultas de Power Query...")
    try:
        # Iniciamos la aplicación de Excel
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False # Que trabaje en segundo plano
        excel.DisplayAlerts = False
        
        # Abrimos el libro
        wb = excel.Workbooks.Open(file_path)
        
        # Actualizamos todas las conexiones (Power Query)
        wb.RefreshAll()
        
        # Esperamos a que la actualización termine (importante)
        excel.CalculateUntilAsyncQueriesDone()
        
        # Guardamos y cerramos
        wb.Save()
        wb.Close()
        excel.Quit()
        print("Power Query actualizado con éxito.")
    except Exception as e:
        print(f"No se pudo actualizar Power Query: {e}")
        # Intentamos cerrar Excel si quedó abierto por el error
        if 'excel' in locals():
            excel.Quit()

def normalize_text(text):
    if not isinstance(text, str) or text == "N/A":
        return "N/A"
    text = text.strip().capitalize()
    mapping = {
        'Ssph': 'SSPH', 'Difh': 'DIFH', 'Ieeh': 'IEEH', 'Teeh': 'TEEH',
        'Cdheh': 'CDHEH', 'Pgjeh': 'PGJEH', 'Tsjeh': 'TSJEH', 'Uaeh': 'UAEH',
        'Seph': 'SEPH', 'Rrss': 'RRSS', 'Tv': 'TV', 'Oficiala mayor': 'Oficialía Mayor'
    }
    return mapping.get(text, text)

def export_to_json(excel_path, json_path):
    # PASO 1: Forzar la actualización de Power Query antes de leer
    refresh_power_query(excel_path)

    print(f"Procesando datos para el Dashboard...")
    try:
        df = pd.read_excel(excel_path, sheet_name='Fuente')
        
        target_cols = ['Fecha', 'Medio', 'Dependencia', 'Organismo', 'Género', 'Canal', 'Tema', 'Estatus', 'Municipio', 'Región', 'Clasificación', 'Agrupación']
        if len(df.columns) >= 12:
            df.columns = target_cols + list(df.columns[12:])

        fields = ["Dependencia", "Tema", "Estatus", "Municipio", "Región", "Canal", "Medio", "Género"]
        for col in fields:
            if col in df.columns:
                df[col] = df[col].apply(normalize_text)
        
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%d/%m/%Y')
        
        if 'Fecha' in df.columns:
            df['Fecha'] = df['Fecha'].apply(lambda x: x.replace('/01/2025', '/01/2026') if isinstance(x, str) else x)
        
        df = df.fillna("N/A")
        
        data = df.to_dict(orient='records')
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            
        print("Success! Datos actualizados y exportados.")
    except Exception as e:
        print(f"Error crítico en la exportación: {e}")

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    excel_input = os.path.join(base_dir, "Data", "Fuente.xlsm")
    json_output = os.path.join(base_dir, "Data", "datos.json")
    
    export_to_json(excel_input, json_output)