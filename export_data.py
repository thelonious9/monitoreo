import pandas as pd
import json
import os
import win32com.client

def refresh_power_query(file_path):
    print(f"Actualizando Power Query en: {file_path}")
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False 
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(file_path)
        
        for conn in wb.Connections:
            if conn.Type == 1:
                conn.OLEDBConnection.BackgroundQuery = False
        
        wb.RefreshAll()
        wb.Save()
        wb.Close()
        print("Actualización de Excel completada.")
    except Exception as e:
        print(f"Error en Excel: {e}")
    finally:
        if excel:
            excel.Quit()

def export_to_json(excel_path, json_path):
    # 1. Actualizar datos desde el Excel
    refresh_power_query(excel_path)

    print(f"Procesando formatos y exportando a JSON...")
    try:
        df = pd.read_excel(excel_path, sheet_name='Fuente')
        
        # Mapeo de columnas
        target_cols = ['Fecha', 'Medio', 'Dependencia', 'Organismo', 'Género', 'Canal', 'Tema', 'Estatus', 'Municipio', 'Región', 'Clasificación', 'Agrupación']
        if len(df.columns) >= 12:
            df.columns = target_cols + list(df.columns[12:])

        # --- LÓGICA DE FORMATEO ACTUALIZADA ---
        
        # 1. Dependencia: TODO EN MAYÚSCULAS
        if 'Dependencia' in df.columns:
            df['Dependencia'] = df['Dependencia'].astype(str).str.strip().str.upper()
            
        # 2. Tema y Medio: Cada Palabra En Mayúscula (Title Case)
        # Se añadió 'Medio' aquí para que sea "Tipo Título"
        for col in ['Tema', 'Medio']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.title()

        # 3. Otros campos: Solo la primera letra en mayúscula
        other_fields = ["Estatus", "Municipio", "Región", "Canal", "Género"]
        for col in other_fields:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.capitalize()
        
        # Formateo de fechas
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%d/%m/%Y')
        
        # Limpieza final de nulos
        df = df.replace("Nan", "N/A").fillna("N/A")
        
        # Exportación a JSON
        data = df.to_dict(orient='records')
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            
        print(f"¡Éxito! El campo 'Medio' ahora también usa formato de título.")
    except Exception as e:
        print(f"Error crítico: {e}")

if __name__ == "__main__":
    folder_path = r"C:\Users\vivie\Dashboard\Data"
    excel_input = os.path.join(folder_path, "Fuente.xlsm")
    json_output = os.path.join(folder_path, "datos.json")
    
    export_to_json(excel_input, json_output)