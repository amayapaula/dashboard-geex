import os
import pandas as pd
import numpy as np

def universal_date_parser(series):
    """
    Parseador universal de fechas. Las fechas ISO (AAAA-MM-DD, como las de Meta) se leen tal cual;
    el resto (texto DD/MM/AAAA, marcas con zona horaria, valores nativos de Excel) con día primero.
    Con dayfirst=True las ISO se daban vuelta (2026-10-11 pasaba a 10/11).
    """
    if series is None or len(series) == 0:
        return pd.Series(dtype='datetime64[ns]')
    if pd.api.types.is_datetime64_any_dtype(series):
        parsed = series
    else:
        t = series.astype('string').str.strip()
        es_iso = t.str.match(r'^\d{4}-\d{2}-\d{2}').fillna(False).astype(bool)
        parsed = pd.Series(pd.NaT, index=series.index, dtype='datetime64[ns]')
        parsed[es_iso] = pd.to_datetime(t[es_iso].str[:10], format='%Y-%m-%d', errors='coerce')
        resto = t[~es_iso]
        if len(resto):
            try:
                parsed[~es_iso] = pd.to_datetime(resto, format='mixed', dayfirst=True, errors='coerce')
            except Exception:
                parsed[~es_iso] = pd.to_datetime(resto, dayfirst=True, errors='coerce')
    # Eliminar zonas horarias (si vienen de Botmaker u otras plataformas)
    if hasattr(parsed.dt, 'tz') and parsed.dt.tz is not None:
        parsed = parsed.dt.tz_localize(None)
    return parsed

def process_all_excel_sheets():
    carpeta_out = "datos"
    os.makedirs(carpeta_out, exist_ok=True)
    
    archivo_excel = 'DASHHHHH.xlsx'
    if not os.path.exists(archivo_excel):
        print(f"Error: No se encontró el archivo {archivo_excel}")
        return
        
    print("Cargando y unificando pestañas de Excel...")
    xls = pd.ExcelFile(archivo_excel)
    
    sheets_mapping = {
        'RAW_horas': ('horas.parquet', ['Fecha']),
        'RAW_portaAR': ('porta.parquet', ['Fecha', 'Fecha inicio', 'Fecha estado actual', 'Fecha portabilidad']),
        'RAW_vozAR': ('ln.parquet', ['Fecha', 'Fecha Inicio', 'Fecha estado actual', 'Fecha de Activacion']),
        'RAW_baf': ('baf.parquet', ['Fecha', 'Fecha Inicio', 'Fecha estado actual', 'Fecha de Instalación']),
        'RAW_botmaker': ('botmaker.parquet', ['Fecha', 'Fecha Sesión']),
        'RAW_botmaker_anuncios': ('bot_an.parquet', ['Fecha', 'Día']),
        'RAW_landing': ('landing.parquet', ['Fecha', 'Fecha Ingreso', 'Fecha Primer Contacto', 'Fecha Ult Modific']),
        'RAW_meta': ('meta.parquet', ['Fecha', 'Día']),
        'RAW_google': ('google.parquet', ['Fecha', 'Día'])
    }
    
    for sheet_name, (parquet_file, date_cols) in sheets_mapping.items():
        if sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name)
            
            # Asegurar columna 'Fecha' estandarizada en todos los dataframes
            if 'Fecha' not in df.columns:
                for candidate in ['Fecha inicio', 'Fecha Inicio', 'Fecha Sesión', 'Fecha Ingreso', 'Día']:
                    if candidate in df.columns:
                        df['Fecha'] = df[candidate]
                        break
            
            # Limpiar y normalizar todas las columnas de fechas detectadas
            for col in date_cols:
                if col in df.columns:
                    df[col] = universal_date_parser(df[col])
            
            # Guardar en la carpeta datos
            out_path = os.path.join(carpeta_out, parquet_file)
            try:
                df.to_parquet(out_path, index=False)
                print(f"  ✓ {sheet_name:<22} -> {parquet_file} ({len(df)} filas procesadas)")
            except Exception as e:
                # Respaldo si no esta instalado pyarrow/fastparquet
                csv_path = out_path.replace('.parquet', '.csv')
                df.to_csv(csv_path, index=False)
                print(f"  ✓ {sheet_name:<22} -> {os.path.basename(csv_path)} ({len(df)} filas procesadas)")

    print("¡Procesamiento completo de datos finalizado con éxito!")

if __name__ == '__main__':
    process_all_excel_sheets()