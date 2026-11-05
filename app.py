import os
import streamlit as st
import pandas as pd
import numpy as np

# ==========================================
# 1. CONFIGURACIÓN INICIAL Y DISEÑO GEEX
# ==========================================
st.set_page_config(page_title="Dashboard Geex", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    .block-container { padding-top: 1rem !important; }
    .metric-card {
        background-color: #FFFFFF;
        padding: 8px;
        border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        text-align: center;
        margin-bottom: 10px;
        border: 1px solid #E0E0E0;
        border-top: 4px solid #0B192C;
        white-space: nowrap; 
    }
    .metric-value { font-size: 19px; font-weight: bold; color: #0B192C; }
    .metric-label { font-size: 9px; color: #555555; text-transform: uppercase; font-weight: 700; letter-spacing: 0px; }
    .section-header {
        font-size: 18px;
        font-weight: bold;
        color: #0B192C;
        margin-top: 15px;
        margin-bottom: 10px;
        border-bottom: 2px solid #0B192C;
        padding-bottom: 5px;
    }
    .sub-header { font-size: 14px; font-weight: bold; color: #333333; margin-top: 20px; }
    </style>
""", unsafe_allow_html=True)

col_titulo, col_filtro = st.columns([7, 1])

with col_titulo:
    st.title("Dashboard Geex")

with col_filtro:
    st.write("") 
    meses_disponibles = {
        "Octubre": (2026, 10),
        "Noviembre": (2026, 11)
    }
    mes_seleccionado = st.selectbox("Mes", list(meses_disponibles.keys()), label_visibility="collapsed")
    anio_sel, mes_sel = meses_disponibles[mes_seleccionado]

# ==========================================
# 2. CARGA DE DATOS
# ==========================================
@st.cache_data
def load_raw_data_final():
    carpeta = "datos"
    if os.path.exists(os.path.join(carpeta, "horas.parquet")):
        horas = pd.read_parquet(os.path.join(carpeta, "horas.parquet"))
        porta = pd.read_parquet(os.path.join(carpeta, "porta.parquet"))
        ln = pd.read_parquet(os.path.join(carpeta, "ln.parquet"))
        baf = pd.read_parquet(os.path.join(carpeta, "baf.parquet"))
        botmaker = pd.read_parquet(os.path.join(carpeta, "botmaker.parquet"))
        bot_anuncios = pd.read_parquet(os.path.join(carpeta, "bot_an.parquet"))
        landing = pd.read_parquet(os.path.join(carpeta, "landing.parquet"))
        meta_df = pd.read_parquet(os.path.join(carpeta, "meta.parquet"))
        google_df = pd.read_parquet(os.path.join(carpeta, "google.parquet"))
        
        porta['Fecha'] = porta['Fecha']
        porta['Fecha inicio'] = porta['Fecha']
        ln['Fecha'] = ln['Fecha']
        ln['Fecha Inicio'] = ln['Fecha']
        baf['Fecha'] = baf['Fecha']
        baf['Fecha Inicio'] = baf['Fecha']
        botmaker['Fecha Sesión'] = botmaker['Fecha']
        bot_anuncios['Fecha'] = pd.NaT
        landing['Fecha Ingreso'] = landing['Fecha']
        meta_df['Día'] = meta_df['Fecha']
        google_df['Día'] = google_df['Fecha']
    else:
        xls = pd.ExcelFile('DASHHHHH.xlsx')
        horas = pd.read_excel(xls, 'RAW_horas')
        porta = pd.read_excel(xls, 'RAW_portaAR')
        ln = pd.read_excel(xls, 'RAW_vozAR')
        baf = pd.read_excel(xls, 'RAW_baf')
        botmaker = pd.read_excel(xls, 'RAW_botmaker')
        bot_anuncios = pd.read_excel(xls, 'RAW_botmaker_anuncios') 
        landing = pd.read_excel(xls, 'RAW_landing')
        meta_df = pd.read_excel(xls, 'RAW_meta')
        google_df = pd.read_excel(xls, 'RAW_google')
        
        horas['Fecha'] = pd.to_datetime(horas['Fecha'], dayfirst=True, errors='coerce')
        porta['Fecha'] = pd.to_datetime(porta['Fecha inicio'], dayfirst=True, errors='coerce')
        ln['Fecha'] = pd.to_datetime(ln['Fecha Inicio'], dayfirst=True, errors='coerce')
        baf['Fecha'] = pd.to_datetime(baf['Fecha Inicio'], dayfirst=True, errors='coerce')
        botmaker['Fecha Sesión'] = pd.to_datetime(botmaker['Fecha Sesión'], dayfirst=True, errors='coerce')
        bot_anuncios['Fecha'] = pd.to_datetime(bot_anuncios['Fecha'], errors='coerce')
        landing['Fecha Ingreso'] = pd.to_datetime(landing['Fecha Ingreso'], dayfirst=True, errors='coerce')
        meta_df['Día'] = pd.to_datetime(meta_df['Día'], errors='coerce')
        google_df['Día'] = pd.to_datetime(google_df['Día'], errors='coerce')
        
    return horas, porta, ln, baf, botmaker, bot_anuncios, landing, meta_df, google_df

# ==========================================
# 3. CÁLCULOS SEGÚN MES
# ==========================================
@st.cache_data
def calculate_geex_metrics_final(anio, mes):
    horas_raw, porta_raw, ln_raw, baf_raw, bot_raw, bot_an_raw, land_raw, meta_raw, google_raw = load_raw_data_final()
    
    horas = horas_raw.copy()
    porta = porta_raw.copy()
    ln = ln_raw.copy()
    baf = baf_raw.copy()
    botmaker = bot_raw.copy()
    bot_an = bot_an_raw.copy()
    landing = land_raw.copy()
    
    def clean_sales(df, state_col):
        if state_col in df.columns:
            return df[df[state_col].astype(str).str.strip().str.upper() != 'NO APROBADO']
        return df

    p_val = clean_sales(porta, 'Estado actual')
    l_val = clean_sales(ln, 'Estado Actual')
    b_val = clean_sales(baf, 'Estado Actual')
    
    sups_digital = ['RULFO, YAMIL', 'PUENTES, SOLEDAD']
    
    p_dig = p_val[(p_val['Supervisor'].isin(sups_digital)) & (p_val['Fecha'].dt.month == mes) & (p_val['Fecha'].dt.year == anio)].copy()
    l_dig = l_val[(l_val['Supervisor'].isin(sups_digital)) & (l_val['Fecha'].dt.month == mes) & (l_val['Fecha'].dt.year == anio)].copy()
    b_dig = b_val[(b_val['Supervisor'].isin(sups_digital)) & (b_val['Fecha'].dt.month == mes) & (b_val['Fecha'].dt.year == anio)].copy()
    
    for df_s in [p_dig, l_dig, b_dig]:
        if 'BOTMAKER' in df_s.columns: df_s['BOTMAKER'] = df_s['BOTMAKER'].astype(str).str.strip()
        if 'LANDING' in df_s.columns: df_s['LANDING'] = df_s['LANDING'].astype(str).str.strip()
    
    bot_m = botmaker[(botmaker['PAIS'] == 'AR') & (botmaker['Fecha Sesión'].dt.month == mes) & (botmaker['Fecha Sesión'].dt.year == anio)].copy()
    land_m = landing[(landing['Pais'] == 'AR') & (landing['Fecha Ingreso'].dt.month == mes) & (landing['Fecha Ingreso'].dt.year == anio)].copy()
    
    total_leads = len(bot_m) + len(land_m)
    
    horas['Horas'] = pd.to_numeric(horas['Horas'], errors='coerce').fillna(0)
    horas_m = horas[(horas['Fecha'].dt.month == mes) & (horas['Fecha'].dt.year == anio) & (horas['Supervisor'].isin(sups_digital))]
    
    h_yam = horas_m[horas_m['Supervisor'] == 'RULFO, YAMIL']['Horas'].sum()
    h_pue = horas_m[horas_m['Supervisor'] == 'PUENTES, SOLEDAD']['Horas'].sum()
    
    v_porta = len(p_dig)
    v_ln = len(l_dig)
    v_baf = len(b_dig)
    v_totales = v_porta + v_ln + v_baf
    h_totales = h_yam + h_pue

    # --- TABLA DIARIA (Productividad) ---
    start_date = pd.Timestamp(year=anio, month=mes, day=1)
    end_date = start_date + pd.offsets.MonthEnd(1)
    dates_m = pd.date_range(start=start_date, end=end_date)
    daily_rows = []
    
    for d in dates_m:
        d_date = d.date()
        is_weekend = d.weekday() >= 5
        is_holiday = (mes == 10 and d.day == 12) or (mes == 11 and d.day in [2, 23])
        is_non_working = is_weekend or is_holiday
        
        h_d = horas_m[horas_m['Fecha'].dt.date == d_date]
        if h_d['Horas'].sum() == 0: continue
            
        p_d = p_dig[p_dig['Fecha'].dt.date == d_date]
        l_d = l_dig[l_dig['Fecha'].dt.date == d_date]
        b_d = b_dig[b_dig['Fecha'].dt.date == d_date]
        
        h_y = h_d[h_d['Supervisor'] == 'RULFO, YAMIL']
        p_y = p_d[p_d['Supervisor'] == 'RULFO, YAMIL']
        l_y = l_d[l_d['Supervisor'] == 'RULFO, YAMIL']
        b_y = b_d[b_d['Supervisor'] == 'RULFO, YAMIL']
        
        pres_m = int(h_y[h_y['Horas'] > 0]['Usuario'].nunique()) if h_y['Horas'].sum() > 0 else 0
        horas_mo = int(h_y['Horas'].sum())
        porta_mo, ln_mo, baf_mo = len(p_y), len(l_y), len(b_y)
        tot_mo = porta_mo + ln_mo + baf_mo
        
        sph_m = tot_mo / horas_mo if horas_mo > 0 else 0
        sph_bp_m = (porta_mo + baf_mo) / horas_mo if horas_mo > 0 else 0
        
        h_p = h_d[h_d['Supervisor'] == 'PUENTES, SOLEDAD']
        p_p = p_d[p_d['Supervisor'] == 'PUENTES, SOLEDAD']
        l_p = l_d[l_d['Supervisor'] == 'PUENTES, SOLEDAD']
        b_p = b_d[b_d['Supervisor'] == 'PUENTES, SOLEDAD']
        
        pres_t = int(h_p[h_p['Horas'] > 0]['Usuario'].nunique()) if h_p['Horas'].sum() > 0 else 0
        horas_t = int(h_p['Horas'].sum())
        porta_t, ln_t, baf_t = len(p_p), len(l_p), len(b_p)
        tot_t = porta_t + ln_t + baf_t
        
        sph_t = tot_t / horas_t if horas_t > 0 else 0
        sph_bp_t = (porta_t + baf_t) / horas_t if horas_t > 0 else 0
        
        tot_porta = porta_mo + porta_t
        tot_ln = ln_mo + ln_t
        tot_baf = baf_mo + baf_t
        tot_ventas = tot_mo + tot_t
        horas_tot = horas_mo + horas_t
        
        sph_tot = tot_ventas / horas_tot if horas_tot > 0 else 0
        sph_bp_tot = (tot_porta + tot_baf) / horas_tot if horas_tot > 0 else 0
        
        leads_dia = len(bot_m[bot_m['Fecha Sesión'].dt.date == d_date]) + len(land_m[land_m['Fecha Ingreso'].dt.date == d_date])
        efect = tot_ventas / leads_dia if leads_dia > 0 else 0
        
        daily_rows.append({
            "Día": d.strftime('%d-%b'), "_No_Laborable": is_non_working,
            "Presentes M": pres_m if pres_m > 0 else "", "Horas M": horas_mo if horas_mo > 0 else "",
            "Porta M": porta_mo, "LN M": ln_mo, "BAF M": baf_mo, "Total M": tot_mo,
            "SPH M": f"{sph_m:.2f}", "SPH BAF+Porta M": f"{sph_bp_m:.2f}",
            "Presentes T": pres_t if pres_t > 0 else "", "Horas T": horas_t if horas_t > 0 else "",
            "Porta T": porta_t, "LN T": ln_t, "BAF T": baf_t, "Total T": tot_t,
            "SPH T": f"{sph_t:.2f}", "SPH BAF+Porta T": f"{sph_bp_t:.2f}",
            "Porta Tot": tot_porta, "LN Tot": tot_ln, "BAF Tot": tot_baf, "Total Ventas": tot_ventas,
            "Horas Totales": horas_tot, "SPH Total": f"{sph_tot:.2f}", "SPH BAF+Porta": f"{sph_bp_tot:.2f}",
            "Leads Totales": leads_dia, "Efectividad": f"{efect:.2%}"
        })

    df_daily = pd.DataFrame(daily_rows)

    # --- TABLA ASESORES ---
    h_sum = horas_m.groupby(['Supervisor', 'Asesor'])['Horas'].sum().reset_index()
    p_sum = p_dig.groupby('Asesor').size().reset_index(name='PORTA')
    l_sum = l_dig.groupby('Asesor').size().reset_index(name='LN')
    b_sum = b_dig.groupby('Asesor').size().reset_index(name='BAF')
    df_asesores = h_sum.merge(p_sum, on='Asesor', how='left').merge(l_sum, on='Asesor', how='left').merge(b_sum, on='Asesor', how='left').fillna(0)
    df_asesores['Horas'] = df_asesores['Horas'].astype(int)
    for c in ['PORTA', 'LN', 'BAF']: df_asesores[c] = df_asesores[c].astype(int)
    df_asesores['Total Ventas'] = df_asesores['PORTA'] + df_asesores['LN'] + df_asesores['BAF']
    df_asesores['SPH Total'] = df_asesores['Total Ventas'] / df_asesores['Horas']
    df_asesores['SPH (baf/porta)'] = (df_asesores['PORTA'] + df_asesores['BAF']) / df_asesores['Horas']
    df_asesores = df_asesores.sort_values(by='SPH (baf/porta)', ascending=False).reset_index(drop=True)

    # --- DATOS PIVOTADOS DE PLANES ---
    p_dig['Día'] = p_dig['Fecha'].dt.day
    b_dig['Día'] = b_dig['Fecha'].dt.day
    p_planes_pivot = pd.pivot_table(p_dig, index='Día', columns='PLANES', aggfunc='size', fill_value=0)
    if not p_planes_pivot.empty: p_planes_pivot = p_planes_pivot[p_planes_pivot.sum().sort_values(ascending=False).index]
    esim_pivot = pd.pivot_table(p_dig, index='Día', columns='ESIM', aggfunc='size', fill_value=0)
    if not esim_pivot.empty: esim_pivot = esim_pivot[esim_pivot.sum().sort_values(ascending=False).index]
    baf_pivot = pd.pivot_table(b_dig, index='Día', columns='PLANES', aggfunc='size', fill_value=0)
    if not baf_pivot.empty: baf_pivot = baf_pivot[baf_pivot.sum().sort_values(ascending=False).index]

    # --- DATOS / LEADS DIARIOS Y POR HORA ---
    bot_d = bot_m.groupby(bot_m['Fecha Sesión'].dt.date).size().reset_index(name='Botmaker')
    land_d = land_m.groupby(land_m['Fecha Ingreso'].dt.date).size().reset_index(name='Landing')
    df_leads_d = pd.merge(bot_d, land_d, left_on='Fecha Sesión', right_on='Fecha Ingreso', how='outer').fillna(0)
    df_leads_d['Día'] = pd.to_datetime(df_leads_d['Fecha Sesión'].combine_first(df_leads_d['Fecha Ingreso'])).dt.strftime('%d-%b')
    df_leads_d['Total'] = df_leads_d['Botmaker'] + df_leads_d['Landing']
    df_leads_d = df_leads_d[['Día', 'Botmaker', 'Landing', 'Total']]

    bot_m['Hora'] = pd.to_numeric(bot_m['Hora Sesión'], errors='coerce').fillna(0).astype(int)
    land_m['Hora'] = land_m['Fecha Ingreso'].dt.hour
    bot_h = bot_m.groupby('Hora').size().reset_index(name='Botmaker')
    land_h = land_m.groupby('Hora').size().reset_index(name='Landing')
    df_leads_h = pd.merge(bot_h, land_h, on='Hora', how='outer').fillna(0)
    df_leads_h['Total'] = df_leads_h['Botmaker'] + df_leads_h['Landing']
    df_leads_h = df_leads_h.sort_values('Hora').reset_index(drop=True)

    # --- LIMPIEZA DE ANUNCIOS (EFECTIVIDAD) ---
    bot_m['join_num'] = bot_m['Id Contacto/Número'].astype(str).str.replace(r'\.0$', '', regex=True).str[-10:]
    bot_an['join_num'] = bot_an['NIM'].astype(str).str.replace(r'\.0$', '', regex=True).str[-10:]
    
    bot_an_clean_ef = bot_an.dropna(subset=['Anuncio_Titulo']).copy()
    bot_an_clean_ef['Anuncio_Titulo'] = bot_an_clean_ef['Anuncio_Titulo'].astype(str).str.strip()
    bot_an_unique = bot_an_clean_ef[['join_num', 'Anuncio_Titulo', 'Producto']].drop_duplicates(subset=['join_num'], keep='last')
    bot_m_merged = pd.merge(bot_m, bot_an_unique, on='join_num', how='left')

    land_m_clean = land_m.dropna(subset=['Fuente']).copy()
    land_m_clean['Fuente'] = land_m_clean['Fuente'].astype(str).str.strip()

    # ==========================================
    # --- PRESUPUESTO DIARIO, REALES Y CPL ---
    # ==========================================
    m_ar = meta_raw[meta_raw['Pais'].astype(str).str.strip().str.upper() == 'AR'].copy()
    m_ar['Prod_Clean'] = m_ar['Producto'].astype(str).str.strip().str.upper()

    g_ar = google_raw[google_raw['Pais'].astype(str).str.strip().str.upper() == 'AR'].copy()
    g_ar['Prod_Clean'] = g_ar['Producto'].astype(str).str.strip().str.upper()
    g_ar['Costo_Final'] = g_ar['Costo num'] if 'Costo num' in g_ar.columns else g_ar['Costo']

    b_all = bot_raw.copy()
    l_all = land_raw.copy()

    bot_an_copy = bot_an_raw.copy()
    b_all['join_num'] = b_all['Id Contacto/Número'].astype(str).str.replace(r'\.0$', '', regex=True).str[-10:]
    bot_an_copy['join_num'] = bot_an_copy['NIM'].astype(str).str.replace(r'\.0$', '', regex=True).str[-10:]
    
    bot_an_clean_pres = bot_an_copy.dropna(subset=['Anuncio_Titulo']).drop_duplicates(subset=['join_num'], keep='last')
    b_all = b_all.merge(bot_an_clean_pres[['join_num', 'Plataforma', 'Producto']], on='join_num', how='left')

    l_all['Plataforma'] = l_all['Plataforma'].astype(str).str.strip().str.upper()
    l_all['Producto'] = l_all['Producto'].astype(str).str.strip().str.upper()
    b_all['Plataforma'] = b_all['Plataforma'].astype(str).str.strip().str.upper()
    b_all['Producto'] = b_all['Producto'].astype(str).str.strip().str.upper()

    # Fechas de facturación exactas según mes
    if mes == 10:
        start_d = pd.Timestamp('2026-10-01')
        end_d = pd.Timestamp('2026-10-31')
    elif mes == 11:
        start_d = pd.Timestamp('2026-10-29')
        end_d = pd.Timestamp('2026-11-26')
    else:
        start_d = pd.Timestamp(year=anio, month=mes, day=1)
        end_d = start_d + pd.offsets.MonthEnd(1)

    dates_presupuesto = pd.date_range(start_d, end_d)
    pres_rows = []
    
    for d in dates_presupuesto:
        d_val = d.date()
        
        if mes == 10:
            in_meta = pd.Timestamp('2026-10-02') <= d <= pd.Timestamp('2026-10-28')
            in_goog = pd.Timestamp('2026-10-01') <= d <= pd.Timestamp('2026-10-31')
        elif mes == 11:
            in_meta = pd.Timestamp('2026-10-29') <= d <= pd.Timestamp('2026-11-26')
            in_goog = pd.Timestamp('2026-11-01') <= d <= pd.Timestamp('2026-11-26')
        else:
            in_meta = True; in_goog = True

        if in_meta:
            m_day = m_ar[m_ar['Día'].dt.date == d_val]
            p_meta = m_day['Monto gastado (ARS)'].sum()
            p_m_baf = m_day[m_day['Prod_Clean'] == 'BAF']['Monto gastado (ARS)'].sum()
            p_m_porta = m_day[m_day['Prod_Clean'] == 'PORTA']['Monto gastado (ARS)'].sum()
        else:
            p_meta = p_m_baf = p_m_porta = 0
            
        if in_goog:
            g_day = g_ar[g_ar['Día'].dt.date == d_val]
            p_goog = g_day['Costo_Final'].sum()
            p_g_baf = g_day[g_day['Prod_Clean'] == 'BAF']['Costo_Final'].sum()
            p_g_porta = g_day[g_day['Prod_Clean'] == 'PORTA']['Costo_Final'].sum()
        else:
            p_goog = p_g_baf = p_g_porta = 0

        tot_pres = p_meta + p_goog
        p_baf = p_m_baf + p_g_baf
        p_porta = p_m_porta + p_g_porta

        b_day = b_all[b_all['Fecha Sesión'].dt.date == d_val]
        l_day = l_all[l_all['Fecha Ingreso'].dt.date == d_val]

        bot_no_id = b_day['Plataforma'].isin(['NAN', 'NONE', '']).sum() + b_day['Plataforma'].isna().sum()
        bot_m_meta = len(b_day[b_day['Plataforma'] == 'META'])
        land_m_meta = len(l_day[l_day['Plataforma'] == 'META'])
        bot_m_goog = len(b_day[b_day['Plataforma'] == 'GOOGLE'])
        land_m_goog = len(l_day[l_day['Plataforma'] == 'GOOGLE'])

        leads_meta = bot_no_id + bot_m_meta + land_m_meta
        leads_goog = bot_m_goog + land_m_goog
        leads_tot = leads_meta + leads_goog

        bot_m_baf = len(b_day[b_day['Producto'] == 'BAF'])
        land_m_baf = len(l_day[l_day['Producto'] == 'BAF'])
        leads_baf = bot_no_id + bot_m_baf + land_m_baf
        
        bot_m_porta = len(b_day[b_day['Producto'] == 'PORTA'])
        land_m_porta = len(l_day[l_day['Producto'] == 'PORTA'])
        leads_porta = bot_m_porta + land_m_porta

        cpl_meta = p_meta / leads_meta if leads_meta > 0 else 0
        cpl_goog = p_goog / leads_goog if leads_goog > 0 else 0
        cpl_baf = p_baf / leads_baf if leads_baf > 0 else 0
        cpl_porta = p_porta / leads_porta if leads_porta > 0 else 0

        pres_rows.append({
            'Fecha': d.strftime('%d/%m/%Y'),
            'Total ($)': tot_pres,
            'Meta ($)': p_meta, '% Meta ($)': p_meta / tot_pres if tot_pres > 0 else 0,
            'Google ($)': p_goog, '% Google ($)': p_goog / tot_pres if tot_pres > 0 else 0,
            'Baf ($)': p_baf, '% Baf ($)': p_baf / tot_pres if tot_pres > 0 else 0,
            'Porta ($)': p_porta, '% Porta ($)': p_porta / tot_pres if tot_pres > 0 else 0,
            'Total Leads': leads_tot,
            'Botmaker no se identifica': bot_no_id,
            'Meta': leads_meta, 'Landing Meta': land_m_meta, 'Botmaker Meta': bot_m_meta, '% Meta (Leads)': leads_meta / leads_tot if leads_tot > 0 else 0,
            'Google': leads_goog, 'Landing Google': land_m_goog, '% Google (Leads)': leads_goog / leads_tot if leads_tot > 0 else 0,
            'BAF': leads_baf, 'Landing BAF': land_m_baf, 'Botmaker BAF': bot_m_baf, '% BAF': leads_baf / leads_tot if leads_tot > 0 else 0,
            'Porta': leads_porta, 'Landing Porta': land_m_porta, 'Botmaker Porta': bot_m_porta, '% Porta': leads_porta / leads_tot if leads_tot > 0 else 0,
            'cpl meta': cpl_meta, 'cpl google': cpl_goog,
            'cpl baf': cpl_baf, 'cpl porta': cpl_porta
        })
        
    df_pres_master = pd.DataFrame(pres_rows)

    return {
        'v_porta': v_porta, 'v_ln': v_ln, 'v_baf': v_baf, 'v_totales': v_totales,
        'leads_tot': total_leads, 'h_totales': h_totales,
        'sph_total': v_totales / h_totales if h_totales > 0 else 0,
        'sph_baf_porta': (v_porta + v_baf) / h_totales if h_totales > 0 else 0,
        'efectividad': v_totales / total_leads if total_leads > 0 else 0,
        'df_daily': df_daily, 'df_asesores': df_asesores,
        'horas_oct': horas_m, 'p_dig': p_dig, 'l_dig': l_dig, 'b_dig': b_dig,
        'df_leads_d': df_leads_d, 'df_leads_h': df_leads_h,
        'p_planes_pivot': p_planes_pivot, 'esim_pivot': esim_pivot, 'baf_pivot': baf_pivot,
        'bot_m_merged': bot_m_merged, 'land_m': land_m_clean,
        'df_pres_master': df_pres_master
    }

# ==========================================
# 4. CAPTURA DE ERRORES Y RENDERIZADO
# ==========================================
try:
    m = calculate_geex_metrics_final(anio_sel, mes_sel)
except Exception as e:
    st.error(f"Ocurrió un error procesando los datos: {e}")
    st.stop()

tab_historico, tab_prod, tab_asesores, tab_datos, tab_efectividad, tab_presupuesto = st.tabs([
    "Resumen Histórico", "Productividad", "Asesores", "Datos", "Efectividad", "Presupuesto"
])

# --- 1. PESTAÑA HISTÓRICO (Con flechas comparativas respecto al mes anterior) ---
with tab_historico:
    st.markdown("<div class='section-header'>Resumen Histórico Operativo (CC 6 a CC 11)</div>", unsafe_allow_html=True)
    st.info("Vista consolidada por mes con comparativa respecto al mes anterior.")

    # Función para agregar flechas comparativas
    def add_comparison_arrows(data_list, num_cols, lower_is_better_cols=None):
        if lower_is_better_cols is None: lower_is_better_cols = []
        out = []
        for i, row in enumerate(data_list):
            new_row = {}
            for k, v in row.items():
                if k in num_cols:
                    is_lower_better = k in lower_is_better_cols
                    if i == 0:
                        arrow = ""
                    else:
                        prev_v = data_list[i-1][k]
                        if v > prev_v:
                            arrow = " 🔴 ↑" if is_lower_better else " 🟢 ↑"
                        elif v < prev_v:
                            arrow = " 🟢 ↓" if is_lower_better else " 🔴 ↓"
                        else:
                            arrow = " ➖"
                    
                    if k in ["SPH Total", "SPH BAF+Porta", "SPH Mañana", "SPH Tarde"]:
                        val_str = f"{v:.2f}"
                    elif k in ["Efectividad"]:
                        val_str = f"{v:.0%}" if isinstance(v, float) and v <= 1 else f"{v}"
                    elif "Presupuesto" in k or "CPL" in k:
                        val_str = f"${v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    elif k in ["Leads Totales", "Horas Total", "Horas Mañana", "Horas Tarde"]:
                        val_str = f"{v:,}".replace(",", ".")
                    else:
                        val_str = f"{v}"
                    
                    new_row[k] = f"{val_str}{arrow}"
                else:
                    new_row[k] = v
            out.append(new_row)
        return pd.DataFrame(out)

    datos_generales_raw = [
        {"CC": 6, "Mes": "Junio", "Porta": 518, "LN": 350, "BAF": 560, "Ventas Totales": 1428, "Horas Total": 3150, "SPH Total": 0.45, "SPH BAF+Porta": 0.34, "Leads Totales": 16834, "Efectividad": 0.08, "Presupuesto Ciclo": 33831810.37, "CPL Mensual": 2009.73, "CPL Ciclo": 1638.50},
        {"CC": 7, "Mes": "Julio", "Porta": 399, "LN": 395, "BAF": 540, "Ventas Totales": 1334, "Horas Total": 3156, "SPH Total": 0.42, "SPH BAF+Porta": 0.30, "Leads Totales": 18110, "Efectividad": 0.07, "Presupuesto Ciclo": 34736333.59, "CPL Mensual": 1918.07, "CPL Ciclo": 2092.93},
        {"CC": 8, "Mes": "Agosto", "Porta": 542, "LN": 365, "BAF": 494, "Ventas Totales": 1401, "Horas Total": 3030, "SPH Total": 0.46, "SPH BAF+Porta": 0.34, "Leads Totales": 15567, "Efectividad": 0.09, "Presupuesto Ciclo": 37593975.01, "CPL Mensual": 2414.98, "CPL Ciclo": 2673.25},
        {"CC": 9, "Mes": "Septiembre", "Porta": 511, "LN": 248, "BAF": 444, "Ventas Totales": 1203, "Horas Total": 3282, "SPH Total": 0.37, "SPH BAF+Porta": 0.29, "Leads Totales": 10866, "Efectividad": 0.11, "Presupuesto Ciclo": 37823181.54, "CPL Mensual": 3480.87, "CPL Ciclo": 2963.73},
        {"CC": 10, "Mes": "Octubre", "Porta": 366, "LN": 230, "BAF": 469, "Ventas Totales": 1065, "Horas Total": 2208, "SPH Total": 0.48, "SPH BAF+Porta": 0.38, "Leads Totales": 12412, "Efectividad": 0.09, "Presupuesto Ciclo": 29823696.35, "CPL Mensual": 2402.81, "CPL Ciclo": 2606.97},
        {"CC": 11, "Mes": "Noviembre", "Porta": 29, "LN": 20, "BAF": 37, "Ventas Totales": 86, "Horas Total": 216, "SPH Total": 0.40, "SPH BAF+Porta": 0.31, "Leads Totales": 1382, "Efectividad": 0.06, "Presupuesto Ciclo": 457590.32, "CPL Mensual": 331.11, "CPL Ciclo": 227.09}
    ]
    cols_gen = ["Porta", "LN", "BAF", "Ventas Totales", "Horas Total", "SPH Total", "SPH BAF+Porta", "Leads Totales", "Efectividad", "Presupuesto Ciclo", "CPL Mensual", "CPL Ciclo"]
    df_gen_show = add_comparison_arrows(datos_generales_raw, cols_gen, ["CPL Mensual", "CPL Ciclo"])
    st.dataframe(df_gen_show, use_container_width=True, hide_index=True)

    st.markdown("<div class='section-header'>Desglose Histórico por Turnos (Mañana vs Tarde)</div>", unsafe_allow_html=True)
    sub_m, sub_t = st.tabs(["Turno Mañana (Yamil)", "Turno Tarde (Soledad)"])
    
    with sub_m:
        datos_manana_raw = [
            {"CC": 6, "Mes": "Junio", "Horas Mañana": 1728, "Porta": 295, "LN": 180, "BAF": 274, "Total Yamil": 749, "SPH Mañana": 0.43, "SPH BAF+Porta": 0.33},
            {"CC": 7, "Mes": "Julio", "Horas Mañana": 1728, "Porta": 203, "LN": 206, "BAF": 265, "Total Yamil": 674, "SPH Mañana": 0.39, "SPH BAF+Porta": 0.27},
            {"CC": 8, "Mes": "Agosto", "Horas Mañana": 1962, "Porta": 375, "LN": 213, "BAF": 336, "Total Yamil": 924, "SPH Mañana": 0.47, "SPH BAF+Porta": 0.36},
            {"CC": 9, "Mes": "Septiembre", "Horas Mañana": 2184, "Porta": 329, "LN": 130, "BAF": 333, "Total Yamil": 792, "SPH Mañana": 0.36, "SPH BAF+Porta": 0.30},
            {"CC": 10, "Mes": "Octubre", "Horas Mañana": 1506, "Porta": 253, "LN": 130, "BAF": 336, "Total Yamil": 719, "SPH Mañana": 0.48, "SPH BAF+Porta": 0.39},
            {"CC": 11, "Mes": "Noviembre", "Horas Mañana": 144, "Porta": 17, "LN": 11, "BAF": 20, "Total Yamil": 48, "SPH Mañana": 0.33, "SPH BAF+Porta": 0.26}
        ]
        cols_m = ["Horas Mañana", "Porta", "LN", "BAF", "Total Yamil", "SPH Mañana", "SPH BAF+Porta"]
        st.dataframe(add_comparison_arrows(datos_manana_raw, cols_m), use_container_width=True, hide_index=True)

    with sub_t:
        datos_tarde_raw = [
            {"CC": 6, "Mes": "Junio", "Horas Tarde": 1422, "Porta": 223, "LN": 170, "BAF": 286, "Total Tarde": 679, "SPH Tarde": 0.48, "SPH BAF+Porta": 0.36},
            {"CC": 7, "Mes": "Julio", "Horas Tarde": 1428, "Porta": 196, "LN": 189, "BAF": 275, "Total Tarde": 660, "SPH Tarde": 0.46, "SPH BAF+Porta": 0.33},
            {"CC": 8, "Mes": "Agosto", "Horas Tarde": 1068, "Porta": 167, "LN": 152, "BAF": 158, "Total Tarde": 477, "SPH Tarde": 0.45, "SPH BAF+Porta": 0.30},
            {"CC": 9, "Mes": "Septiembre", "Horas Tarde": 1098, "Porta": 182, "LN": 118, "BAF": 111, "Total Tarde": 411, "SPH Tarde": 0.37, "SPH BAF+Porta": 0.27},
            {"CC": 10, "Mes": "Octubre", "Horas Tarde": 702, "Porta": 113, "LN": 100, "BAF": 133, "Total Tarde": 346, "SPH Tarde": 0.49, "SPH BAF+Porta": 0.35},
            {"CC": 11, "Mes": "Noviembre", "Horas Tarde": 72, "Porta": 12, "LN": 9, "BAF": 17, "Total Tarde": 38, "SPH Tarde": 0.53, "SPH BAF+Porta": 0.40}
        ]
        cols_t = ["Horas Tarde", "Porta", "LN", "BAF", "Total Tarde", "SPH Tarde", "SPH BAF+Porta"]
        st.dataframe(add_comparison_arrows(datos_tarde_raw, cols_t), use_container_width=True, hide_index=True)

# --- 2. PESTAÑA PRODUCTIVIDAD ---
with tab_prod:
    c1, c2, c3, c4, c5, c6, c7, c8, c9 = st.columns(9)
    with c1: st.markdown(f"<div class='metric-card'><div class='metric-label'>Ventas Brutas</div><div class='metric-value'>{m['v_totales']}</div></div>", unsafe_allow_html=True)
    with c2: st.markdown(f"<div class='metric-card'><div class='metric-label'>Portabilidad</div><div class='metric-value'>{m['v_porta']}</div></div>", unsafe_allow_html=True)
    with c3: st.markdown(f"<div class='metric-card'><div class='metric-label'>Línea Nueva</div><div class='metric-value'>{m['v_ln']}</div></div>", unsafe_allow_html=True)
    with c4: st.markdown(f"<div class='metric-card'><div class='metric-label'>BAF</div><div class='metric-value'>{m['v_baf']}</div></div>", unsafe_allow_html=True)
    with c5: st.markdown(f"<div class='metric-card'><div class='metric-label'>Leads Totales</div><div class='metric-value'>{m['leads_tot']:,}</div></div>", unsafe_allow_html=True)
    with c6: st.markdown(f"<div class='metric-card'><div class='metric-label'>Horas Totales</div><div class='metric-value'>{m['h_totales']:,.0f}</div></div>", unsafe_allow_html=True)
    with c7: st.markdown(f"<div class='metric-card'><div class='metric-label'>SPH Total</div><div class='metric-value'>{m['sph_total']:.2f}</div></div>", unsafe_allow_html=True)
    with c8: st.markdown(f"<div class='metric-card'><div class='metric-label'>SPH BAF+Porta</div><div class='metric-value'>{m['sph_baf_porta']:.2f}</div></div>", unsafe_allow_html=True)
    with c9: st.markdown(f"<div class='metric-card'><div class='metric-label'>Efectividad</div><div class='metric-value'>{m['efectividad']:.2%}</div></div>", unsafe_allow_html=True)

    st.markdown("---")
    
    if m['df_daily'].empty:
        st.warning(f"No hay registros de productividad para {mes_seleccionado}.")
    else:
        non_working_flags = m['df_daily']['_No_Laborable'].values
        df_clean = m['df_daily'].drop(columns=['_No_Laborable'])

        def style_geex_dataframe(row):
            idx = row.name
            is_non_working = non_working_flags[idx]
            styles = []
            for col in row.index:
                style_str = ''
                if is_non_working: style_str = 'background-color: #EAEAEA; color: #555555;'
                elif 'M' in col and col not in ['Día', 'Efectividad', 'Leads Totales', 'Horas Totales', 'Total Ventas']: style_str = 'background-color: #EBF3FA;'
                elif 'T' in col and col not in ['Total Ventas', 'Porta Tot', 'LN Tot', 'BAF Tot', 'Horas Totales', 'SPH Total', 'SPH BAF+Porta', 'Leads Totales', 'Efectividad']: style_str = 'background-color: #FAF4EB;'
                elif col in ['Porta Tot', 'LN Tot', 'BAF Tot', 'Total Ventas', 'Horas Totales']: style_str = 'background-color: #E8F8F5;'
                
                if 'SPH' in col:
                    try:
                        val = float(row[col])
                        if val >= 0.45: style_str += ' background-color: #D4EFDF; color: #145A32; font-weight: bold;'
                        elif val >= 0.35: style_str += ' background-color: #FCF3CF; color: #7D6608; font-weight: bold;'
                        else: style_str += ' background-color: #FADBD8; color: #78281F; font-weight: bold;'
                    except: pass
                styles.append(style_str)
            return styles

        st.dataframe(df_clean.style.apply(style_geex_dataframe, axis=1), use_container_width=True, hide_index=True, height=600)

# --- 3. PESTAÑA ASESORES ---
with tab_asesores:
    modo_vista = st.radio("Seleccionar Vista", ["Resumen Mensual", "Evolución Diaria / Por Asesor"], horizontal=True, key="vista_asesores_radio")
    df_as = m['df_asesores'].copy()
    
    def color_sph(val):
        try:
            v = float(val)
            if v >= 0.45: return 'background-color: #D4EFDF; color: #145A32; font-weight: bold;'
            elif v >= 0.35: return 'background-color: #FCF3CF; color: #7D6608; font-weight: bold;'
            else: return 'background-color: #FADBD8; color: #78281F; font-weight: bold;'
        except: return ''

    if df_as.empty:
        st.warning(f"No hay asesores con actividad para {mes_seleccionado}.")
    else:
        if modo_vista == "Resumen Mensual":
            filtro_sup = st.selectbox("Filtrar por Turno / Supervisor", ["Todos", "RULFO, YAMIL (Mañana)", "PUENTES, SOLEDAD (Tarde)"], key="filtro_sup_as")
            df_filtered = df_as.copy()
            if "RULFO" in filtro_sup: df_filtered = df_as[df_as['Supervisor'] == 'RULFO, YAMIL']
            elif "PUENTES" in filtro_sup: df_filtered = df_as[df_as['Supervisor'] == 'PUENTES, SOLEDAD']
                
            df_show_as = df_filtered.copy()
            df_show_as['SPH Total'] = df_show_as['SPH Total'].map(lambda x: f"{x:.2f}")
            df_show_as['SPH (baf/porta)'] = df_show_as['SPH (baf/porta)'].map(lambda x: f"{x:.2f}")
            
            tot_horas = df_filtered['Horas'].sum()
            tot_porta = df_filtered['PORTA'].sum()
            tot_baf = df_filtered['BAF'].sum()
            tot_ln = df_filtered['LN'].sum()
            tot_ventas = df_filtered['Total Ventas'].sum()
            sph_tot_calc = tot_ventas / tot_horas if tot_horas > 0 else 0
            sph_bp_calc = (tot_porta + tot_baf) / tot_horas if tot_horas > 0 else 0

            tot_row = pd.DataFrame({
                'Supervisor': ['TOTAL'], 'Asesor': [''], 'Horas': [tot_horas],
                'PORTA': [tot_porta], 'BAF': [tot_baf], 'LN': [tot_ln], 'Total Ventas': [tot_ventas],
                'SPH Total': [f"{sph_tot_calc:.2f}"], 'SPH (baf/porta)': [f"{sph_bp_calc:.2f}"]
            })
            tot_row.columns = df_show_as.columns
            df_show_as = pd.concat([df_show_as, tot_row], ignore_index=True)

            def style_asesores(row):
                styles = [''] * len(row)
                sup = row['Supervisor']
                bg_sup = '#FAF4EB' if sup == 'PUENTES, SOLEDAD' else ('#EBF3FA' if sup == 'RULFO, YAMIL' else '#D5D8DC')
                for i, col in enumerate(row.index):
                    if col == 'Supervisor': styles[i] = f'background-color: {bg_sup}; font-weight: bold;'
                    if col in ['SPH (baf/porta)', 'SPH Total']: styles[i] = color_sph(row[col])
                return styles

            st.dataframe(df_show_as.style.apply(style_asesores, axis=1), use_container_width=True, hide_index=True, height=650)
        else:
            col_f1, col_f2 = st.columns(2)
            asesor_seleccionado = col_f1.selectbox("Filtrar por Asesor", sorted(df_as['Asesor'].unique()), key="sel_asesor_evol")
            
            h_asesor = m['horas_oct'][m['horas_oct']['Asesor'] == asesor_seleccionado]
            p_asesor = m['p_dig'][m['p_dig']['Asesor'] == asesor_seleccionado]
            l_asesor = m['l_dig'][m['l_dig']['Asesor'] == asesor_seleccionado]
            b_asesor = m['b_dig'][m['b_dig']['Asesor'] == asesor_seleccionado]
            
            start_date = pd.Timestamp(year=anio_sel, month=mes_sel, day=1)
            end_date = start_date + pd.offsets.MonthEnd(1)
            dates_m = pd.date_range(start_date, end_date)
            evol_rows = []
            
            for d in dates_m:
                d_date = d.date()
                h_d = h_asesor[h_asesor['Fecha'].dt.date == d_date]['Horas'].sum()
                if h_d == 0: continue
                p_d = len(p_asesor[p_asesor['Fecha'].dt.date == d_date])
                l_d = len(l_asesor[l_asesor['Fecha'].dt.date == d_date])
                b_d = len(b_asesor[b_asesor['Fecha'].dt.date == d_date])
                tot_d = p_d + l_d + b_d
                sph_tot_d = tot_d / h_d if h_d > 0 else 0
                sph_bp_d = (p_d + b_d) / h_d if h_d > 0 else 0
                
                evol_rows.append({
                    "Día": d.strftime('%d-%b'), "Horas": int(h_d), "PORTA": p_d, "BAF": b_d,
                    "LN": l_d, "Total Ventas": tot_d, "SPH Total": f"{sph_tot_d:.2f}", "SPH (baf/porta)": f"{sph_bp_d:.2f}"
                })
                
            df_evol = pd.DataFrame(evol_rows)
            
            if not df_evol.empty:
                tot_h_as = df_evol['Horas'].sum()
                tot_p_as = df_evol['PORTA'].sum()
                tot_b_as = df_evol['BAF'].sum()
                tot_l_as = df_evol['LN'].sum()
                tot_v_as = df_evol['Total Ventas'].sum()
                sph_tot_as = tot_v_as / tot_h_as if tot_h_as > 0 else 0
                sph_bp_as = (tot_p_as + tot_b_as) / tot_h_as if tot_h_as > 0 else 0
                
                tot_row_as = pd.DataFrame({
                    "Día": ["TOTAL"], "Horas": [int(tot_h_as)], "PORTA": [tot_p_as], "BAF": [tot_b_as],
                    "LN": [tot_l_as], "Total Ventas": [tot_v_as], "SPH Total": [f"{sph_tot_as:.2f}"], "SPH (baf/porta)": [f"{sph_bp_as:.2f}"]
                })
                df_evol = pd.concat([df_evol, tot_row_as], ignore_index=True)

            def style_evol(row):
                styles = [''] * len(row)
                is_total_row = row['Día'] == 'TOTAL'
                try: is_zero = float(row['SPH (baf/porta)']) == 0.0
                except: is_zero = False
                    
                for i, col in enumerate(row.index):
                    if is_total_row:
                        styles[i] = 'background-color: #E8F8F5; font-weight: bold;'
                        if col in ['SPH (baf/porta)', 'SPH Total']: styles[i] += f" {color_sph(row[col])}"
                    elif is_zero: styles[i] = 'background-color: #FADBD8; color: #78281F; font-weight: bold;'
                    else:
                        if col in ['SPH (baf/porta)', 'SPH Total']: styles[i] = color_sph(row[col])
                return styles

            st.markdown(f"**Evolución Diaria para: {asesor_seleccionado}**")
            st.dataframe(df_evol.style.apply(style_evol, axis=1), use_container_width=True, hide_index=True, height=650)

# --- 4. PESTAÑA DATOS ---
with tab_datos:
    st.markdown("#### Ingresos de Leads por Día (Landing vs Botmaker)")
    if m['df_leads_d'].empty:
        st.info("No hay datos de leads diarios para este mes.")
    else:
        st.dataframe(m['df_leads_d'], use_container_width=True, hide_index=True, height=350)
    
    st.markdown("---")
    
    st.markdown("#### Ingresos de Leads por Hora (Turnos Mañana y Tarde)")
    df_h = m['df_leads_h'].copy()
    def get_turno(h):
        if 9 <= h <= 14: return "Turno Mañana (09-14)"
        elif 15 <= h <= 21: return "Turno Tarde (15-21)"
        else: return "Fuera de Turno"
            
    df_h['Turno'] = df_h['Hora'].apply(get_turno)
    filtro_turno = st.selectbox("Filtrar Análisis Horario", ["Todos", "Turno Mañana (09-14)", "Turno Tarde (15-21)"], key="filtro_turno_datos")
    df_h_filtered = df_h.copy()
    if filtro_turno != "Todos": df_h_filtered = df_h[df_h['Turno'] == filtro_turno]
        
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Total Ingresos Selección", f"{df_h_filtered['Total'].sum():,}")
    col_m2.metric("Total Botmaker (Turno)", f"{df_h_filtered['Botmaker'].sum():,}")
    col_m3.metric("Total Landing (Turno)", f"{df_h_filtered['Landing'].sum():,}")

    def style_hours(row):
        styles = [''] * len(row)
        h = row['Hora']
        bg = ''
        if 9 <= h <= 14: bg = 'background-color: #EBF3FA;'
        elif 15 <= h <= 21: bg = 'background-color: #FAF4EB;'
        for i, col in enumerate(row.index): styles[i] = bg
        return styles

    st.dataframe(df_h_filtered.style.apply(style_hours, axis=1), use_container_width=True, hide_index=True, height=350)
    
    st.markdown("---")
    
    def display_planes_filtered(pivot_df, title, filter_key):
        if pivot_df.empty:
            st.info(f"No hay datos para {title}")
            return
        all_cols = list(pivot_df.columns)
        selected_cols = st.multiselect(title, all_cols, default=all_cols, key=filter_key)
        if selected_cols:
            total_general_dia = pivot_df.sum(axis=1).replace(0, 1)
            total_general_mes = pivot_df.sum().sum()
            if total_general_mes == 0: total_general_mes = 1
            df_show = pd.DataFrame(index=pivot_df.index)
            tot_sel_dia = pd.Series(0, index=pivot_df.index)
            totales_dict = {}
            total_seleccionado_mes = 0
            for col in selected_cols:
                df_show[col] = pivot_df[col]
                df_show[f"% {col}"] = (pivot_df[col] / total_general_dia * 100).map(lambda x: f"{x:.1f}%")
                tot_sel_dia += pivot_df[col]
                sum_col = pivot_df[col].sum()
                totales_dict[col] = sum_col
                totales_dict[f"% {col}"] = f"{(sum_col / total_general_mes * 100):.1f}%"
                total_seleccionado_mes += sum_col
            df_show['Total Seleccionado'] = tot_sel_dia
            df_show['% Selección'] = (tot_sel_dia / total_general_dia * 100).map(lambda x: f"{x:.1f}%")
            totales_dict['Total Seleccionado'] = total_seleccionado_mes
            totales_dict['% Selección'] = f"{(total_seleccionado_mes / total_general_mes * 100):.1f}%"
            df_show.index = df_show.index.astype(str)
            df_total = pd.DataFrame(totales_dict, index=['TOTAL'])
            df_final = pd.concat([df_show, df_total])
            def style_total_row(row):
                if row.name == 'TOTAL': return ['background-color: #E8F8F5; font-weight: bold;'] * len(row)
                return [''] * len(row)
            st.dataframe(df_final.style.apply(style_total_row, axis=1), use_container_width=True, height=450)
        else:
            st.info("Selecciona al menos un ítem para visualizar.")

    sub_t1, sub_t2, sub_t3 = st.tabs(["Porta Planes", "Porta eSIM", "Planes BAF"])
    with sub_t1: display_planes_filtered(m['p_planes_pivot'], "Filtrar Planes de Portabilidad", f"f_porta_{mes_sel}")
    with sub_t2: display_planes_filtered(m['esim_pivot'], "Filtrar Tipo eSIM", f"f_esim_{mes_sel}")
    with sub_t3: display_planes_filtered(m['baf_pivot'], "Filtrar Planes BAF", f"f_baf_{mes_sel}")

# --- 5. PESTAÑA EFECTIVIDAD ---
with tab_efectividad:
    format_dict = {
        'Leads': '{:,.0f}', 'Efectividad': '{:.2%}',
        'Porta': '{:,.0f}', '% Porta': '{:.2%}',
        'LN': '{:,.0f}', '% LN': '{:.2%}',
        'BAF': '{:,.0f}', '% BAF': '{:.2%}',
        'Total Ventas': '{:,.0f}'
    }

    meta_leads_cnt = len(m['bot_m_merged'])
    goog_leads_cnt = len(m['land_m'])
    tot_leads_all = meta_leads_cnt + goog_leads_cnt

    def calc_summary_row(name, p_df, l_df, b_df, leads_base):
        p_cnt, l_cnt, b_cnt = len(p_df), len(l_df), len(b_df)
        tot_v = p_cnt + l_cnt + b_cnt
        ef = tot_v / leads_base if leads_base > 0 else 0
        return {
            'Segmento': name, 'Leads': leads_base,
            'Porta': p_cnt, '% Porta': p_cnt / tot_v if tot_v > 0 else 0,
            'LN': l_cnt, '% LN': l_cnt / tot_v if tot_v > 0 else 0,
            'BAF': b_cnt, '% BAF': b_cnt / tot_v if tot_v > 0 else 0,
            'Total Ventas': tot_v, 'Efectividad': ef
        }

    st.markdown("<div class='section-header'>Resumen General por Plataforma</div>", unsafe_allow_html=True)
    df_sum_plat = pd.DataFrame([
        calc_summary_row("Meta", m['p_dig'][m['p_dig']['BOTMAKER'].isin(m['bot_m_merged']['Anuncio_Titulo'])], m['l_dig'][m['l_dig']['BOTMAKER'].isin(m['bot_m_merged']['Anuncio_Titulo'])], m['b_dig'][m['b_dig']['BOTMAKER'].isin(m['bot_m_merged']['Anuncio_Titulo'])], meta_leads_cnt),
        calc_summary_row("Google", m['p_dig'][m['p_dig']['LANDING'].isin(m['land_m']['Fuente'])], m['l_dig'][m['l_dig']['LANDING'].isin(m['land_m']['Fuente'])], m['b_dig'][m['b_dig']['LANDING'].isin(m['land_m']['Fuente'])], goog_leads_cnt),
        calc_summary_row("TOTAL", m['p_dig'], m['l_dig'], m['b_dig'], tot_leads_all)
    ]).set_index('Segmento')[['Leads', 'Porta', '% Porta', 'LN', '% LN', 'BAF', '% BAF', 'Total Ventas', 'Efectividad']]
    st.dataframe(df_sum_plat.style.format(format_dict), use_container_width=True)

    st.markdown("<div class='section-header'>Resumen General por Herramienta</div>", unsafe_allow_html=True)
    df_sum_tool = pd.DataFrame([
        calc_summary_row("Botmaker", m['p_dig'][m['p_dig']['BOTMAKER'].isin(m['bot_m_merged']['Anuncio_Titulo'])], m['l_dig'][m['l_dig']['BOTMAKER'].isin(m['bot_m_merged']['Anuncio_Titulo'])], m['b_dig'][m['b_dig']['BOTMAKER'].isin(m['bot_m_merged']['Anuncio_Titulo'])], meta_leads_cnt),
        calc_summary_row("Landing", m['p_dig'][m['p_dig']['LANDING'].isin(m['land_m']['Fuente'])], m['l_dig'][m['l_dig']['LANDING'].isin(m['land_m']['Fuente'])], m['b_dig'][m['b_dig']['LANDING'].isin(m['land_m']['Fuente'])], goog_leads_cnt),
        calc_summary_row("TOTAL", m['p_dig'], m['l_dig'], m['b_dig'], tot_leads_all)
    ]).set_index('Segmento')[['Leads', 'Porta', '% Porta', 'LN', '% LN', 'BAF', '% BAF', 'Total Ventas', 'Efectividad']]
    st.dataframe(df_sum_tool.style.format(format_dict), use_container_width=True)

    st.markdown("<div class='section-header'>Resumen General por Producto de Campaña</div>", unsafe_allow_html=True)
    porta_leads_cnt = len(m['bot_m_merged'][m['bot_m_merged']['Producto']=='PORTA']) + len(m['land_m'][m['land_m']['Producto']=='PORTA'])
    baf_leads_cnt = len(m['bot_m_merged'][m['bot_m_merged']['Producto']=='BAF']) + len(m['land_m'][m['land_m']['Producto']=='BAF'])
    
    df_sum_prod = pd.DataFrame([
        calc_summary_row("Porta", m['p_dig'], m['l_dig'], m['b_dig'], porta_leads_cnt if porta_leads_cnt>0 else 1),
        calc_summary_row("BAF", m['p_dig'], m['l_dig'], m['b_dig'], baf_leads_cnt if baf_leads_cnt>0 else 1),
        calc_summary_row("TOTAL", m['p_dig'], m['l_dig'], m['b_dig'], tot_leads_all)
    ]).set_index('Segmento')[['Leads', 'Porta', '% Porta', 'LN', '% LN', 'BAF', '% BAF', 'Total Ventas', 'Efectividad']]
    st.dataframe(df_sum_prod.style.format(format_dict), use_container_width=True)

    st.markdown("---")

    st.markdown("### Filtro para Cuadros de Anuncios")
    prod_filtro_ads = st.selectbox("Filtrar por Producto Vendido (Aplica a los cuadros de anuncios abajo)", ["Todos", "PORTA", "LN", "BAF"], key="filtro_prod_ads")
    
    def get_filtered_sales_ads():
        if prod_filtro_ads == "PORTA": return m['p_dig'], m['p_dig'].iloc[0:0], m['p_dig'].iloc[0:0]
        if prod_filtro_ads == "LN": return m['l_dig'].iloc[0:0], m['l_dig'], m['l_dig'].iloc[0:0]
        if prod_filtro_ads == "BAF": return m['b_dig'].iloc[0:0], m['b_dig'].iloc[0:0], m['b_dig']
        return m['p_dig'], m['l_dig'], m['b_dig']

    p_fa, l_fa, b_fa = get_filtered_sales_ads()

    def get_ad_table(sub_df, ad_col, sales_col):
        ads_sz = sub_df.groupby(ad_col).size()
        p_sub = p_fa.groupby(sales_col).size()
        l_sub = l_fa.groupby(sales_col).size()
        b_sub = b_fa.groupby(sales_col).size()
        
        df = pd.DataFrame({'Leads': ads_sz, 'Porta': p_sub, 'LN': l_sub, 'BAF': b_sub}).fillna(0)
        df['Total Ventas'] = df['Porta'] + df['LN'] + df['BAF']
        df = df[(df['Leads'] > 0) | (df['Total Ventas'] > 0)].copy()
        df['Efectividad'] = (df['Total Ventas'] / df['Leads']).replace([np.inf, -np.inf], 0).fillna(0)
        df['% Porta'] = (df['Porta'] / df['Total Ventas']).replace([np.inf, -np.inf], 0).fillna(0)
        df['% LN'] = (df['LN'] / df['Total Ventas']).replace([np.inf, -np.inf], 0).fillna(0)
        df['% BAF'] = (df['BAF'] / df['Total Ventas']).replace([np.inf, -np.inf], 0).fillna(0)
        return df[['Leads', 'Efectividad', 'Porta', '% Porta', 'LN', '% LN', 'BAF', '% BAF', 'Total Ventas']].sort_values(by='Efectividad', ascending=False)

    st.markdown("<div class='section-header'>Anuncios por Plataforma (Ordenados por Mayor Efectividad)</div>", unsafe_allow_html=True)
    col_meta, col_goog = st.columns(2)
    with col_meta:
        st.markdown("#### Meta (Botmaker)")
        st.dataframe(get_ad_table(m['bot_m_merged'], 'Anuncio_Titulo', 'BOTMAKER').style.format(format_dict), use_container_width=True, height=350)
    with col_goog:
        st.markdown("#### Google (Landing)")
        st.dataframe(get_ad_table(m['land_m'], 'Fuente', 'LANDING').style.format(format_dict), use_container_width=True, height=350)

    st.markdown("<div class='section-header'>Anuncios por Producto de Campaña (Ordenados por Mayor Efectividad)</div>", unsafe_allow_html=True)
    col_porta_c, col_baf_c = st.columns(2)

    def get_campaign_ad_table(prod_name):
        b_sub = m['bot_m_merged'][m['bot_m_merged']['Producto'].astype(str).str.upper() == prod_name]
        l_sub = m['land_m'][m['land_m']['Producto'].astype(str).str.upper() == prod_name]
        ads_all = pd.concat([b_sub.groupby('Anuncio_Titulo').size(), l_sub.groupby('Fuente').size()]).groupby(level=0).sum()
        
        p_sub = p_fa.groupby('BOTMAKER' if 'BOTMAKER' in p_fa.columns else 'LANDING').size()
        l_sub_v = l_fa.groupby('BOTMAKER' if 'BOTMAKER' in l_fa.columns else 'LANDING').size()
        b_sub_v = b_fa.groupby('BOTMAKER' if 'BOTMAKER' in b_fa.columns else 'LANDING').size()
        
        df = pd.DataFrame({'Leads': ads_all, 'Porta': p_sub, 'LN': l_sub_v, 'BAF': b_sub_v}).fillna(0)
        df['Total Ventas'] = df['Porta'] + df['LN'] + df['BAF']
        df = df[(df['Leads'] > 0) | (df['Total Ventas'] > 0)].copy()
        df['Efectividad'] = (df['Total Ventas'] / df['Leads']).replace([np.inf, -np.inf], 0).fillna(0)
        df['% Porta'] = (df['Porta'] / df['Total Ventas']).replace([np.inf, -np.inf], 0).fillna(0)
        df['% LN'] = (df['LN'] / df['Total Ventas']).replace([np.inf, -np.inf], 0).fillna(0)
        df['% BAF'] = (df['BAF'] / df['Total Ventas']).replace([np.inf, -np.inf], 0).fillna(0)
        return df[['Leads', 'Efectividad', 'Porta', '% Porta', 'LN', '% LN', 'BAF', '% BAF', 'Total Ventas']].sort_values(by='Efectividad', ascending=False)

    with col_porta_c:
        st.markdown("#### Campaña: Porta")
        st.dataframe(get_campaign_ad_table('PORTA').style.format(format_dict), use_container_width=True, height=350)
    with col_baf_c:
        st.markdown("#### Campaña: BAF")
        st.dataframe(get_campaign_ad_table('BAF').style.format(format_dict), use_container_width=True, height=350)

# --- 6. PESTAÑA PRESUPUESTO ---
with tab_presupuesto:
    st.markdown("<div class='section-header'>Presupuesto Diario y Periodos de Facturación</div>", unsafe_allow_html=True)
    
    if mes_sel == 10:
        st.info("Período de Facturación Octubre: Meta (02/10 al 28/10) | Google (01/10 al 31/10)")
    else:
        st.info("Período de Facturación Noviembre: Meta (29/10 al 26/11) | Google (01/11 al 26/11)")

    df_master = m['df_pres_master'].copy()

    def append_total_row(df, money_cols, sum_cols):
        if df.empty: return df
        total_row = {'Fecha': 'TOTAL'}
        for c in df.columns:
            if c in money_cols or c in sum_cols:
                total_row[c] = df[c].sum()
        
        if 'Meta ($)' in df.columns:
            tot_pres = total_row.get('Total ($)', 0)
            total_row['% Meta ($)'] = total_row.get('Meta ($)', 0) / tot_pres if tot_pres > 0 else 0
            total_row['% Google ($)'] = total_row.get('Google ($)', 0) / tot_pres if tot_pres > 0 else 0
            total_row['% Baf ($)'] = total_row.get('Baf ($)', 0) / tot_pres if tot_pres > 0 else 0
            total_row['% Porta ($)'] = total_row.get('Porta ($)', 0) / tot_pres if tot_pres > 0 else 0
            
        elif 'cpl meta' in df.columns:
            tot_l = total_row.get('Total Leads', 0)
            total_row['% Meta (Leads)'] = total_row.get('Meta', 0) / tot_l if tot_l > 0 else 0
            total_row['% Google (Leads)'] = total_row.get('Google', 0) / tot_l if tot_l > 0 else 0
            tot_meta_dinero = df['Meta ($)'].sum() if 'Meta ($)' in df.columns else df_master['Meta ($)'].sum()
            tot_goog_dinero = df['Google ($)'].sum() if 'Google ($)' in df.columns else df_master['Google ($)'].sum()
            total_row['cpl meta'] = tot_meta_dinero / total_row.get('Meta', 1) if total_row.get('Meta', 1) > 0 else 0
            total_row['cpl google'] = tot_goog_dinero / total_row.get('Google', 1) if total_row.get('Google', 1) > 0 else 0

        elif 'cpl baf' in df.columns:
            tot_l = total_row.get('Total Leads', 0)
            total_row['% BAF'] = total_row.get('BAF', 0) / tot_l if tot_l > 0 else 0
            total_row['% Porta'] = total_row.get('Porta', 0) / tot_l if tot_l > 0 else 0
            tot_baf_dinero = df['Baf ($)'].sum() if 'Baf ($)' in df.columns else df_master['Baf ($)'].sum()
            tot_porta_dinero = df['Porta ($)'].sum() if 'Porta ($)' in df.columns else df_master['Porta ($)'].sum()
            total_row['cpl baf'] = tot_baf_dinero / total_row.get('BAF', 1) if total_row.get('BAF', 1) > 0 else 0
            total_row['cpl porta'] = tot_porta_dinero / total_row.get('Porta', 1) if total_row.get('Porta', 1) > 0 else 0
            
        return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    def style_tot(row):
        return ['background-color: #E8F8F5; font-weight: bold;' if row['Fecha'] == 'TOTAL' else ''] * len(row)

    st.markdown("<div class='sub-header'>1. PRESUPUESTO PLATAFORMA Y PRODUCTO</div>", unsafe_allow_html=True)
    df_gastos = df_master[['Fecha', 'Total ($)', 'Meta ($)', '% Meta ($)', 'Google ($)', '% Google ($)', 'Baf ($)', '% Baf ($)', 'Porta ($)', '% Porta ($)']].copy()
    # Filtrar solo días con presupuesto o datos
    df_gastos = df_gastos[df_gastos['Total ($)'] > 0].copy()
    if df_gastos.empty:
        st.info("No hay datos de presupuesto registrados aún para el período seleccionado.")
    else:
        df_gastos = append_total_row(df_gastos, ['Total ($)', 'Meta ($)', 'Google ($)', 'Baf ($)', 'Porta ($)'], [])
        st.dataframe(df_gastos.style.format({
            'Total ($)': '${:,.2f}', 'Meta ($)': '${:,.2f}', 'Google ($)': '${:,.2f}', 'Baf ($)': '${:,.2f}', 'Porta ($)': '${:,.2f}',
            '% Meta ($)': '{:.0%}', '% Google ($)': '{:.0%}', '% Baf ($)': '{:.0%}', '% Porta ($)': '{:.0%}'
        }).apply(style_tot, axis=1), use_container_width=True, height=450)

    st.markdown("<div class='sub-header'>2. REALES Y CPL (Por Plataforma)</div>", unsafe_allow_html=True)
    df_reales_plat = df_master[['Fecha', 'Total Leads', 'Botmaker no se identifica', 'Meta', 'Landing Meta', 'Botmaker Meta', '% Meta (Leads)', 'Google', 'Landing Google', '% Google (Leads)', 'cpl meta', 'cpl google']].copy()
    df_reales_plat['cpl meta'] = df_reales_plat['cpl meta'].replace([np.inf, -np.inf], 0)
    df_reales_plat['cpl google'] = df_reales_plat['cpl google'].replace([np.inf, -np.inf], 0)
    
    # Filtrar solo días con leads registrados
    df_reales_plat = df_reales_plat[df_reales_plat['Total Leads'] > 0].copy()
    if df_reales_plat.empty:
        st.info("No hay datos de leads por plataforma registrados aún.")
    else:
        sum_cols_plat = ['Total Leads', 'Botmaker no se identifica', 'Meta', 'Landing Meta', 'Botmaker Meta', 'Google', 'Landing Google']
        df_reales_plat = append_total_row(df_reales_plat, [], sum_cols_plat)
        st.dataframe(df_reales_plat.style.format({
            'cpl meta': '${:,.2f}', 'cpl google': '${:,.2f}',
            '% Meta (Leads)': '{:.0%}', '% Google (Leads)': '{:.0%}'
        }).apply(style_tot, axis=1), use_container_width=True, height=450)

    st.markdown("<div class='sub-header'>3. REALES Y CPL (Por Producto)</div>", unsafe_allow_html=True)
    df_reales_prod = df_master[['Fecha', 'Total Leads', 'Botmaker no se identifica', 'BAF', 'Landing BAF', 'Botmaker BAF', '% BAF', 'Porta', 'Landing Porta', 'Botmaker Porta', '% Porta', 'cpl baf', 'cpl porta']].copy()
    df_reales_prod['cpl baf'] = df_reales_prod['cpl baf'].replace([np.inf, -np.inf], 0)
    df_reales_prod['cpl porta'] = df_reales_prod['cpl porta'].replace([np.inf, -np.inf], 0)
    
    # Filtrar solo días con leads registrados
    df_reales_prod = df_reales_prod[df_reales_prod['Total Leads'] > 0].copy()
    if df_reales_prod.empty:
        st.info("No hay datos de leads por producto registrados aún.")
    else:
        sum_cols_prod = ['Total Leads', 'Botmaker no se identifica', 'BAF', 'Landing BAF', 'Botmaker BAF', 'Porta', 'Landing Porta', 'Botmaker Porta']
        df_reales_prod = append_total_row(df_reales_prod, [], sum_cols_prod)
        st.dataframe(df_reales_prod.style.format({
            'cpl baf': '${:,.2f}', 'cpl porta': '${:,.2f}',
            '% BAF': '{:.0%}', '% Porta': '{:.0%}'
        }).apply(style_tot, axis=1), use_container_width=True, height=450)