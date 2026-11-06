# --- PRESUPUESTO DIARIO, REALES Y CPL ---
    m_ar = meta_raw.copy()
    m_ar['Pais_Clean'] = m_ar['Pais'].astype(str).str.strip().str.upper()
    m_ar = m_ar[m_ar['Pais_Clean'].str.startswith('AR')].copy()
    m_ar['Fecha_Clean'] = pd.to_datetime(m_ar['Día'], errors='coerce').dt.date
    m_ar['Prod_Clean'] = m_ar['Producto'].astype(str).str.strip().str.upper()
    
    col_spend_meta = 'Monto gastado (ARS)' if 'Monto gastado (ARS)' in m_ar.columns else 'Monto gastado'
    m_ar['Gasto_Clean'] = safe_numeric_spend(m_ar[col_spend_meta])
    leads_col_meta = m_ar['Resultados'] if 'Resultados' in m_ar.columns else pd.Series(0, index=m_ar.index)
    m_ar['Leads_Plataforma'] = pd.to_numeric(leads_col_meta, errors='coerce').fillna(0)

    g_ar = google_raw.copy()
    g_ar['Pais_Clean'] = g_ar['Pais'].astype(str).str.strip().str.upper()
    g_ar = g_ar[g_ar['Pais_Clean'].str.startswith('AR')].copy()
    g_ar['Fecha_Clean'] = pd.to_datetime(g_ar['Día'], errors='coerce').dt.date
    g_ar['Prod_Clean'] = g_ar['Producto'].astype(str).str.strip().str.upper()
    
    col_spend_goog = 'Costo num' if 'Costo num' in g_ar.columns else ('Costo' if 'Costo' in g_ar.columns else 'Gasto')
    g_ar['Gasto_Clean'] = safe_numeric_spend(g_ar[col_spend_goog])
    if 'Conversiones num' in g_ar.columns:
        leads_col_goog = g_ar['Conversiones num']
    elif 'Conversiones' in g_ar.columns:
        leads_col_goog = g_ar['Conversiones']
    else:
        leads_col_goog = pd.Series(0, index=g_ar.index)
    g_ar['Leads_Plataforma'] = pd.to_numeric(leads_col_goog, errors='coerce').fillna(0)

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
    b_all['Producto_Ad'] = b_all['Producto'].astype(str).str.strip().str.upper()

    # Obtener dinámicamente las fechas disponibles en los datos de Meta y Google para el mes seleccionado
    all_dates_meta = m_ar[pd.to_datetime(m_ar['Fecha_Clean']).dt.month == mes]['Fecha_Clean'].unique()
    all_dates_goog = g_ar[pd.to_datetime(g_ar['Fecha_Clean']).dt.month == mes]['Fecha_Clean'].unique()
    all_valid_dates = sorted(list(set(list(all_dates_meta) + list(all_dates_goog))))
    
    if not all_valid_dates:
        # Fallback al mes calendario completo si no hay filtro estricto
        start_d = pd.Timestamp(year=anio, month=mes, day=1)
        end_d = start_d + pd.offsets.MonthEnd(1)
        dates_presupuesto = pd.date_range(start_d, end_d)
    else:
        dates_presupuesto = pd.to_datetime(all_valid_dates)

    pres_rows = []
    
    for d in dates_presupuesto:
        d_val = d.date() if hasattr(d, 'date') else d
        
        m_day = m_ar[m_ar['Fecha_Clean'] == d_val]
        p_meta = m_day['Gasto_Clean'].sum()
        p_m_baf = m_day[m_day['Prod_Clean'].str.contains('BAF')]['Gasto_Clean'].sum()
        p_m_porta = m_day[m_day['Prod_Clean'].str.contains('PORTA')]['Gasto_Clean'].sum()
        leads_plat_meta = m_day['Leads_Plataforma'].sum()
            
        g_day = g_ar[g_ar['Fecha_Clean'] == d_val]
        p_goog = g_day['Gasto_Clean'].sum()
        p_g_baf = g_day[g_day['Prod_Clean'].str.contains('BAF')]['Gasto_Clean'].sum()
        p_g_porta = g_day[g_day['Prod_Clean'].str.contains('PORTA')]['Gasto_Clean'].sum()
        leads_plat_goog = g_day['Leads_Plataforma'].sum()

        tot_pres = p_meta + p_goog
        p_baf = p_m_baf + p_g_baf
        p_porta = p_m_porta + p_g_porta
        tot_leads_plataforma = leads_plat_meta + leads_plat_goog

        b_day = b_all[pd.to_datetime(b_all['Fecha Sesión'], dayfirst=True, errors='coerce').dt.date == d_val]
        l_day = l_all[pd.to_datetime(l_all['Fecha Ingreso'], dayfirst=True, errors='coerce').dt.date == d_val]

        bot_no_id = b_day['Plataforma'].isin(['NAN', 'NONE', '']).sum() + b_day['Plataforma'].isna().sum()
        bot_m_meta = len(b_day[b_day['Plataforma'] == 'META'])
        land_m_meta = len(l_day[l_day['Plataforma'] == 'META'])
        bot_m_goog = len(b_day[b_day['Plataforma'] == 'GOOGLE'])
        land_m_goog = len(l_day[l_day['Plataforma'] == 'GOOGLE'])

        leads_meta = bot_no_id + bot_m_meta + land_m_meta
        leads_goog = bot_m_goog + land_m_goog
        leads_tot_reales = leads_meta + leads_goog

        bot_m_baf = len(b_day[b_day['Producto_Ad'] == 'BAF'])
        land_m_baf = len(l_day[l_day['Producto'] == 'BAF'])
        leads_baf = bot_no_id + bot_m_baf + land_m_baf
        
        bot_m_porta = len(b_day[b_day['Producto_Ad'] == 'PORTA'])
        land_m_porta = len(l_day[l_day['Producto'] == 'PORTA'])
        leads_porta = bot_m_porta + land_m_porta

        cpl_general = tot_pres / leads_tot_reales if leads_tot_reales > 0 else 0
        cpl_meta = p_meta / leads_meta if leads_meta > 0 else 0
        cpl_goog = p_goog / leads_goog if leads_goog > 0 else 0
        cpl_baf = p_baf / leads_baf if leads_baf > 0 else 0
        cpl_porta = p_porta / leads_porta if leads_porta > 0 else 0

        diferencia_leads = (leads_tot_reales / tot_leads_plataforma - 1) if tot_leads_plataforma > 0 else 0

        formatted_date = pd.to_datetime(d_val).strftime('%d/%m/%Y')

        pres_rows.append({
            'Fecha': formatted_date,
            'Total ($)': tot_pres,
            'Meta ($)': p_meta, '% Meta': p_meta / tot_pres if tot_pres > 0 else 0,
            'Google ($)': p_goog, '% Google': p_goog / tot_pres if tot_pres > 0 else 0,
            'Baf ($)': p_baf, '% Baf': p_baf / tot_pres if tot_pres > 0 else 0,
            'Porta ($)': p_porta, '% Porta': p_porta / tot_pres if tot_pres > 0 else 0,
            
            'Plataforma (Leads)': tot_leads_plataforma,
            'Reales (Leads)': leads_tot_reales,
            'Diferencia': diferencia_leads,
            'CPL': cpl_general,

            'Botmaker no se identifica': bot_no_id,
            'Meta': leads_meta, 'Landing Meta': land_m_meta, 'Botmaker Meta': bot_m_meta, '% Meta (L)': leads_meta / leads_tot_reales if leads_tot_reales > 0 else 0,
            'Google': leads_goog, 'Landing Google': land_m_goog, 'Botmaker Google': bot_m_goog, '% Google (L)': leads_goog / leads_tot_reales if leads_tot_reales > 0 else 0,
            
            'BAF': leads_baf, 'Landing BAF': land_m_baf, 'Botmaker BAF': bot_m_baf, '% BAF': leads_baf / leads_tot_reales if leads_tot_reales > 0 else 0,
            'Porta': leads_porta, 'Landing Porta': land_m_porta, 'Botmaker Porta': bot_m_porta, '% Porta': leads_porta / leads_tot_reales if leads_tot_reales > 0 else 0,
            
            'cpl meta': cpl_meta, 'cpl google': cpl_goog,
            'cpl baf': cpl_baf, 'cpl porta': cpl_porta
        })
        
    df_pres_master = pd.DataFrame(pres_rows)