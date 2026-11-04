import streamlit as st
import pandas as pd
import numpy as np

# ==========================================================
# 0. PARÁMETROS 
# ==========================================================
ARCHIVO = "DASHHHHH.xlsx"                       
SUPERVISORES_DIGITAL = ["RULFO, YAMIL", "PUENTES, SOLEDAD"]
TURNO = {"RULFO, YAMIL": "Mañana", "PUENTES, SOLEDAD": "Tarde"}
ESTADO_EXCLUIDO = "NO APROBADO"                 
FERIADOS = pd.to_datetime([                     
    "2026-10-12", "2026-11-23", "2026-12-07", "2026-12-08", "2026-12-25"])
SPH_VERDE, SPH_AMARILLO = 0.45, 0.35            

PERIODOS_FACTURACION = {
    (2026, 10): {"META": ("2026-10-02", "2026-10-28"), "GOOGLE": ("2026-10-01", "2026-10-31")},
    (2026, 11): {"META": ("2026-10-29", "2026-11-26"), "GOOGLE": ("2026-11-01", "2026-11-26")},
}

SIN_ID_SUMA_A_PLATAFORMA = "META"
SIN_ID_SUMA_A_PRODUCTO = "BAF"

BRAND = {
    "primario": "#0B192C",      
    "acento": "#44B3E1",        
    "acento2": "#7030A0",       
    "manana": "#EBF3FA", "tarde": "#FAF4EB", "total": "#E8F8F5", "fin_semana": "#EAEAEA",
    "ok_bg": "#D4EFDF", "ok_tx": "#145A32",
    "med_bg": "#FCF3CF", "med_tx": "#7D6608",
    "bad_bg": "#FADBD8", "bad_tx": "#78281F",
}
MESES_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio",
            8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}

st.set_page_config(page_title="Dashboard Geex", layout="wide", initial_sidebar_state="collapsed")
st.markdown(f"""
    <style>
    .block-container {{ padding-top: 1rem !important; }}
    h1 {{ color: {BRAND['primario']}; }}
    .metric-card {{
        background-color: #FFFFFF; padding: 8px; border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1); text-align: center; margin-bottom: 10px;
        border: 1px solid #E0E0E0; border-top: 4px solid {BRAND['acento']}; white-space: nowrap;
    }}
    .metric-value {{ font-size: 19px; font-weight: bold; color: {BRAND['primario']}; }}
    .metric-label {{ font-size: 9px; color: #555555; text-transform: uppercase; font-weight: 700; letter-spacing: 0px; }}
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] {{ color: {BRAND['primario']}; border-bottom: 3px solid {BRAND['acento']}; }}
    .stTabs [data-baseweb="tab-highlight"] {{ background-color: {BRAND['acento']}; }}
    </style>
""", unsafe_allow_html=True)

def card(label, value): return f"<div class='metric-card'><div class='metric-label'>{label}</div><div class='metric-value'>{value}</div></div>"
def style_sph(v):
    try: v = float(v)
    except (TypeError, ValueError): return ""
    if v >= SPH_VERDE: return f"background-color: {BRAND['ok_bg']}; color: {BRAND['ok_tx']}; font-weight: bold;"
    if v >= SPH_AMARILLO: return f"background-color: {BRAND['med_bg']}; color: {BRAND['med_tx']}; font-weight: bold;"
    return f"background-color: {BRAND['bad_bg']}; color: {BRAND['bad_tx']}; font-weight: bold;"
def limpio(s): return s.astype(str).str.strip().where(~s.astype(str).str.strip().isin(["", "nan", "None", "NaT"]), np.nan)

# ==========================================================
# 1. CARGA DE DATOS
# ==========================================================
@st.cache_data
def cargar_datos():
    xls = pd.ExcelFile(ARCHIVO)
    d = {n: pd.read_excel(xls, h) for n, h in [
        ("horas", "RAW_horas"), ("porta", "RAW_portaAR"), ("ln", "RAW_vozAR"), ("baf", "RAW_baf"),
        ("botmaker", "RAW_botmaker"), ("bot_an", "RAW_botmaker_anuncios"), ("landing", "RAW_landing"),
        ("meta", "RAW_meta"), ("google", "RAW_google")]}
    d["horas"]["Fecha"] = pd.to_datetime(d["horas"]["Fecha"], dayfirst=True, errors="coerce")
    d["porta"]["Fecha"] = pd.to_datetime(d["porta"]["Fecha inicio"], dayfirst=True, errors="coerce")
    d["ln"]["Fecha"] = pd.to_datetime(d["ln"]["Fecha Inicio"], dayfirst=True, errors="coerce")
    d["baf"]["Fecha"] = pd.to_datetime(d["baf"]["Fecha Inicio"], dayfirst=True, errors="coerce")
    d["botmaker"]["Fecha"] = pd.to_datetime(d["botmaker"]["Fecha Sesión"], dayfirst=True, errors="coerce")
    d["landing"]["Fecha"] = pd.to_datetime(d["landing"]["Fecha Ingreso"], dayfirst=True, errors="coerce")
    d["meta"]["Fecha"] = pd.to_datetime(d["meta"]["Día"], errors="coerce")
    d["google"]["Fecha"] = pd.to_datetime(d["google"]["Día"], errors="coerce")
    
    h = d["horas"].dropna(subset=["Fecha", "Usuario"]).copy()
    h["Horas"] = pd.to_numeric(h["Horas"], errors="coerce").fillna(0)
    d["horas"] = h.sort_values("Horas", ascending=False).drop_duplicates(["Fecha", "Usuario"]).sort_values("Fecha")
    return d

try: DATOS = cargar_datos()
except Exception as e: st.error(f"No pude leer el archivo: {e}"); st.stop()

_h = DATOS["horas"]
_h = _h[_h["Supervisor"].isin(SUPERVISORES_DIGITAL) & (_h["Horas"] > 0)]
meses_disponibles = {f"{MESES_ES[mm]} {aa}": (aa, mm) for aa, mm in sorted(set(zip(_h["Fecha"].dt.year, _h["Fecha"].dt.month)))}
if not meses_disponibles: st.error("No hay horas cargadas en RAW_horas."); st.stop()

st.title("Dashboard Geex")
st.markdown("<br>", unsafe_allow_html=True)
nombres_meses = list(meses_disponibles.keys())
mes_seleccionado = st.selectbox("Filtro de Mes", nombres_meses, index=len(nombres_meses) - 1)
anio_sel, mes_sel = meses_disponibles[mes_seleccionado]

# ==========================================================
# 2. CÁLCULOS DEL MES SELECCIONADO
# ==========================================================
@st.cache_data
def calcular_metricas(anio, mes):
    D = cargar_datos()
    en_mes = lambda f: (f.dt.month == mes) & (f.dt.year == anio)

    def ventas_validas(df, col_estado, producto):
        ok = df[col_estado].astype(str).str.strip().str.upper() != ESTADO_EXCLUIDO
        out = df[ok & df["Supervisor"].isin(SUPERVISORES_DIGITAL) & en_mes(df["Fecha"])].copy()
        out["Producto_venta"] = producto
        return out

    p_dig, l_dig, b_dig = ventas_validas(D["porta"], "Estado actual", "PORTA"), ventas_validas(D["ln"], "Estado Actual", "LN"), ventas_validas(D["baf"], "Estado Actual", "BAF")

    bot_m = D["botmaker"][(D["botmaker"]["PAIS"] == "AR") & en_mes(D["botmaker"]["Fecha"])].copy()
    land_all = D["landing"][(D["landing"]["Pais"] == "AR") & en_mes(D["landing"]["Fecha"])].copy()
    total_leads = len(bot_m) + len(land_all)

    horas_m = D["horas"][en_mes(D["horas"]["Fecha"]) & D["horas"]["Supervisor"].isin(SUPERVISORES_DIGITAL)].copy()
    h_totales, v_porta, v_ln, v_baf = horas_m["Horas"].sum(), len(p_dig), len(l_dig), len(b_dig)
    v_totales = v_porta + v_ln + v_baf

    dates_m = pd.date_range(pd.Timestamp(year=anio, month=mes, day=1), periods=1, freq="MS")
    dates_m = pd.date_range(dates_m[0], dates_m[0] + pd.offsets.MonthEnd(1))
    daily_rows = []
    for d in dates_m:
        dd = d.date()
        h_d = horas_m[horas_m["Fecha"].dt.date == dd]
        p_d, l_d, b_d = [x[x["Fecha"].dt.date == dd] for x in (p_dig, l_dig, b_dig)]
        
        leads_dia = int((bot_m["Fecha"].dt.date == dd).sum() + (land_all["Fecha"].dt.date == dd).sum())
        if h_d["Horas"].sum() == 0 and len(p_d) + len(l_d) + len(b_d) == 0 and leads_dia == 0: continue
            
        no_laborable = d.weekday() >= 5 or d in FERIADOS
        fila = {"Día": d.strftime("%d-%b"), "_No_Laborable": no_laborable}
        tot = {"Porta": 0, "LN": 0, "BAF": 0, "Horas": 0}
        for sup, suf in (("RULFO, YAMIL", "M"), ("PUENTES, SOLEDAD", "T")):
            h_s = h_d[h_d["Supervisor"] == sup]
            pres, hs = int(h_s[h_s["Horas"] > 0]["Usuario"].nunique()), int(h_s["Horas"].sum())
            po, ln_, ba = int((p_d["Supervisor"] == sup).sum()), int((l_d["Supervisor"] == sup).sum()), int((b_d["Supervisor"] == sup).sum())
            tt = po + ln_ + ba
            fila.update({f"Presentes {suf}": pres or None, f"Horas {suf}": hs or None, f"Porta {suf}": po, f"LN {suf}": ln_, f"BAF {suf}": ba, f"Total {suf}": tt, f"SPH {suf}": f"{tt / hs:.2f}" if hs > 0 else "0.00", f"SPH BAF+Porta {suf}": f"{(po + ba) / hs:.2f}" if hs > 0 else "0.00"})
            tot["Porta"] += po; tot["LN"] += ln_; tot["BAF"] += ba; tot["Horas"] += hs
        tv = tot["Porta"] + tot["LN"] + tot["BAF"]
        fila.update({"Porta Tot": tot["Porta"], "LN Tot": tot["LN"], "BAF Tot": tot["BAF"], "Total Ventas": tv, "Horas Totales": tot["Horas"], "SPH Total": f"{tv / tot['Horas']:.2f}" if tot["Horas"] > 0 else "0.00", "SPH BAF+Porta": f"{(tot['Porta'] + tot['BAF']) / tot['Horas']:.2f}" if tot["Horas"] > 0 else "0.00", "Leads Totales": leads_dia, "Efectividad": f"{(tv / leads_dia if leads_dia else 0):.2%}"})
        daily_rows.append(fila)
    df_daily = pd.DataFrame(daily_rows)
    for c in [c for c in df_daily.columns if c.startswith("Presentes") or c in ("Horas M", "Horas T")]: df_daily[c] = df_daily[c].astype("Int64")

    df_as = horas_m.groupby(["Supervisor", "Asesor"])["Horas"].sum().reset_index()
    for nombre, df_v in (("PORTA", p_dig), ("LN", l_dig), ("BAF", b_dig)): df_as = df_as.merge(df_v.groupby("Asesor").size().reset_index(name=nombre), on="Asesor", how="left")
    df_as = df_as.fillna(0)
    for c in ["Horas", "PORTA", "LN", "BAF"]: df_as[c] = df_as[c].astype(int)
    
    df_as = df_as[df_as["Horas"] > 0].copy()
    df_as["Total Ventas"] = df_as["PORTA"] + df_as["LN"] + df_as["BAF"]
    df_as["SPH Total"] = (df_as["Total Ventas"] / df_as["Horas"].replace(0, np.nan)).fillna(0)
    df_as["SPH (baf/porta)"] = ((df_as["PORTA"] + df_as["BAF"]) / df_as["Horas"].replace(0, np.nan)).fillna(0)
    df_as = df_as.sort_values("SPH (baf/porta)", ascending=False).reset_index(drop=True)

    bot_d = bot_m.groupby(bot_m["Fecha"].dt.date).size().reset_index(name="Botmaker")
    land_d = land_all.groupby(land_all["Fecha"].dt.date).size().reset_index(name="Landing")
    df_leads_d = pd.merge(bot_d, land_d, on="Fecha", how="outer").fillna(0).sort_values("Fecha")
    if not df_leads_d.empty:
        df_leads_d["Día"] = pd.to_datetime(df_leads_d["Fecha"]).dt.strftime("%d-%b")
        df_leads_d["Total"] = df_leads_d["Botmaker"] + df_leads_d["Landing"]
        df_leads_d = df_leads_d[["Día", "Botmaker", "Landing", "Total"]]

    bot_m["Hora"] = pd.to_numeric(bot_m["Hora Sesión"], errors="coerce").fillna(0).astype(int)
    land_all["Hora"] = land_all["Fecha"].dt.hour
    df_leads_h = pd.merge(bot_m.groupby("Hora").size().reset_index(name="Botmaker"), land_all.groupby("Hora").size().reset_index(name="Landing"), on="Hora", how="outer").fillna(0)
    df_leads_h["Total"] = df_leads_h["Botmaker"] + df_leads_h["Landing"]
    df_leads_h = df_leads_h.sort_values("Hora").reset_index(drop=True)

    # --- PLANES ---
    p_dig["Día_Num"] = p_dig["Fecha"].dt.day
    b_dig["Día_Num"] = b_dig["Fecha"].dt.day
    def pivote(df, col):
        if col not in df.columns: return pd.DataFrame()
        pv = pd.pivot_table(df, index="Día_Num", columns=col, aggfunc="size", fill_value=0)
        return pv[pv.sum().sort_values(ascending=False).index] if not pv.empty else pv
    p_planes_pivot, esim_pivot, baf_pivot = pivote(p_dig, "PLANES"), pivote(p_dig, "ESIM"), pivote(b_dig, "PLANES")

    bot_m["join_num"] = bot_m["Id Contacto/Número"].astype(str).str.replace(r"\.0$", "", regex=True).str[-10:]
    bot_an, land_all["Fuente"] = D["bot_an"].copy(), limpio(land_all["Fuente"])
    bot_an["join_num"] = bot_an["NIM"].astype(str).str.replace(r"\.0$", "", regex=True).str[-10:]
    bot_an_u = bot_an.dropna(subset=["Anuncio_Titulo"])[["join_num", "Anuncio_Titulo", "Producto"]].drop_duplicates("join_num", keep="last")
    bot_m_merged = pd.merge(bot_m, bot_an_u, on="join_num", how="left")
    land_all["Plataforma"], land_all["Producto"] = land_all["Plataforma"].astype(str).str.upper().str.strip(), land_all["Producto"].astype(str).str.upper().str.strip()
    fuente_plat = (D["landing"].assign(F=limpio(D["landing"]["Fuente"]), P=D["landing"]["Plataforma"].astype(str).str.upper().str.strip()).dropna(subset=["F"]).groupby("F")["P"].agg(lambda s: s.value_counts().index[0]).to_dict())

    def con_origen(df):
        bm, ld = limpio(df["BOTMAKER"]), limpio(df["LANDING"])
        out = pd.DataFrame({"Producto": df["Producto_venta"], "Asesor": df["Asesor"], "Supervisor": df["Supervisor"], "Fecha": df["Fecha"]}, index=df.index)
        out["Origen"] = np.select([bm.notna() & (bm != "0"), ld.notna() & (ld != "0"), bm.eq("0")], ["Botmaker (anuncio)", "Landing", "Botmaker (orgánico)"], default="Sin identificar")
        out["Clave"] = bm.where(bm.notna() & (bm != "0"), ld.where(ld.notna() & (ld != "0"), np.nan))
        out["Plataforma"] = np.select([out["Origen"] == "Botmaker (anuncio)", out["Origen"] == "Landing", out["Origen"] == "Botmaker (orgánico)"], ["META", out["Clave"].map(fuente_plat).fillna("SIN IDENTIFICAR"), "ORGÁNICO"], default="SIN IDENTIFICAR")
        return out
    ventas_all = pd.concat([con_origen(p_dig), con_origen(l_dig), con_origen(b_dig)], ignore_index=True)

    def moda(k, p):
        t = pd.DataFrame({"k": k, "p": p}).dropna()
        return t[t["p"].isin(["PORTA", "BAF"])].groupby("k")["p"].agg(lambda x: x.value_counts().index[0]).to_dict() if not t.empty else {}
    clave_prod = {**moda(bot_m_merged["Anuncio_Titulo"], bot_m_merged["Producto"].astype(str).str.upper().str.strip()), **moda(land_all["Fuente"], land_all["Producto"])}

    return {"v_porta": v_porta, "v_ln": v_ln, "v_baf": v_baf, "v_totales": v_totales, "leads_tot": total_leads, "h_totales": h_totales, "sph_total": v_totales / h_totales if h_totales > 0 else 0, "sph_baf_porta": (v_porta + v_baf) / h_totales if h_totales > 0 else 0, "efectividad": v_totales / total_leads if total_leads > 0 else 0, "df_daily": df_daily, "df_asesores": df_as, "horas_m": horas_m, "p_dig": p_dig, "l_dig": l_dig, "b_dig": b_dig, "bot_m_merged": bot_m_merged, "land_all": land_all, "ventas_all": ventas_all, "clave_prod": clave_prod, "df_leads_d": df_leads_d, "df_leads_h": df_leads_h, "p_planes_pivot": p_planes_pivot, "esim_pivot": esim_pivot, "baf_pivot": baf_pivot}

# ==========================================================
# 4. RENDERIZADO DE PESTAÑAS
# ==========================================================

def render_historico():
    st.markdown("#### Histórico General")
    df_hist = pd.DataFrame({
        'Mes': ['Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre'],
        'Horas Totales': [2892, 2790, 3102, 3150, 3156, 3030, 3282, 2208, 120],
        'Total Ventas': [1880, 1692, 1376, 1428, 1334, 1401, 1203, 1065, 45],
        'SPH Total': [0.65, 0.61, 0.44, 0.45, 0.42, 0.46, 0.37, 0.48, 0.38],
        'SPH BAF+Porta': [np.nan, np.nan, np.nan, 0.34, 0.30, 0.34, 0.29, 0.38, 0.30],
        'Porta Total': [695, 562, 604, 518, 399, 542, 511, 366, 12],
        'LN Total': [580, 503, 339, 350, 395, 365, 248, 230, 9],
        'BAF Total': [605, 627, 433, 560, 540, 494, 444, 469, 24],
        'Leads Totales': [16716, 17406, 18494, 16834, 18110, 15567, 10866, 12412, 954],
        'Efectividad': [0.11, 0.10, 0.07, 0.08, 0.07, 0.09, 0.11, 0.09, 0.05],
        'Presupuesto ($)': [np.nan, np.nan, np.nan, 33831810.37, 34736333.59, 37593975.01, 37823181.54, 29823696.35, 2848547.26],
        'CPL Mensual ($)': [np.nan, np.nan, np.nan, 2009.73, 1918.07, 2414.98, 3480.87, 2402.81, 2985.90],
        'CPL Ciclo ($)': [np.nan, np.nan, np.nan, 1638.50, 2092.93, 2673.25, 2963.73, 2606.97, 1721.18]
    }).set_index('Mes')

    df_hist_manana = pd.DataFrame({
        'Mes': ['Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre'],
        'Horas Mañana': [1500, 1350, 1584, 1728, 1728, 1962, 2184, 1506, 78],
        'Porta': [327, 251, 259, 295, 203, 375, 329, 253, 7],
        'LN': [273, 124, 133, 180, 206, 213, 130, 130, 3],
        'BAF': [316, 224, 207, 274, 265, 336, 333, 336, 13],
        'Total Ventas': [916, 599, 599, 749, 674, 924, 792, 719, 23],
        'SPH Mañana': [0.61, 0.44, 0.38, 0.43, 0.39, 0.47, 0.36, 0.48, 0.29],
        'SPH BAF+Porta': [np.nan, np.nan, np.nan, 0.33, 0.27, 0.36, 0.30, 0.39, 0.26],
    }).set_index('Mes')

    df_hist_tarde = pd.DataFrame({
        'Mes': ['Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre'],
        'Horas Tarde': [1392, 1440, 1518, 1422, 1428, 1068, 1098, 702, 42],
        'Porta': [368, 311, 345, 223, 196, 167, 182, 113, 5],
        'LN': [307, 379, 206, 170, 189, 152, 118, 100, 6],
        'BAF': [289, 403, 226, 286, 275, 158, 111, 133, 11],
        'Total Ventas': [964, 1093, 777, 679, 660, 477, 411, 346, 22],
        'SPH Tarde': [0.69, 0.76, 0.51, 0.48, 0.46, 0.45, 0.37, 0.49, 0.52],
        'SPH BAF+Porta': [np.nan, np.nan, np.nan, 0.36, 0.33, 0.30, 0.27, 0.35, 0.38],
    }).set_index('Mes')

    mes_str_actual = mes_seleccionado.split()[0]
    def apply_historical_heatmap(data):
        styles = pd.DataFrame('', index=data.index, columns=data.columns)
        for i in range(1, len(data)):
            if str(data.index[i]) == mes_str_actual: continue 
            for col in data.columns:
                curr, prev = data.iloc[i][col], data.iloc[i-1][col]
                if pd.isna(curr) or pd.isna(prev) or prev == 0: continue
                if "Presupuesto" in col or "CPL" in col: 
                    styles.iloc[i, data.columns.get_loc(col)] = 'color: #C0392B; font-weight: bold; background-color: #FADBD8;' if float(curr) > float(prev) else 'color: #1E8449; font-weight: bold; background-color: #D4EFDF;'
                else: 
                    styles.iloc[i, data.columns.get_loc(col)] = 'color: #1E8449; font-weight: bold; background-color: #D4EFDF;' if float(curr) > float(prev) else 'color: #C0392B; font-weight: bold; background-color: #FADBD8;'
        return styles

    st.dataframe(df_hist.style.format({
        'Total Ventas': '{:,.0f}', 'Horas Totales': '{:,.0f}', 'Porta Total': '{:,.0f}', 'LN Total': '{:,.0f}', 'BAF Total': '{:,.0f}', 'Leads Totales': '{:,.0f}',
        'SPH Total': '{:.2f}', 'SPH BAF+Porta': '{:.2f}', 'Efectividad': '{:.0%}',
        'Presupuesto ($)': '${:,.2f}', 'CPL Mensual ($)': '${:,.2f}', 'CPL Ciclo ($)': '${:,.2f}'
    }, na_rep="-").apply(apply_historical_heatmap, axis=None), use_container_width=True, height=350)
    
    col_m, col_t = st.columns(2)
    with col_m:
        st.markdown("#### Histórico Turno Mañana")
        st.dataframe(df_hist_manana.style.format({'Horas Mañana': '{:,.0f}', 'Porta': '{:,.0f}', 'LN': '{:,.0f}', 'BAF': '{:,.0f}', 'Total Ventas': '{:,.0f}', 'SPH Mañana': '{:.2f}', 'SPH BAF+Porta': '{:.2f}'}, na_rep="-").apply(apply_historical_heatmap, axis=None), use_container_width=True, height=350)
    with col_t:
        st.markdown("#### Histórico Turno Tarde")
        st.dataframe(df_hist_tarde.style.format({'Horas Tarde': '{:,.0f}', 'Porta': '{:,.0f}', 'LN': '{:,.0f}', 'BAF': '{:,.0f}', 'Total Ventas': '{:,.0f}', 'SPH Tarde': '{:.2f}', 'SPH BAF+Porta': '{:.2f}'}, na_rep="-").apply(apply_historical_heatmap, axis=None), use_container_width=True, height=350)

def render_productividad(m):
    cs = st.columns(9)
    for c, (lab, val) in zip(cs, [("Ventas Brutas", f"{m['v_totales']:,}"), ("Portabilidad", m["v_porta"]), ("Línea Nueva", m["v_ln"]), ("BAF", m["v_baf"]), ("Leads Totales", f"{m['leads_tot']:,}"), ("Horas Totales", f"{m['h_totales']:,.0f}"), ("SPH Total", f"{m['sph_total']:.2f}"), ("SPH BAF+Porta", f"{m['sph_baf_porta']:.2f}"), ("Efectividad", f"{m['efectividad']:.2%}")]): c.markdown(card(lab, val), unsafe_allow_html=True)
    st.markdown("---")

    if m["df_daily"].empty: st.warning(f"No hay registros de productividad para {mes_seleccionado}.")
    else:
        flags = m["df_daily"]["_No_Laborable"].values
        df_clean = m["df_daily"].drop(columns=["_No_Laborable"]).set_index("Día")
        
        def style_diario(row):
            idx = df_clean.index.get_loc(row.name)
            out = []
            for col in row.index:
                s = f"background-color: {BRAND['fin_semana']}; color: #555555;" if flags[idx] else (f"background-color: {BRAND['manana']};" if col.endswith(" M") else (f"background-color: {BRAND['tarde']};" if col.endswith(" T") else (f"background-color: {BRAND['total']};" if col in ["Porta Tot", "LN Tot", "BAF Tot", "Total Ventas", "Horas Totales"] else "")))
                if "SPH" in col: s += style_sph(row[col])
                out.append(s)
            return out
        st.dataframe(df_clean.style.apply(style_diario, axis=1), use_container_width=True, height=600)

def render_asesores(m):
    modo = st.radio("Seleccionar Vista", ["Resumen Mensual", "Evolución Diaria / Por Asesor"], horizontal=True)
    df_as = m["df_asesores"].copy()
    if df_as.empty: st.warning(f"No hay asesores con actividad para {mes_seleccionado}.")
    elif modo == "Resumen Mensual":
        filtro_sup = st.selectbox("Filtrar por Turno / Supervisor", ["Todos", "RULFO, YAMIL (Mañana)", "PUENTES, SOLEDAD (Tarde)"])
        df_f = df_as[df_as["Supervisor"] == "RULFO, YAMIL"] if "RULFO" in filtro_sup else (df_as[df_as["Supervisor"] == "PUENTES, SOLEDAD"] if "PUENTES" in filtro_sup else df_as.copy())
        th, tp, tb, tl, tv = df_f["Horas"].sum(), df_f["PORTA"].sum(), df_f["BAF"].sum(), df_f["LN"].sum(), df_f["Total Ventas"].sum()
        show = df_f.copy()
        show["SPH Total"], show["SPH (baf/porta)"] = show["SPH Total"].map(lambda x: f"{x:.2f}"), show["SPH (baf/porta)"].map(lambda x: f"{x:.2f}")
        show = pd.concat([show, pd.DataFrame([{"Supervisor": "", "Asesor": "TOTAL", "Horas": th, "PORTA": tp, "LN": tl, "BAF": tb, "Total Ventas": tv, "SPH Total": f"{tv / th:.2f}" if th else "0.00", "SPH (baf/porta)": f"{(tp + tb) / th:.2f}" if th else "0.00"}])], ignore_index=True)
        show = show.set_index("Asesor")

        def style_as(row):
            bg = {"PUENTES, SOLEDAD": BRAND["tarde"], "RULFO, YAMIL": BRAND["manana"]}.get(row["Supervisor"], BRAND["total"])
            return [f"background-color: {bg}; font-weight: bold;" if col == "Supervisor" or row.name == "TOTAL" else (style_sph(row[col]) if col in ("SPH (baf/porta)", "SPH Total") else "") for col in row.index]
        st.dataframe(show.style.apply(style_as, axis=1), use_container_width=True, height=650)
    else:
        asesor_sel = st.selectbox("Filtrar por Asesor", sorted(df_as["Asesor"].unique()))
        h_a = m["horas_m"][m["horas_m"]["Asesor"] == asesor_sel]
        ventas_a = {k: m[k][m[k]["Asesor"] == asesor_sel] for k in ("p_dig", "l_dig", "b_dig")}
        rows = []
        for d in pd.date_range(pd.Timestamp(year=anio_sel, month=mes_sel, day=1), pd.Timestamp(year=anio_sel, month=mes_sel, day=1) + pd.offsets.MonthEnd(1)):
            dd = d.date()
            h_d = h_a[h_a["Fecha"].dt.date == dd]["Horas"].sum()
            po, ln_, ba = [int((ventas_a[k]["Fecha"].dt.date == dd).sum()) for k in ("p_dig", "l_dig", "b_dig")]
            if h_d == 0 and po + ln_ + ba == 0: continue
            rows.append({"Día": d.strftime("%d-%b"), "Horas": int(h_d), "PORTA": po, "BAF": ba, "LN": ln_, "Total Ventas": po + ln_ + ba, "SPH Total": f"{(po + ln_ + ba) / h_d:.2f}" if h_d > 0 else "-", "SPH (baf/porta)": f"{(po + ba) / h_d:.2f}" if h_d > 0 else "-"})
        df_ev = pd.DataFrame(rows)
        if not df_ev.empty:
            th, tp, tb, tl, tv = df_ev["Horas"].sum(), df_ev["PORTA"].sum(), df_ev["BAF"].sum(), df_ev["LN"].sum(), df_ev["Total Ventas"].sum()
            df_ev = pd.concat([df_ev, pd.DataFrame([{"Día": "TOTAL", "Horas": int(th), "PORTA": tp, "BAF": tb, "LN": tl, "Total Ventas": tv, "SPH Total": f"{tv / th:.2f}" if th else "0.00", "SPH (baf/porta)": f"{(tp + tb) / th:.2f}" if th else "0.00"}])], ignore_index=True)
        
        df_ev = df_ev.set_index("Día")
        def style_ev(row):
            try: en_cero = float(row["SPH (baf/porta)"]) == 0.0
            except ValueError: en_cero = False
            return [f"background-color: {BRAND['total']}; font-weight: bold;" + (style_sph(row[col]) if col in ("SPH (baf/porta)", "SPH Total") else "") if row.name == "TOTAL" else (f"background-color: {BRAND['bad_bg']}; color: {BRAND['bad_tx']}; font-weight: bold;" if en_cero else (style_sph(row[col]) if col in ("SPH (baf/porta)", "SPH Total") else "")) for col in row.index]
        st.dataframe(df_ev.style.apply(style_ev, axis=1), use_container_width=True, height=650)

def render_datos(m):
    st.markdown("#### Ingresos de Leads por Día (Landing vs Botmaker)")
    if "df_leads_d" not in m or m["df_leads_d"].empty:
        st.info("No hay datos diarios")
    else:
        st.dataframe(m["df_leads_d"].set_index("Día"), use_container_width=True, height=350)
    
    st.markdown("---")
    st.markdown("#### Ingresos de Leads por Hora")
    if "df_leads_h" in m and not m["df_leads_h"].empty:
        df_h = m["df_leads_h"].copy()
        df_h["Turno"] = df_h["Hora"].apply(lambda h: "Turno Mañana (09-14)" if 9<=h<=14 else ("Turno Tarde (15-20)" if 15<=h<=20 else "Turno Noche (21-08)"))
        filtro_turno = st.selectbox("Filtrar Análisis Horario", ["Todos", "Turno Mañana (09-14)", "Turno Tarde (15-20)", "Turno Noche (21-08)"])
        df_h_f = df_h[df_h["Turno"] == filtro_turno] if filtro_turno != "Todos" else df_h
        
        dias_activos = m["df_leads_d"][m["df_leads_d"]["Total"] > 0]["Día"].nunique() if not m["df_leads_d"].empty else 1
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Promedio Diario Leads (Selección)", f"{df_h_f['Total'].sum() / dias_activos:,.1f}")
        c2.metric("Promedio Botmaker (Diario)", f"{df_h_f['Botmaker'].sum() / dias_activos:,.1f}")
        c3.metric("Promedio Landing (Diario)", f"{df_h_f['Landing'].sum() / dias_activos:,.1f}")
        
        def style_hours(row):
            h = row.name
            if 9<=h<=14: bg = f"background-color: {BRAND['manana']};"
            elif 15<=h<=20: bg = f"background-color: {BRAND['tarde']};"
            else: bg = f"background-color: {BRAND['fin_semana']};" 
            return [bg] * len(row)
        st.dataframe(df_h_f.set_index("Hora").style.apply(style_hours, axis=1), use_container_width=True, height=350)

    st.markdown("---")
    st.markdown("#### Detalle de Planes Vendidos")
    def planes_filtrados(pivot_df, titulo, key):
        if pivot_df.empty: st.info(f"No hay datos para {titulo}"); return
        todas = list(pivot_df.columns)
        sel = st.multiselect(titulo, todas, default=todas, key=key)
        if not sel: return
        
        tot_dia = pivot_df.sum(axis=1).replace(0, 1)
        tot_mes = pivot_df.sum().sum() or 1
        show = pd.DataFrame(index=pivot_df.index)
        tot_sel_dia = pd.Series(0, index=pivot_df.index)
        fila_total, total_sel_mes = {}, 0
        
        for c in sel:
            show[c] = pivot_df[c]
            show[f"% {c}"] = (pivot_df[c] / tot_dia * 100).map(lambda x: f"{x:.1f}%")
            tot_sel_dia += pivot_df[c]
            fila_total[c], fila_total[f"% {c}"] = pivot_df[c].sum(), f"{pivot_df[c].sum() / tot_mes * 100:.1f}%"
            total_sel_mes += pivot_df[c].sum()
            
        show["Total Seleccionado"], show["% Selección"] = tot_sel_dia, (tot_sel_dia / tot_dia * 100).map(lambda x: f"{x:.1f}%")
        fila_total["Total Seleccionado"], fila_total["% Selección"] = total_sel_mes, f"{total_sel_mes / tot_mes * 100:.1f}%"
        
        show.index = show.index.astype(str)
        final = pd.concat([show, pd.DataFrame(fila_total, index=["TOTAL"])])
        st.dataframe(final.style.apply(lambda r: [f"background-color: {BRAND['total']}; font-weight: bold;"] * len(r) if r.name == "TOTAL" else [""] * len(r), axis=1), use_container_width=True, height=450)

    s1, s2, s3 = st.tabs(["Porta Planes", "Porta eSIM", "Planes BAF"])
    with s1: planes_filtrados(m["p_planes_pivot"], "Filtrar Planes de Portabilidad", f"f_porta_{mes_sel}")
    with s2: planes_filtrados(m["esim_pivot"], "Filtrar Tipo eSIM", f"f_esim_{mes_sel}")
    with s3: planes_filtrados(m["baf_pivot"], "Filtrar Planes BAF", f"f_baf_{mes_sel}")

FORMATO_EF = {"Leads": "{:,.0f}", "Efectividad": "{:.2%}", "Porta": "{:,.0f}", "% Porta": "{:.2%}", "LN": "{:,.0f}", "% LN": "{:.2%}", "BAF": "{:,.0f}", "% BAF": "{:.2%}", "Total Ventas": "{:,.0f}"}
ORDEN_EF = ["Leads", "Porta", "% Porta", "LN", "% LN", "BAF", "% BAF", "Total Ventas", "Efectividad"]

def completar_ef(df):
    for c in ("Leads", "Porta", "LN", "BAF"):
        if c not in df: df[c] = 0
    df = df.fillna(0)
    df["Total Ventas"] = df[["Porta", "LN", "BAF"]].sum(axis=1)
    df["Efectividad"] = (df["Total Ventas"] / df["Leads"].replace(0, np.nan)).fillna(0)
    for c in ("Porta", "LN", "BAF"): df[f"% {c}"] = (df[c] / df["Total Ventas"].replace(0, np.nan)).fillna(0)
    return df[ORDEN_EF]

def conteo_ventas(v): return {"Porta": int((v["Producto"] == "PORTA").sum()), "LN": int((v["Producto"] == "LN").sum()), "BAF": int((v["Producto"] == "BAF").sum())}
def cuadro_resumen(filas, v_all):
    df = pd.DataFrame({nombre: {"Leads": leads, **conteo_ventas(v_all[mask(v_all)])} for nombre, leads, mask in filas}).T
    df.loc["TOTAL"] = df.sum()
    return completar_ef(df)
def cuadro_por_clave(leads, ventas):
    v = pd.DataFrame(columns=["Porta", "LN", "BAF"]) if ventas.empty else ventas.groupby(["Clave", "Producto"]).size().unstack(fill_value=0).reindex(columns=["PORTA", "LN", "BAF"], fill_value=0)
    v.columns = ["Porta", "LN", "BAF"]
    df = completar_ef(pd.concat([leads.rename("Leads"), v], axis=1))
    return df[(df["Leads"] > 0) | (df["Total Ventas"] > 0)].sort_values("Efectividad", ascending=False)
def estilo_ef(df): return df.style.format(FORMATO_EF, na_rep="-").apply(lambda r: [f"background-color: {BRAND['total']}; font-weight: bold;"] * len(r) if r.name == "TOTAL" else [""] * len(r), axis=1)

def render_efectividad(m):
    bm, ld, v_all, clave_prod = m["bot_m_merged"], m["land_all"], m["ventas_all"], m["clave_prod"]
    bm_anuncio, n_org = bm[bm["Anuncio_Titulo"].notna()], int(bm["Anuncio_Titulo"].isna().sum())
    ld_meta, ld_goog, ld_otro = ld[ld["Plataforma"] == "META"], ld[ld["Plataforma"] == "GOOGLE"], ld[~ld["Plataforma"].isin(["META", "GOOGLE"])]
    es, es_land = lambda origen: (lambda v: v["Origen"] == origen), lambda plat: (lambda v: (v["Origen"] == "Landing") & (v["Plataforma"] == plat))
    def mostrar(titulo, df): st.markdown(f"<div class='section-header'>{titulo}</div>", unsafe_allow_html=True); st.dataframe(estilo_ef(df), use_container_width=True)
    mostrar("Resumen General por Plataforma", cuadro_resumen([("Meta", len(bm_anuncio) + len(ld_meta), lambda v: es("Botmaker (anuncio)")(v) | es_land("META")(v)), ("Google", len(ld_goog), es_land("GOOGLE")), ("Landing sin plataforma", len(ld_otro), lambda v: (v["Origen"] == "Landing") & ~v["Plataforma"].isin(["META", "GOOGLE"])), ("Botmaker orgánico", n_org, es("Botmaker (orgánico)")), ("Sin lead", 0, es("Sin identificar"))], v_all))
    mostrar("Resumen General por Herramienta", cuadro_resumen([("Botmaker anuncio", len(bm_anuncio), es("Botmaker (anuncio)")), ("Botmaker orgánico", n_org, es("Botmaker (orgánico)")), ("Landing", len(ld), es("Landing")), ("Sin lead", 0, es("Sin identificar"))], v_all))
    leads_prod = lambda p: int((bm_anuncio["Producto"].astype(str).str.upper().str.strip() == p).sum() + (ld["Producto"] == p).sum())
    mostrar("Resumen por Producto de Campaña", cuadro_resumen([("Campaña Porta", leads_prod("PORTA"), lambda v: v["Clave"].map(clave_prod) == "PORTA"), ("Campaña BAF", leads_prod("BAF"), lambda v: v["Clave"].map(clave_prod) == "BAF"), ("Sin identificar", len(bm) + len(ld) - leads_prod("PORTA") - leads_prod("BAF"), lambda v: ~v["Clave"].map(clave_prod).isin(["PORTA", "BAF"]))], v_all))

    st.markdown("---")
    filtro = st.selectbox("Filtrar Producto Vendido", ["Todos", "PORTA", "LN", "BAF"], key="filtro_prod_ads")
    vf = v_all if filtro == "Todos" else v_all[v_all["Producto"] == filtro]
    def l_clave(bm_sub, ld_sub):
        lds = pd.concat([bm_sub.groupby("Anuncio_Titulo").size(), ld_sub.dropna(subset=["Fuente"]).groupby("Fuente").size()]).groupby(level=0).sum()
        if int(ld_sub["Fuente"].isna().sum()): lds.loc["(landing sin fuente)"] = int(ld_sub["Fuente"].isna().sum())
        return lds
    st.markdown("<div class='section-header'>Anuncios por Plataforma</div>", unsafe_allow_html=True)
    c_meta, c_goog = st.columns(2)
    with c_meta: st.markdown("#### Meta"); st.dataframe(cuadro_por_clave(l_clave(bm_anuncio, ld_meta), vf[es("Botmaker (anuncio)")(vf) | es_land("META")(vf)]).style.format(FORMATO_EF), use_container_width=True, height=350)
    with c_goog: st.markdown("#### Google"); st.dataframe(cuadro_por_clave(l_clave(bm_anuncio.iloc[0:0], ld_goog), vf[es_land("GOOGLE")(vf)]).style.format(FORMATO_EF), use_container_width=True, height=350)
    st.markdown("<div class='section-header'>Anuncios por Producto de Campaña</div>", unsafe_allow_html=True)
    c_por, c_baf = st.columns(2)
    for col, prod, nm in ((c_por, "PORTA", "Porta"), (c_baf, "BAF", "BAF")):
        with col:
            st.markdown(f"#### Campaña: {nm}")
            lds = l_clave(bm_anuncio[bm_anuncio["Producto"].astype(str).str.upper().str.strip() == prod], ld[ld["Producto"] == prod])
            st.dataframe(cuadro_por_clave(lds, vf[vf["Clave"].isin(lds.index)]).style.format(FORMATO_EF), use_container_width=True, height=350)

def periodos_mes(anio, mes):
    p = PERIODOS_FACTURACION.get((anio, mes))
    if p: return {k: (pd.Timestamp(a), pd.Timestamp(b)) for k, (a, b) in p.items()}
    ini = pd.Timestamp(year=anio, month=mes, day=1)
    return {"META": (ini, ini + pd.offsets.MonthEnd(1)), "GOOGLE": (ini, ini + pd.offsets.MonthEnd(1))}

def calcular_presupuesto(anio, mes):
    D, per = cargar_datos(), periodos_mes(anio, mes)
    dias = pd.date_range(min(v[0] for v in per.values()), max(v[1] for v in per.values()))

    def gasto_de(df, plat, col):
        d = df[(df["Pais"] == "AR") & (df["Fecha"] >= per[plat][0]) & (df["Fecha"] <= per[plat][1])]
        return pd.DataFrame({"Fecha": d["Fecha"].dt.normalize(), "Plataforma": plat, "Producto": d["Producto"].astype(str).str.upper().str.strip(), "Gasto": pd.to_numeric(d[col], errors="coerce").fillna(0)})
    gasto = pd.concat([gasto_de(D["meta"], "META", "Monto gastado (ARS)"), gasto_de(D["google"], "GOOGLE", "Costo num")], ignore_index=True)

    bm, an, ld = D["botmaker"][(D["botmaker"]["PAIS"] == "AR") & (D["botmaker"]["Fecha"] >= dias[0]) & (D["botmaker"]["Fecha"] < dias[-1] + pd.Timedelta(days=1))].copy(), D["bot_an"].copy(), D["landing"][(D["landing"]["Pais"] == "AR") & (D["landing"]["Fecha"] >= dias[0]) & (D["landing"]["Fecha"] < dias[-1] + pd.Timedelta(days=1))].copy()
    bm["join_num"] = bm["Id Contacto/Número"].astype(str).str.replace(r"\.0$", "", regex=True).str[-10:]
    an["join_num"] = an["NIM"].astype(str).str.replace(r"\.0$", "", regex=True).str[-10:]
    bm = bm.merge(an[limpio(an["Anuncio_Titulo"]).notna()].drop_duplicates("join_num", keep="last")[["join_num", "Plataforma", "Producto"]], on="join_num", how="left")
    
    def norm(s, validos): return s.astype(str).str.upper().str.strip().where(s.astype(str).str.upper().str.strip().isin(validos), "SIN ID")
    leads = pd.concat([pd.DataFrame({"Fecha": bm["Fecha"].dt.normalize(), "Herr": "Botmaker", "Plat": norm(bm["Plataforma"], ["META", "GOOGLE"]), "Prod": norm(bm["Producto"], ["PORTA", "BAF"])}), pd.DataFrame({"Fecha": ld["Fecha"].dt.normalize(), "Herr": "Landing", "Plat": norm(ld["Plataforma"], ["META", "GOOGLE"]), "Prod": norm(ld["Producto"], ["PORTA", "BAF"])})], ignore_index=True)
    
    b, l = leads["Herr"] == "Botmaker", leads["Herr"] == "Landing"
    cnt = pd.DataFrame({"total": 1, "sin_id": b & (leads["Plat"] == "SIN ID"), "meta_bot": b & (leads["Plat"] == "META"), "meta_land": l & (leads["Plat"] == "META"), "goog_bot": b & (leads["Plat"] == "GOOGLE"), "goog_land": l & (leads["Plat"] == "GOOGLE"), "baf_bot": b & (leads["Prod"] == "BAF"), "baf_land": l & (leads["Prod"] == "BAF"), "porta_bot": b & (leads["Prod"] == "PORTA"), "porta_land": l & (leads["Prod"] == "PORTA")}).astype(int).groupby(leads["Fecha"]).sum().reindex(dias, fill_value=0)
    g_pl, g_pr = gasto.pivot_table(index="Fecha", columns="Plataforma", values="Gasto", aggfunc="sum").reindex(dias).fillna(0), gasto.pivot_table(index="Fecha", columns="Producto", values="Gasto", aggfunc="sum").reindex(dias).fillna(0)
    c = lambda df, k: df[k] if k in df else pd.Series(0.0, index=dias)

    M = pd.DataFrame(index=dias)
    M["Total ($)"], M["Meta ($)"], M["Google ($)"], M["Baf ($)"], M["Porta ($)"] = c(g_pl, "META") + c(g_pl, "GOOGLE"), c(g_pl, "META"), c(g_pl, "GOOGLE"), c(g_pr, "BAF"), c(g_pr, "PORTA")
    M["Otro ($)"] = M["Total ($)"] - M["Baf ($)"] - M["Porta ($)"]
    M["Total Leads"], M["Botmaker no se identifica"], M["Landing Meta"], M["Botmaker Meta"], M["Landing Google"], M["Botmaker Google"], M["Landing BAF"], M["Botmaker BAF"], M["Landing Porta"], M["Botmaker Porta"] = cnt["total"], cnt["sin_id"], cnt["meta_land"], cnt["meta_bot"], cnt["goog_land"], cnt["goog_bot"], cnt["baf_land"], cnt["baf_bot"], cnt["porta_land"], cnt["porta_bot"]
    M["Meta"], M["Google"] = M["Landing Meta"] + M["Botmaker Meta"] + (M["Botmaker no se identifica"] if SIN_ID_SUMA_A_PLATAFORMA == "META" else 0), M["Landing Google"] + M["Botmaker Google"] + (M["Botmaker no se identifica"] if SIN_ID_SUMA_A_PLATAFORMA == "GOOGLE" else 0)
    M["BAF"], M["Porta"] = M["Landing BAF"] + M["Botmaker BAF"] + (M["Botmaker no se identifica"] if SIN_ID_SUMA_A_PRODUCTO == "BAF" else 0), M["Landing Porta"] + M["Botmaker Porta"] + (M["Botmaker no se identifica"] if SIN_ID_SUMA_A_PRODUCTO == "PORTA" else 0)
    
    M = M[(M["Total ($)"] > 0) | (M["Total Leads"] > 0)].copy()
    return {"M": M, "periodos": per, "en_meta": (M.index >= per["META"][0]) & (M.index <= per["META"][1]), "en_goog": (M.index >= per["GOOGLE"][0]) & (M.index <= per["GOOGLE"][1])}

def render_presupuesto(m):
    P = calcular_presupuesto(anio_sel, mes_sel)
    M, per, f = P["M"], P["periodos"], lambda t: t.strftime("%d/%m")
    st.markdown("<div class='section-header'>Presupuesto Diario y Períodos de Facturación</div>", unsafe_allow_html=True)
    st.info(f"Período de facturación {mes_seleccionado}: Meta ({f(per['META'][0])} al {f(per['META'][1])}) | Google ({f(per['GOOGLE'][0])} al {f(per['GOOGLE'][1])})")

    T, cpl = M.sum(), lambda g_, l_: g_ / l_ if l_ else 0
    T_meta_l, T_goog_l = M.loc[P["en_meta"], "Meta"].sum(), M.loc[P["en_goog"], "Google"].sum()

    def armar(cols_fijas): return pd.DataFrame({"Fecha": M.index.strftime("%d/%m/%Y"), **{c: M[c].values for c in cols_fijas}})
    def con_total(df, total_dict): return pd.concat([df, pd.DataFrame([total_dict])], ignore_index=True).set_index("Fecha")
    def estilo(df): return df.style.apply(lambda r: [f"background-color: {BRAND['total']}; font-weight: bold;"] * len(r) if r.name == "TOTAL" else [""] * len(r), axis=1)

    DINERO, ENTERO, PCT, div = "${:,.2f}", "{:,.0f}", "{:.2%}", lambda a, b: a / b if b else 0

    d1 = armar(["Total ($)", "Meta ($)", "Google ($)", "Baf ($)", "Porta ($)", "Otro ($)"])
    for c in ("Meta ($)", "Google ($)", "Baf ($)", "Porta ($)", "Otro ($)"): d1[f"% {c[:-4]} ($)"] = (M[c] / M["Total ($)"].replace(0, np.nan)).fillna(0).values
    t1 = {"Fecha": "TOTAL", **{c: T[c] for c in ("Total ($)", "Meta ($)", "Google ($)", "Baf ($)", "Porta ($)", "Otro ($)")}, **{f"% {c} ($)": div(T[f"{c} ($)"], T["Total ($)"]) for c in ("Meta", "Google", "Baf", "Porta", "Otro")}}
    d1 = con_total(d1, t1)[["Total ($)", "Meta ($)", "% Meta ($)", "Google ($)", "% Google ($)", "Baf ($)", "% Baf ($)", "Porta ($)", "% Porta ($)", "Otro ($)", "% Otro ($)"]]
    if T["Otro ($)"] == 0: d1 = d1.drop(columns=["Otro ($)", "% Otro ($)"])
    st.markdown("<div class='sub-header'>1. PRESUPUESTO PLATAFORMA Y PRODUCTO</div>", unsafe_allow_html=True)
    st.dataframe(estilo(d1).format({**{c: DINERO for c in d1.columns if c.endswith("($)") and not c.startswith("%")}, **{c: PCT for c in d1.columns if c.startswith("%")}}), use_container_width=True, height=450)

    d2 = armar(["Total Leads", "Botmaker no se identifica", "Meta", "Landing Meta", "Botmaker Meta", "Google", "Landing Google"])
    d2["% Meta (Leads)"], d2["% Google (Leads)"] = (M["Meta"] / M["Total Leads"].replace(0, np.nan)).fillna(0).values, (M["Google"] / M["Total Leads"].replace(0, np.nan)).fillna(0).values
    d2["cpl meta"], d2["cpl google"] = (M["Meta ($)"] / M["Meta"].replace(0, np.nan)).fillna(0).values, (M["Google ($)"] / M["Google"].replace(0, np.nan)).fillna(0).values
    t2 = {"Fecha": "TOTAL", **{c: T[c] for c in ("Total Leads", "Botmaker no se identifica", "Meta", "Landing Meta", "Botmaker Meta", "Google", "Landing Google")}, "% Meta (Leads)": div(T["Meta"], T["Total Leads"]), "% Google (Leads)": div(T["Google"], T["Total Leads"]), "cpl meta": cpl(T["Meta ($)"], T_meta_l), "cpl google": cpl(T["Google ($)"], T_goog_l)}
    d2 = con_total(d2, t2)[["Total Leads", "Botmaker no se identifica", "Meta", "Landing Meta", "Botmaker Meta", "% Meta (Leads)", "Google", "Landing Google", "% Google (Leads)", "cpl meta", "cpl google"]]
    st.markdown("<div class='sub-header'>2. REALES Y CPL (Por Plataforma)</div>", unsafe_allow_html=True)
    st.dataframe(estilo(d2).format({
        "Total Leads": ENTERO, "Botmaker no se identifica": ENTERO, "Meta": ENTERO, "Landing Meta": ENTERO, "Botmaker Meta": ENTERO, "Google": ENTERO, "Landing Google": ENTERO,
        "cpl meta": DINERO, "cpl google": DINERO, "% Meta (Leads)": PCT, "% Google (Leads)": PCT
    }), use_container_width=True, height=450)

    d3 = armar(["Total Leads", "Botmaker no se identifica", "BAF", "Landing BAF", "Botmaker BAF", "Porta", "Landing Porta", "Botmaker Porta"])
    d3["% BAF"], d3["% Porta"] = (M["BAF"] / M["Total Leads"].replace(0, np.nan)).fillna(0).values, (M["Porta"] / M["Total Leads"].replace(0, np.nan)).fillna(0).values
    d3["cpl baf"], d3["cpl porta"] = (M["Baf ($)"] / M["BAF"].replace(0, np.nan)).fillna(0).values, (M["Porta ($)"] / M["Porta"].replace(0, np.nan)).fillna(0).values
    t3 = {"Fecha": "TOTAL", **{c: T[c] for c in ("Total Leads", "Botmaker no se identifica", "BAF", "Landing BAF", "Botmaker BAF", "Porta", "Landing Porta", "Botmaker Porta")}, "% BAF": div(T["BAF"], T["Total Leads"]), "% Porta": div(T["Porta"], T["Total Leads"]), "cpl baf": cpl(T["Baf ($)"], T["BAF"]), "cpl porta": cpl(T["Porta ($)"], T["Porta"])}
    d3 = con_total(d3, t3)[["Total Leads", "Botmaker no se identifica", "BAF", "Landing BAF", "Botmaker BAF", "% BAF", "Porta", "Landing Porta", "Botmaker Porta", "% Porta", "cpl baf", "cpl porta"]]
    st.markdown("<div class='sub-header'>3. REALES Y CPL (Por Producto)</div>", unsafe_allow_html=True)
    st.dataframe(estilo(d3).format({
        "Total Leads": ENTERO, "Botmaker no se identifica": ENTERO, "BAF": ENTERO, "Landing BAF": ENTERO, "Botmaker BAF": ENTERO, "Porta": ENTERO, "Landing Porta": ENTERO, "Botmaker Porta": ENTERO,
        "cpl baf": DINERO, "cpl porta": DINERO, "% BAF": PCT, "% Porta": PCT
    }), use_container_width=True, height=450)

# ==========================================================
# 5. ARMADO DE LA PÁGINA
# ==========================================================
try:
    m = calcular_metricas(anio_sel, mes_sel)
except Exception as e:
    st.error(f"Ocurrió un error procesando los datos: {e}")
    st.stop()

tab_hist, tab_prod, tab_asesores, tab_datos, tab_efectividad, tab_presupuesto = st.tabs(["Resumen Histórico", "Productividad", "Asesores", "Datos", "Efectividad", "Presupuesto"])

with tab_hist: render_historico()
with tab_prod: render_productividad(m)
with tab_asesores: render_asesores(m)
with tab_datos: render_datos(m)
with tab_efectividad: render_efectividad(m)
with tab_presupuesto: render_presupuesto(m)

import os
import streamlit as st
from google import genai
import pandas as pd # Asegúrate de que pandas esté importado

st.subheader("Chat Asistente Geex")

# Tomar la clave directamente desde Easypanel
api_key = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=api_key)

if "messages" not in st.session_state:
    st.session_state.messages = []

for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

if prompt := st.chat_input("Consulta métricas sobre las ventas o agentes..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # Leer el Excel para darle contexto a Gemini
    try:
        df = pd.read_excel('DASHHHHH.xlsx')
        contexto = df.to_string(index=False)
    except Exception as e:
        contexto = "No se pudo cargar el Excel."

    prompt_con_contexto = f"Basado en los siguientes datos del dashboard de Geex:\n\n{contexto}\n\nResponde a esta pregunta: {prompt}"

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt_con_contexto,
    )

    with st.chat_message("assistant"):
        st.markdown(response.text)
    st.session_state.messages.append({"role": "assistant", "content": response.text})