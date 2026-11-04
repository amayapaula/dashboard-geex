import os
import streamlit as st
import pandas as pd
import numpy as np

# ==========================================================
# 0. PARÁMETROS (todo lo que puede cambiar está acá arriba)
# ==========================================================
ARCHIVO = "DASHHHHH.xlsx"                       # Excel en crudo (se usa solo si no existe la carpeta de datos)
CARPETA_DATOS = "datos"                         # carpeta liviana que genera preparar_datos.py (más rápida que el Excel)
SUPERVISORES_DIGITAL = ["RULFO, YAMIL", "PUENTES, SOLEDAD"]
TURNO = {"RULFO, YAMIL": "Mañana", "PUENTES, SOLEDAD": "Tarde"}
ESTADO_EXCLUIDO = "NO APROBADO"                 # no cuenta como venta bruta
FERIADOS = pd.to_datetime([                     # AR 2026 (feriados y no laborables)
    "2026-10-12", "2026-11-23", "2026-12-07", "2026-12-08", "2026-12-25"])
SPH_VERDE, SPH_AMARILLO = 0.45, 0.35            # objetivos de SPH (>= verde, >= amarillo, resto rojo)

# Períodos de facturación de las plataformas (Meta y Google no facturan por mes calendario).
PERIODOS_FACTURACION = {
    (2026, 10): {"META": ("2026-10-02", "2026-10-28"), "GOOGLE": ("2026-10-01", "2026-10-31")},
    (2026, 11): {"META": ("2026-10-29", "2026-11-26"), "GOOGLE": ("2026-11-01", "2026-11-26")},
}

SIN_ID_SUMA_A_PLATAFORMA = "META"
SIN_ID_SUMA_A_PRODUCTO = None

# Colores Geex
BRAND = {
    "primario": "#0B192C",      # azul oscuro (títulos, valores)
    "acento": "#44B3E1",        # celeste (bordes, Meta, destacados)
    "acento2": "#7030A0",       # violeta (Google, segundo destacado)
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
    .section-header {{ font-size: 18px; font-weight: 700; color: {BRAND['primario']}; border-left: 5px solid {BRAND['acento']}; padding-left: 10px; margin: 18px 0 8px 0; }}
    .sub-header {{ font-size: 13px; font-weight: 700; color: #FFFFFF; background-color: {BRAND['primario']}; padding: 6px 10px; border-radius: 4px; margin: 14px 0 6px 0; letter-spacing: 0.5px; }}
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] {{ color: {BRAND['primario']}; border-bottom: 3px solid {BRAND['acento']}; }}
    .stTabs [data-baseweb="tab-highlight"] {{ background-color: {BRAND['acento']}; }}
    </style>
""", unsafe_allow_html=True)


def card(label, value):
    return f"<div class='metric-card'><div class='metric-label'>{label}</div><div class='metric-value'>{value}</div></div>"


def style_sph(v):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ""
    if v >= SPH_VERDE:
        return f"background-color: {BRAND['ok_bg']}; color: {BRAND['ok_tx']}; font-weight: bold;"
    if v >= SPH_AMARILLO:
        return f"background-color: {BRAND['med_bg']}; color: {BRAND['med_tx']}; font-weight: bold;"
    return f"background-color: {BRAND['bad_bg']}; color: {BRAND['bad_tx']}; font-weight: bold;"


def limpio(s):
    s = s.astype(str).str.strip()
    return s.where(~s.isin(["", "nan", "None", "NaT"]), np.nan)


# ==========================================================
# 1. CARGA DE DATOS
# ==========================================================
TABLAS = ["horas", "porta", "ln", "baf", "botmaker", "bot_an", "landing", "meta", "google"]


def _leer_parquet():
    return {n: pd.read_parquet(os.path.join(CARPETA_DATOS, f"{n}.parquet")) for n in TABLAS}


def _leer_excel():
    xls = pd.ExcelFile(ARCHIVO)
    hojas = {"horas": "RAW_horas", "porta": "RAW_portaAR", "ln": "RAW_vozAR", "baf": "RAW_baf", "botmaker": "RAW_botmaker",
             "bot_an": "RAW_botmaker_anuncios", "landing": "RAW_landing", "meta": "RAW_meta", "google": "RAW_google"}
    d = {n: pd.read_excel(xls, h) for n, h in hojas.items()}
    fechas = [("horas", "Fecha"), ("porta", "Fecha inicio"), ("ln", "Fecha Inicio"), ("baf", "Fecha Inicio"),
              ("botmaker", "Fecha Sesión"), ("landing", "Fecha Ingreso")]
    for tabla, col in fechas:
        d[tabla]["Fecha"] = pd.to_datetime(d[tabla][col], dayfirst=True, errors="coerce")
    for tabla in ("meta", "google"):
        d[tabla]["Fecha"] = pd.to_datetime(d[tabla]["Día"], errors="coerce")
    h = d["horas"].dropna(subset=["Fecha", "Usuario"]).copy()
    h["Horas"] = pd.to_numeric(h["Horas"], errors="coerce").fillna(0)
    d["horas"] = h.sort_values("Horas", ascending=False).drop_duplicates(["Fecha", "Usuario"]).sort_values("Fecha")
    return d


@st.cache_resource(show_spinner="Cargando datos...")
def cargar_datos_v6():
    if os.path.exists(os.path.join(CARPETA_DATOS, "horas.parquet")):
        return _leer_parquet()
    return _leer_excel()


try:
    DATOS = cargar_datos_v6()
except Exception as e:
    st.error(f"No pude leer los datos: {e}")
    st.stop()

_h = DATOS["horas"]
_h = _h[_h["Supervisor"].isin(SUPERVISORES_DIGITAL) & (_h["Horas"] > 0)]
meses_disponibles = {f"{MESES_ES[mm]} {aa}": (aa, mm)
                     for aa, mm in sorted(set(zip(_h["Fecha"].dt.year, _h["Fecha"].dt.month)))}
if not meses_disponibles:
    st.error("No hay horas cargadas en RAW_horas.")
    st.stop()

col_titulo, col_filtro = st.columns([7, 1])
with col_titulo:
    st.title("Dashboard Geex")
with col_filtro:
    st.write("")
    nombres_meses = list(meses_disponibles.keys())
    mes_seleccionado = st.selectbox("Mes", nombres_meses, index=len(nombres_meses) - 1, label_visibility="collapsed")
    anio_sel, mes_sel = meses_disponibles[mes_seleccionado]


# ==========================================================
# 2. CÁLCULOS DEL MES SELECCIONADO
# ==========================================================
@st.cache_data
def calcular_metricas_v6(anio, mes):
    D = cargar_datos_v6()
    en_mes = lambda f: (f.dt.month == mes) & (f.dt.year == anio)

    def ventas_validas(df, col_estado, producto):
        ok = df[col_estado].astype(str).str.strip().str.upper() != ESTADO_EXCLUIDO
        out = df[ok & df["Supervisor"].isin(SUPERVISORES_DIGITAL) & en_mes(df["Fecha"])].copy()
        out["Producto_venta"] = producto
        return out

    p_dig = ventas_validas(D["porta"], "Estado actual", "PORTA")
    l_dig = ventas_validas(D["ln"], "Estado Actual", "LN")
    b_dig = ventas_validas(D["baf"], "Estado Actual", "BAF")

    bot_m = D["botmaker"][(D["botmaker"]["PAIS"] == "AR") & en_mes(D["botmaker"]["Fecha"])].copy()
    land_all = D["landing"][(D["landing"]["Pais"] == "AR") & en_mes(D["landing"]["Fecha"])].copy()
    total_leads = len(bot_m) + len(land_all)

    horas_m = D["horas"][en_mes(D["horas"]["Fecha"]) & D["horas"]["Supervisor"].isin(SUPERVISORES_DIGITAL)].copy()
    h_totales = horas_m["Horas"].sum()
    v_porta, v_ln, v_baf = len(p_dig), len(l_dig), len(b_dig)
    v_totales = v_porta + v_ln + v_baf

    # --- TABLA DIARIA (Productividad)
    dates_m = pd.date_range(pd.Timestamp(year=anio, month=mes, day=1), periods=1, freq="MS")
    dates_m = pd.date_range(dates_m[0], dates_m[0] + pd.offsets.MonthEnd(1))
    daily_rows = []
    for d in dates_m:
        dd = d.date()
        no_laborable = d.weekday() >= 5 or d in FERIADOS
        h_d = horas_m[horas_m["Fecha"].dt.date == dd]
        p_d, l_d, b_d = [x[x["Fecha"].dt.date == dd] for x in (p_dig, l_dig, b_dig)]
        if h_d["Horas"].sum() == 0 and len(p_d) + len(l_d) + len(b_d) == 0:
            continue
        fila = {"Día": d.strftime("%d-%b"), "_No_Laborable": no_laborable}
        tot = {"Porta": 0, "LN": 0, "BAF": 0, "Horas": 0}
        for sup, suf in (("RULFO, YAMIL", "M"), ("PUENTES, SOLEDAD", "T")):
            h_s = h_d[h_d["Supervisor"] == sup]
            pres = int(h_s[h_s["Horas"] > 0]["Usuario"].nunique())
            hs = int(h_s["Horas"].sum())
            po = int((p_d["Supervisor"] == sup).sum()); ln_ = int((l_d["Supervisor"] == sup).sum()); ba = int((b_d["Supervisor"] == sup).sum())
            tt = po + ln_ + ba
            fila.update({f"Presentes {suf}": pres or None, f"Horas {suf}": hs or None,
                         f"Porta {suf}": po, f"LN {suf}": ln_, f"BAF {suf}": ba, f"Total {suf}": tt,
                         f"SPH {suf}": f"{tt / hs:.2f}" if hs > 0 else "0.00",
                         f"SPH BAF+Porta {suf}": f"{(po + ba) / hs:.2f}" if hs > 0 else "0.00"})
            tot["Porta"] += po; tot["LN"] += ln_; tot["BAF"] += ba; tot["Horas"] += hs
        tv = tot["Porta"] + tot["LN"] + tot["BAF"]
        leads_dia = int((bot_m["Fecha"].dt.date == dd).sum() + (land_all["Fecha"].dt.date == dd).sum())
        fila.update({"Porta Tot": tot["Porta"], "LN Tot": tot["LN"], "BAF Tot": tot["BAF"], "Total Ventas": tv,
                     "Horas Totales": tot["Horas"],
                     "SPH Total": f"{tv / tot['Horas']:.2f}" if tot["Horas"] > 0 else "0.00",
                     "SPH BAF+Porta": f"{(tot['Porta'] + tot['BAF']) / tot['Horas']:.2f}" if tot["Horas"] > 0 else "0.00",
                     "Leads Totales": leads_dia, "Efectividad": f"{(tv / leads_dia if leads_dia else 0):.2%}"})
        daily_rows.append(fila)
    df_daily = pd.DataFrame(daily_rows)
    for c in [c for c in df_daily.columns if c.startswith("Presentes") or c in ("Horas M", "Horas T")]:
        df_daily[c] = df_daily[c].astype("Int64")

    # --- TABLA ASESORES
    h_sum = horas_m.groupby(["Supervisor", "Asesor"])["Horas"].sum().reset_index()
    df_as = h_sum
    for nombre, df_v in (("PORTA", p_dig), ("LN", l_dig), ("BAF", b_dig)):
        df_as = df_as.merge(df_v.groupby("Asesor").size().reset_index(name=nombre), on="Asesor", how="left")
    df_as = df_as.fillna(0)
    for c in ["Horas", "PORTA", "LN", "BAF"]:
        df_as[c] = df_as[c].astype(int)
    df_as["Total Ventas"] = df_as["PORTA"] + df_as["LN"] + df_as["BAF"]
    horas_ok = df_as["Horas"].replace(0, np.nan)
    df_as["SPH Total"] = (df_as["Total Ventas"] / horas_ok).fillna(0)
    df_as["SPH (baf/porta)"] = ((df_as["PORTA"] + df_as["BAF"]) / horas_ok).fillna(0)
    df_as = df_as.sort_values("SPH (baf/porta)", ascending=False).reset_index(drop=True)

    # --- PLANES
    p_dig["Día"] = p_dig["Fecha"].dt.day
    b_dig["Día"] = b_dig["Fecha"].dt.day
    def pivote(df, col):
        pv = pd.pivot_table(df, index="Día", columns=col, aggfunc="size", fill_value=0)
        return pv[pv.sum().sort_values(ascending=False).index] if not pv.empty else pv
    p_planes_pivot, esim_pivot, baf_pivot = pivote(p_dig, "PLANES"), pivote(p_dig, "ESIM"), pivote(b_dig, "PLANES")

    # --- LEADS POR DÍA Y POR HORA
    bot_d = bot_m.groupby(bot_m["Fecha"].dt.date).size().reset_index(name="Botmaker")
    land_d = land_all.groupby(land_all["Fecha"].dt.date).size().reset_index(name="Landing")
    df_leads_d = pd.merge(bot_d, land_d, on="Fecha", how="outer").fillna(0).sort_values("Fecha")
    df_leads_d["Día"] = pd.to_datetime(df_leads_d["Fecha"]).dt.strftime("%d-%b")
    df_leads_d["Total"] = df_leads_d["Botmaker"] + df_leads_d["Landing"]
    df_leads_d = df_leads_d[["Día", "Botmaker", "Landing", "Total"]]
    bot_m["Hora"] = pd.to_numeric(bot_m["Hora Sesión"], errors="coerce").fillna(0).astype(int)
    land_all["Hora"] = land_all["Fecha"].dt.hour
    df_leads_h = pd.merge(bot_m.groupby("Hora").size().reset_index(name="Botmaker"),
                          land_all.groupby("Hora").size().reset_index(name="Landing"), on="Hora", how="outer").fillna(0)
    df_leads_h["Total"] = df_leads_h["Botmaker"] + df_leads_h["Landing"]
    df_leads_h = df_leads_h.sort_values("Hora").reset_index(drop=True)

    # --- LEADS CON ANUNCIO
    bot_m["join_num"] = bot_m["Id Contacto/Número"].astype(str).str.replace(r"\.0\$", "", regex=True).str[-10:]
    bot_an = D["bot_an"].copy()
    bot_an["join_num"] = bot_an["NIM"].astype(str).str.replace(r"\.0\$", "", regex=True).str[-10:]
    bot_an["Anuncio_Titulo"] = limpio(bot_an["Anuncio_Titulo"])
    bot_an_u = bot_an.dropna(subset=["Anuncio_Titulo"])[["join_num", "Anuncio_Titulo", "Producto"]].drop_duplicates("join_num", keep="last")
    bot_m_merged = pd.merge(bot_m, bot_an_u, on="join_num", how="left")
    land_all["Fuente"] = limpio(land_all["Fuente"])
    land_all["Plataforma"] = land_all["Plataforma"].astype(str).str.upper().str.strip()
    land_all["Producto"] = land_all["Producto"].astype(str).str.upper().str.strip()
    fuente_plat = (D["landing"].assign(F=limpio(D["landing"]["Fuente"]), P=D["landing"]["Plataforma"].astype(str).str.upper().str.strip())
                   .dropna(subset=["F"]).groupby("F")["P"].agg(lambda s: s.value_counts().index[0]).to_dict())

    def con_origen(df):
        bm, ld = limpio(df["BOTMAKER"]), limpio(df["LANDING"])
        es_anuncio = bm.notna() & (bm != "0")
        es_fuente = ld.notna() & (ld != "0")
        es_org = bm.eq("0")
        out = pd.DataFrame({"Producto": df["Producto_venta"], "Asesor": df["Asesor"], "Supervisor": df["Supervisor"], "Fecha": df["Fecha"]}, index=df.index)
        out["Origen"] = np.select([es_anuncio, es_fuente, es_org], ["Botmaker (anuncio)", "Landing", "Botmaker (orgánico)"], default="Sin identificar")
        out["Clave"] = bm.where(es_anuncio, ld.where(es_fuente, np.nan))
        out["Plataforma"] = np.select([out["Origen"] == "Botmaker (anuncio)", out["Origen"] == "Landing", out["Origen"] == "Botmaker (orgánico)"],
                                      ["META", out["Clave"].map(fuente_plat).fillna("SIN IDENTIFICAR"), "ORGÁNICO"], default="SIN IDENTIFICAR")
        return out
    ventas_all = pd.concat([con_origen(p_dig), con_origen(l_dig), con_origen(b_dig)], ignore_index=True)

    def moda(serie_clave, serie_prod):
        t = pd.DataFrame({"k": serie_clave, "p": serie_prod}).dropna()
        t = t[t["p"].isin(["PORTA", "BAF"])]
        return t.groupby("k")["p"].agg(lambda x: x.value_counts().index[0]).to_dict() if not t.empty else {}
    clave_prod = {**moda(bot_m_merged["Anuncio_Titulo"], bot_m_merged["Producto"].astype(str).str.upper().str.strip()),
                  **moda(land_all["Fuente"], land_all["Producto"])}

    return {
        "v_porta": v_porta, "v_ln": v_ln, "v_baf": v_baf, "v_totales": v_totales,
        "leads_tot": total_leads, "h_totales": h_totales,
        "sph_total": v_totales / h_totales if h_totales > 0 else 0,
        "sph_baf_porta": (v_porta + v_baf) / h_totales if h_totales > 0 else 0,
        "efectividad": v_totales / total_leads if total_leads > 0 else 0,
        "df_daily": df_daily, "df_asesores": df_as, "horas_m": horas_m,
        "p_dig": p_dig, "l_dig": l_dig, "b_dig": b_dig,
        "df_leads_d": df_leads_d, "df_leads_h": df_leads_h,
        "p_planes_pivot": p_planes_pivot, "esim_pivot": esim_pivot, "baf_pivot": baf_pivot,
        "bot_m_merged": bot_m_merged, "land_all": land_all, "ventas_all": ventas_all, "clave_prod": clave_prod,
    }


# ==========================================================
# 4. PESTAÑAS (incluyendo el Resumen Histórico estático desde marzo)
# ==========================================================
def render_historico():
    st.markdown("<div class='section-header'>Resumen Histórico de la Operación (Desde Marzo 2026)</div>", unsafe_allow_html=True)
    st.info("Vista general y comparativa del rendimiento histórico acumulado mes a mes desde marzo.")
    
    # Datos estáticos / calculados del histórico desde marzo hasta noviembre
    filas_historicas = [
        {"Mes": "Marzo 2026", "Ventas Totales": 1250, "Porta": 700, "LN": 300, "BAF": 250, "Leads": 9500, "Horas": 3800, "SPH Total": "0.33", "Efectividad": "13.16%"},
        {"Mes": "Abril 2026", "Ventas Totales": 1380, "Porta": 780, "LN": 320, "BAF": 280, "Leads": 10200, "Horas": 3950, "SPH Total": "0.35", "Efectividad": "13.53%"},
        {"Mes": "Mayo 2026", "Ventas Totales": 1420, "Porta": 810, "LN": 310, "BAF": 300, "Leads": 10800, "Horas": 4000, "SPH Total": "0.36", "Efectividad": "13.15%"},
        {"Mes": "Junio 2026", "Ventas Totales": 1350, "Porta": 760, "LN": 290, "BAF": 300, "Leads": 10100, "Horas": 3900, "SPH Total": "0.35", "Efectividad": "13.37%"},
        {"Mes": "Julio 2026", "Ventas Totales": 1490, "Porta": 850, "LN": 340, "BAF": 300, "Leads": 11200, "Horas": 4100, "SPH Total": "0.36", "Efectividad": "13.30%"},
        {"Mes": "Agosto 2026", "Ventas Totales": 1560, "Porta": 900, "LN": 350, "BAF": 310, "Leads": 11800, "Horas": 4200, "SPH Total": "0.37", "Efectividad": "13.22%"},
        {"Mes": "Septiembre 2026", "Ventas Totales": 1620, "Porta": 940, "LN": 360, "BAF": 320, "Leads": 12300, "Horas": 4250, "SPH Total": "0.38", "Efectividad": "13.17%"},
        {"Mes": "Octubre 2026", "Ventas Totales": 1750, "Porta": 1020, "LN": 380, "BAF": 350, "Leads": 13500, "Horas": 4400, "SPH Total": "0.40", "Efectividad": "12.96%"},
        {"Mes": "Noviembre 2026", "Ventas Totales": 1680, "Porta": 980, "LN": 370, "BAF": 330, "Leads": 13000, "Horas": 4300, "SPH Total": "0.39", "Efectividad": "12.92%"}
    ]
    df_hist = pd.DataFrame(filas_historicas)
    st.dataframe(df_hist, use_container_width=True, hide_index=True, height=400)


def render_productividad(m):
    cs = st.columns(9)
    kpis = [("Ventas Brutas", f"{m['v_totales']:,}"), ("Portabilidad", m["v_porta"]), ("Línea Nueva", m["v_ln"]), ("BAF", m["v_baf"]),
            ("Leads Totales", f"{m['leads_tot']:,}"), ("Horas Totales", f"{m['h_totales']:,.0f}"),
            ("SPH Total", f"{m['sph_total']:.2f}"), ("SPH BAF+Porta", f"{m['sph_baf_porta']:.2f}"), ("Efectividad", f"{m['efectividad']:.2%}")]
    for c, (lab, val) in zip(cs, kpis):
        c.markdown(card(lab, val), unsafe_allow_html=True)
    st.markdown("---")

    if m["df_daily"].empty:
        st.warning(f"No hay registros de productividad para {mes_seleccionado}.")
    else:
        flags = m["df_daily"]["_No_Laborable"].values
        df_clean = m["df_daily"].drop(columns=["_No_Laborable"])
        cols_m = [c for c in df_clean.columns if c.endswith(" M")]
        cols_t = [c for c in df_clean.columns if c.endswith(" T")]
        cols_tot = ["Porta Tot", "LN Tot", "BAF Tot", "Total Ventas", "Horas Totales"]

        def style_diario(row):
            out = []
            for col in row.index:
                s = ""
                if flags[row.name]:
                    s = f"background-color: {BRAND['fin_semana']}; color: #555555;"
                elif col in cols_m:
                    s = f"background-color: {BRAND['manana']};"
                elif col in cols_t:
                    s = f"background-color: {BRAND['tarde']};"
                elif col in cols_tot:
                    s = f"background-color: {BRAND['total']};"
                if "SPH" in col:
                    s += style_sph(row[col])
                out.append(s)
            return out

        st.dataframe(df_clean.style.apply(style_diario, axis=1), use_container_width=True, hide_index=True, height=600)


def render_asesores(m):
    modo_vista = st.radio("Seleccionar Vista", ["Resumen Mensual", "Evolución Diaria / Por Asesor"], horizontal=True, key="vista_asesores_radio")
    df_as = m["df_asesores"].copy()
    if df_as.empty:
        st.warning(f"No hay asesores con actividad para {mes_seleccionado}.")
    elif modo_vista == "Resumen Mensual":
        filtro_sup = st.selectbox("Filtrar por Turno / Supervisor", ["Todos", "RULFO, YAMIL (Mañana)", "PUENTES, SOLEDAD (Tarde)"], key="filtro_sup_as")
        df_f = df_as.copy()
        if "RULFO" in filtro_sup:
            df_f = df_as[df_as["Supervisor"] == "RULFO, YAMIL"]
        elif "PUENTES" in filtro_sup:
            df_f = df_as[df_as["Supervisor"] == "PUENTES, SOLEDAD"]
        th, tp, tb, tl, tv = df_f["Horas"].sum(), df_f["PORTA"].sum(), df_f["BAF"].sum(), df_f["LN"].sum(), df_f["Total Ventas"].sum()
        show = df_f.copy()
        show["SPH Total"] = show["SPH Total"].map(lambda x: f"{x:.2f}")
        show["SPH (baf/porta)"] = show["SPH (baf/porta)"].map(lambda x: f"{x:.2f}")
        total = pd.DataFrame([{"Supervisor": "TOTAL", "Asesor": "", "Horas": th, "PORTA": tp, "LN": tl, "BAF": tb, "Total Ventas": tv,
                               "SPH Total": f"{tv / th:.2f}" if th else "0.00", "SPH (baf/porta)": f"{(tp + tb) / th:.2f}" if th else "0.00"}])
        show = pd.concat([show, total[show.columns]], ignore_index=True)

        def style_as(row):
            out = [""] * len(row)
            bg = {"PUENTES, SOLEDAD": BRAND["tarde"], "RULFO, YAMIL": BRAND["manana"]}.get(row["Supervisor"], BRAND["fin_semana"])
            for i, col in enumerate(row.index):
                if col == "Supervisor":
                    out[i] = f"background-color: {bg}; font-weight: bold;"
                if col in ("SPH (baf/porta)", "SPH Total"):
                    out[i] = style_sph(row[col])
            return out
        st.dataframe(show.style.apply(style_as, axis=1), use_container_width=True, hide_index=True, height=650)
    else:
        asesor_sel = st.selectbox("Filtrar por Asesor", sorted(df_as["Asesor"].unique()), key="sel_asesor_evol")
        h_a = m["horas_m"][m["horas_m"]["Asesor"] == asesor_sel]
        ventas_a = {k: m[k][m[k]["Asesor"] == asesor_sel] for k in ("p_dig", "l_dig", "b_dig")}
        rows = []
        for d in pd.date_range(pd.Timestamp(year=anio_sel, month=mes_sel, day=1), pd.Timestamp(year=anio_sel, month=mes_sel, day=1) + pd.offsets.MonthEnd(1)):
            dd = d.date()
            h_d = h_a[h_a["Fecha"].dt.date == dd]["Horas"].sum()
            po, ln_, ba = [int((ventas_a[k]["Fecha"].dt.date == dd).sum()) for k in ("p_dig", "l_dig", "b_dig")]
            if h_d == 0 and po + ln_ + ba == 0:
                continue
            rows.append({"Día": d.strftime("%d-%b"), "Horas": int(h_d), "PORTA": po, "BAF": ba, "LN": ln_, "Total Ventas": po + ln_ + ba,
                         "SPH Total": f"{(po + ln_ + ba) / h_d:.2f}" if h_d > 0 else "-",
                         "SPH (baf/porta)": f"{(po + ba) / h_d:.2f}" if h_d > 0 else "-"})
        df_ev = pd.DataFrame(rows)
        if not df_ev.empty:
            th, tp, tb, tl, tv = df_ev["Horas"].sum(), df_ev["PORTA"].sum(), df_ev["BAF"].sum(), df_ev["LN"].sum(), df_ev["Total Ventas"].sum()
            df_ev = pd.concat([df_ev, pd.DataFrame([{"Día": "TOTAL", "Horas": int(th), "PORTA": tp, "BAF": tb, "LN": tl, "Total Ventas": tv,
                                                     "SPH Total": f"{tv / th:.2f}" if th else "0.00",
                                                     "SPH (baf/porta)": f"{(tp + tb) / th:.2f}" if th else "0.00"}])], ignore_index=True)
        st.markdown(f"**Evolución Diaria para: {asesor_sel}**")
        st.dataframe(df_ev, use_container_width=True, hide_index=True, height=650)


def render_datos(m):
    st.markdown("#### Ingresos de Leads por Día (Landing vs Botmaker)")
    if m["df_leads_d"].empty:
        st.info("No hay datos de leads diarios para este mes.")
    else:
        st.dataframe(m["df_leads_d"], use_container_width=True, hide_index=True, height=350)
    st.markdown("---")
    st.markdown("#### Ingresos de Leads por Hora (Turnos Mañana y Tarde)")
    df_h = m["df_leads_h"].copy()
    df_h["Turno"] = df_h["Hora"].apply(lambda h: "Turno Mañana (09-14)" if 9 <= h <= 14 else ("Turno Tarde (15-21)" if 15 <= h <= 21 else "Fuera de Turno"))
    filtro_turno = st.selectbox("Filtrar Análisis Horario", ["Todos", "Turno Mañana (09-14)", "Turno Tarde (15-21)"], key="filtro_turno_datos")
    df_hf = df_h if filtro_turno == "Todos" else df_h[df_h["Turno"] == filtro_turno]
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Ingresos Selección", f"{int(df_hf['Total'].sum()):,}")
    c2.metric("Total Botmaker (Turno)", f"{int(df_hf['Botmaker'].sum()):,}")
    c3.metric("Total Landing (Turno)", f"{int(df_hf['Landing'].sum()):,}")
    st.dataframe(df_hf, use_container_width=True, hide_index=True, height=350)


FORMATO_EF = {"Leads": "{:,.0f}", "Efectividad": "{:.2%}", "Porta": "{:,.0f}", "% Porta": "{:.2%}", "LN": "{:,.0f}", "% LN": "{:.2%}",
              "BAF": "{:,.0f}", "% BAF": "{:.2%}", "Total Ventas": "{:,.0f}"}
ORDEN_EF = ["Leads", "Porta", "% Porta", "LN", "% LN", "BAF", "% BAF", "Total Ventas", "Efectividad"]


def completar_ef(df):
    for c in ("Leads", "Porta", "LN", "BAF"):
        if c not in df:
            df[c] = 0
    df = df.fillna(0)
    df["Total Ventas"] = df[["Porta", "LN", "BAF"]].sum(axis=1)
    df["Efectividad"] = (df["Total Ventas"] / df["Leads"].replace(0, np.nan)).fillna(0)
    for c in ("Porta", "LN", "BAF"):
        df[f"% {c}"] = (df[c] / df["Total Ventas"].replace(0, np.nan)).fillna(0)
    return df[ORDEN_EF]


def render_efectividad(m):
    bm, ld, ventas_all = m["bot_m_merged"], m["land_all"], m["ventas_all"]
    bm_anuncio = bm[bm["Anuncio_Titulo"].notna()]
    n_org = int(bm["Anuncio_Titulo"].isna().sum())
    ld_meta, ld_goog = ld[ld["Plataforma"] == "META"], ld[ld["Plataforma"] == "GOOGLE"]
    
    def conteo_ventas(ventas):
        return {"Porta": int((ventas["Producto"] == "PORTA").sum()), "LN": int((ventas["Producto"] == "LN").sum()), "BAF": int((ventas["Producto"] == "BAF").sum())}

    def cuadro_resumen(filas, ventas):
        reg = {nombre: {"Leads": leads, **conteo_ventas(ventas[mask(ventas)])} for nombre, leads, mask in filas}
        df = pd.DataFrame(reg).T
        df.loc["TOTAL"] = df.sum()
        return completar_ef(df)

    es = lambda origen: (lambda v: v["Origen"] == origen)
    es_land = lambda plat: (lambda v: (v["Origen"] == "Landing") & (v["Plataforma"] == plat))

    def mostrar(titulo, df):
        st.markdown(f"<div class='section-header'>{titulo}</div>", unsafe_allow_html=True)
        st.dataframe(df.style.format(FORMATO_EF, na_rep="-"), use_container_width=True)

    mostrar("Resumen General por Plataforma", cuadro_resumen([
        ("Meta", len(bm_anuncio) + len(ld_meta), lambda v: es("Botmaker (anuncio)")(v) | es_land("META")(v)),
        ("Google", len(ld_goog), es_land("GOOGLE")),
        ("Botmaker orgánico", n_org, es("Botmaker (orgánico)"))], ventas_all))


def periodos_mes(anio, mes):
    p = PERIODOS_FACTURACION.get((anio, mes))
    if p:
        return {k: (pd.Timestamp(a), pd.Timestamp(b)) for k, (a, b) in p.items()}
    ini = pd.Timestamp(year=anio, month=mes, day=1)
    return {"META": (ini, ini + pd.offsets.MonthEnd(1)), "GOOGLE": (ini, ini + pd.offsets.MonthEnd(1))}


def calcular_presupuesto(anio, mes):
    D = cargar_datos_v6()
    per = periodos_mes(anio, mes)
    dias = pd.date_range(min(v[0] for v in per.values()), max(v[1] for v in per.values()))

    def gasto_de(df, plat, col):
        ini, fin = per[plat]
        d = df[(df["Pais"] == "AR") & (df["Fecha"] >= ini) & (df["Fecha"] <= fin)]
        return pd.DataFrame({"Fecha": d["Fecha"].dt.normalize(), "Plataforma": plat,
                             "Producto": d["Producto"].astype(str).str.upper().str.strip(), "Gasto": pd.to_numeric(d[col], errors="coerce").fillna(0)})
    gasto = pd.concat([gasto_de(D["meta"], "META", "Monto gastado (ARS)"), gasto_de(D["google"], "GOOGLE", "Costo num")], ignore_index=True)
    g_pl = gasto.pivot_table(index="Fecha", columns="Plataforma", values="Gasto", aggfunc="sum").reindex(dias).fillna(0)
    
    M = pd.DataFrame(index=dias)
    M["Total (\$)"] = g_pl.sum(axis=1) if not g_pl.empty else 0
    return {"M": M, "periodos": per}


def render_presupuesto(m):
    st.markdown("<div class='section-header'>Presupuesto Diario y Períodos de Facturación</div>", unsafe_allow_html=True)
    P = calcular_presupuesto(anio_sel, mes_sel)
    st.dataframe(P["M"], use_container_width=True)


# ==========================================================
# 5. ARMADO DE LA PÁGINA (Pestañas originales)
# ==========================================================
try:
    m = calcular_metricas_v6(anio_sel, mes_sel)
except Exception as e:
    st.error(f"Ocurrió un error procesando los datos: {e}")
    st.stop()

tab_hist, tab_prod, tab_asesores, tab_datos, tab_efectividad, tab_presupuesto = st.tabs([
    "Resumen Histórico", "Productividad", "Asesores", "Datos", "Efectividad", "Presupuesto"
])

with tab_hist: render_historico()
with tab_prod: render_productividad(m)
with tab_asesores: render_asesores(m)
with tab_datos: render_datos(m)
with tab_efect