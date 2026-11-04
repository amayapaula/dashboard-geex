import hashlib
import os
import sys
import pandas as pd

ARCHIVO = sys.argv[1] if len(sys.argv) > 1 else "DASHHHHH.xlsx"
CARPETA = "datos"
SUPERVISORES_DIGITAL = ["RULFO, YAMIL", "PUENTES, SOLEDAD"]
SAL = os.environ.get("SAL_TELEFONOS", "geex")

TABLAS = {
    "horas":    ("RAW_horas", ["Fecha", "Usuario", "Horas", "Asesor", "Supervisor"], "Fecha", False),
    "porta":    ("RAW_portaAR", ["Fecha inicio", "Supervisor", "Asesor", "Estado actual", "BOTMAKER", "LANDING", "PLANES", "ESIM"], "Fecha inicio", True),
    "ln":       ("RAW_vozAR", ["Fecha Inicio", "Supervisor", "Asesor", "Estado Actual", "BOTMAKER", "LANDING"], "Fecha Inicio", True),
    "baf":      ("RAW_baf", ["Fecha Inicio", "Supervisor", "Asesor", "Estado Actual", "BOTMAKER", "LANDING", "PLANES"], "Fecha Inicio", True),
    "botmaker": ("RAW_botmaker", ["PAIS", "Fecha Sesión", "Hora Sesión", "Id Contacto/Número"], "Fecha Sesión", False),
    "bot_an":   ("RAW_botmaker_anuncios", ["NIM", "Anuncio_Titulo", "Producto", "Plataforma"], None, False),
    "landing":  ("RAW_landing", ["Pais", "Fecha Ingreso", "Fuente", "Plataforma", "Producto"], "Fecha Ingreso", False),
    "meta":     ("RAW_meta", ["Pais", "Día", "Producto", "Monto gastado (ARS)"], "Día", False),
    "google":   ("RAW_google", ["Pais", "Día", "Producto", "Costo num"], "Día", False),
}

def codigo_telefono(serie):
    num = serie.astype(str).str.replace(r"\.0$", "", regex=True).str.replace(r"\D", "", regex=True).str[-10:]
    return num.map(lambda n: hashlib.sha256((SAL + n).encode()).hexdigest()[:16] if n else "")

def main():
    os.makedirs(CARPETA, exist_ok=True)
    xls = pd.ExcelFile(ARCHIVO)
    total = 0
    for nombre, (hoja, columnas, col_fecha, solo_digital) in TABLAS.items():
        df = pd.read_excel(xls, hoja)
        df = df[columnas].copy()
        if solo_digital:
            df = df[df["Supervisor"].isin(SUPERVISORES_DIGITAL)]
        if col_fecha:
            df["Fecha"] = pd.to_datetime(df[col_fecha], dayfirst=True, errors="coerce")
            if col_fecha != "Fecha":
                df = df.drop(columns=[col_fecha])
        for c in ("Id Contacto/Número", "NIM"):
            if c in df.columns:
                df[c] = codigo_telefono(df[c])
        if nombre == "horas":
            df = df.dropna(subset=["Fecha", "Usuario"])
            df["Horas"] = pd.to_numeric(df["Horas"], errors="coerce").fillna(0)
            df = df.sort_values("Horas", ascending=False).drop_duplicates(["Fecha", "Usuario"]).sort_values("Fecha")
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].astype("string")
        ruta = os.path.join(CARPETA, f"{nombre}.parquet")
        df.to_parquet(ruta, index=False)
    print("¡Proceso terminado con éxito!")

if __name__ == "__main__":
    main()