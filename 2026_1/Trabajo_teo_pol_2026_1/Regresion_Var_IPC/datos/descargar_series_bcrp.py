"""
Descarga series mensuales del BCRP via scraping HTML
Series: Produccion Nacional, Credito Privado, Credito Publico,
        Tasa Interbancaria, M1
Periodo: Enero 2004 - Diciembre 2025
Fuente: https://estadisticas.bcrp.gob.pe/estadisticas/series/
"""

import requests
import pandas as pd
import numpy as np
import os
import re
import time
from bs4 import BeautifulSoup

BASE_HTML = (
    "https://estadisticas.bcrp.gob.pe/estadisticas/series/"
    "mensuales/resultados/{code}/html/{ini}/{fin}"
)

SERIES = {
    "produccion_nacional":  ("PN01770AM", "Produccion Nacional (Indice 2007=100)"),
    "credito_privado":      ("PN00518MM", "Credito Sector Privado (Mill. S/)"),
    "credito_publico_neto": ("PN00881MM", "Credito Neto Sector Publico (Mill. S/)"),
    "tasa_interbancaria":   ("PN07819NM", "Tasa Interbancaria Prom. MN (% TEA)"),
    "m1_dinero":            ("PN00199MM", "M1 Dinero Sist. Financiero (Mill. S/)"),
}

START = "2004-1"
END   = "2025-12"

MESES_ES = {
    "Ene": 1, "Feb": 2, "Mar": 3, "Abr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Ago": 8, "Set": 9, "Sep": 9, "Oct": 10, "Nov": 11, "Dic": 12,
}

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def parsear_codigo_periodo(codigo):
    """'Ene04' -> Timestamp('2004-01-01'), 'Dic25' -> Timestamp('2025-12-01')"""
    m = re.match(r"^([A-Za-z]{3})(\d{2})$", codigo.strip())
    if not m:
        return pd.NaT
    mes_str, anio_str = m.group(1), m.group(2)
    mes = MESES_ES.get(mes_str.capitalize())
    if mes is None:
        return pd.NaT
    anio = 2000 + int(anio_str) if int(anio_str) < 50 else 1900 + int(anio_str)
    return pd.Timestamp(year=anio, month=mes, day=1)


def fetch_serie(key, code, label):
    url = BASE_HTML.format(code=code, ini=START, fin=END)
    print(f"  Descargando: {label} [{code}]...")
    try:
        r = requests.get(url, headers=HEADERS, timeout=40)
        r.raise_for_status()
    except Exception as e:
        print(f"  ERROR HTTP en {code}: {e}")
        return pd.DataFrame(columns=["fecha", key])

    soup = BeautifulSoup(r.content, "html.parser")
    tables = soup.find_all("table")
    if not tables:
        print(f"  ERROR: no se encontro tabla HTML en {code}")
        return pd.DataFrame(columns=["fecha", key])

    # La tabla de datos es la primera; extraer todas las celdas
    cells = [td.get_text(strip=True) for td in tables[0].find_all(["td", "th"])]

    # Buscar el patron periodo-valor: "Ene04", "8978", "Feb04", "9279", ...
    # Recorrer celdas buscando pares periodo->valor
    registros = []
    i = 0
    while i < len(cells) - 1:
        periodo_raw = cells[i]
        valor_raw   = cells[i + 1]
        if re.match(r"^[A-Za-z]{3}\d{2}$", periodo_raw):
            fecha = parsear_codigo_periodo(periodo_raw)
            try:
                val = float(valor_raw.replace(",", ".").replace(" ", ""))
            except ValueError:
                val = np.nan
            if pd.notna(fecha):
                registros.append({"fecha": fecha, key: val})
            i += 2
        else:
            i += 1

    if not registros:
        print(f"  ADVERTENCIA: no se extrajeron datos para {code}")
        return pd.DataFrame(columns=["fecha", key])

    df = pd.DataFrame(registros).sort_values("fecha").reset_index(drop=True)
    print(f"    -> {len(df)} observaciones extraidas "
          f"({df['fecha'].min().strftime('%b %Y')} - {df['fecha'].max().strftime('%b %Y')})")
    return df


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(script_dir, "datos_var_ipc.xlsx")

    # --- Descargar todas las series ---
    dfs = []
    for key, (code, label) in SERIES.items():
        df = fetch_serie(key, code, label)
        dfs.append(df)
        time.sleep(1.0)

    # --- Merge por fecha ---
    merged = dfs[0]
    for df in dfs[1:]:
        merged = pd.merge(merged, df, on="fecha", how="outer")

    merged = merged.sort_values("fecha").reset_index(drop=True)

    # Filtrar Jan 2004 - Dic 2025
    merged = merged[
        (merged["fecha"] >= "2004-01-01") &
        (merged["fecha"] <= "2025-12-31")
    ].reset_index(drop=True)

    # Columna periodo legible
    merged.insert(0, "periodo", merged["fecha"].dt.strftime("%Y-%m"))

    # --- Log-transformaciones ---
    trans = merged[["periodo", "fecha"]].copy()
    trans["ln_produccion"]    = np.log(merged["produccion_nacional"].where(merged["produccion_nacional"] > 0))
    trans["ln_credito_priv"]  = np.log(merged["credito_privado"].where(merged["credito_privado"] > 0))
    trans["credito_pub_neto"] = merged["credito_publico_neto"]       # puede ser negativo
    trans["tasa_interbancaria"] = merged["tasa_interbancaria"]
    trans["ln_m1"]            = np.log(merged["m1_dinero"].where(merged["m1_dinero"] > 0))

    # --- Metadata ---
    meta_rows = []
    transformaciones = [
        "Nivel directo (Indice 2007=100). En Stata: gen ln_prod=ln(produccion_nacional) luego D.ln_prod",
        "Nivel (Mill. S/). En Stata: gen ln_cred=ln(credito_privado) luego D.ln_cred",
        "Nivel (Mill. S/, puede ser negativo). Usar directamente o en diferencias",
        "Tasa % TEA. Usar directamente (ya estacionaria o cerca)",
        "Nivel (Mill. S/). En Stata: gen ln_m1=ln(m1_dinero) luego D.ln_m1",
    ]
    for (key, (code, label)), t in zip(SERIES.items(), transformaciones):
        meta_rows.append({
            "variable_excel":   key,
            "codigo_bcrp":      code,
            "descripcion_bcrp": label,
            "transformacion_var": t,
            "url_html": BASE_HTML.format(code=code, ini=START, fin=END),
        })
    meta = pd.DataFrame(meta_rows)

    # --- Exportar Excel ---
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="niveles", index=False)
        trans.to_excel(writer,  sheet_name="log_niveles", index=False)
        meta.to_excel(writer,   sheet_name="metadata", index=False)

    # --- Resumen ---
    print(f"\n{'='*55}")
    print(f"Archivo guardado: {out_path}")
    print(f"Observaciones: {len(merged)}  ({merged['periodo'].iloc[0]} a {merged['periodo'].iloc[-1]})")
    print(f"\nResumen por variable:")
    for col in list(SERIES.keys()):
        if col in merged.columns:
            n_ok   = merged[col].notna().sum()
            n_miss = merged[col].isna().sum()
            v_min  = merged[col].min()
            v_max  = merged[col].max()
            print(f"  {col:25s}: {n_ok:3d} validas | {n_miss:2d} faltantes "
                  f"| min={v_min:.1f}  max={v_max:.1f}")
    print(f"{'='*55}")


if __name__ == "__main__":
    main()
