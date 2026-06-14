"""
Script para descargar series macroeconómicas de Chile (mensual, desde 2000)
Variables: Brent, IMACEC, Riesgo País (EMBI), TPM, M1, IPC
"""

import requests
import pandas as pd
from io import BytesIO
import warnings
import sys
from datetime import datetime

warnings.filterwarnings('ignore')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'es-CL,es;q=0.9,en;q=0.8',
}

BOLETIN_BASE = "https://si3.bcentral.cl/estadisticas/Principal1/enlaces/Informes/BOLETIN/listado/"
START_DATE = "2000-01-01"
END_DATE = datetime.today().strftime("%Y-%m-%d")

results = {}  # nombre -> DataFrame con columna 'fecha' y columna de valor

# ─────────────────────────────────────────────────────────────────────────────
# 1. BRENT – EIA (US Energy Information Administration)
#    URL: https://www.eia.gov/dnav/pet/xls/PET_PRI_SPT_S1_M.xls
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 60)
print("1. Descargando Precio Brent (EIA)...")
try:
    url_brent = "https://www.eia.gov/dnav/pet/xls/PET_PRI_SPT_S1_M.xls"
    r = requests.get(url_brent, headers=HEADERS, timeout=60)
    r.raise_for_status()
    xls = pd.ExcelFile(BytesIO(r.content))
    # La hoja "Data 1" contiene los spot prices mensuales
    df_brent_raw = pd.read_excel(BytesIO(r.content), sheet_name="Data 1", skiprows=2)
    # Columnas: Date, WTI, Brent (Europe)
    # El nombre de la columna de Brent varía; buscarla
    brent_col = None
    for col in df_brent_raw.columns:
        if 'brent' in str(col).lower() or 'europe' in str(col).lower():
            brent_col = col
            break
    if brent_col is None:
        # Intentar por posición (col 2 suele ser Brent)
        brent_col = df_brent_raw.columns[1]

    df_brent = df_brent_raw[['Date', brent_col]].copy()
    df_brent.columns = ['fecha', 'brent_usd_bbl']
    df_brent['fecha'] = pd.to_datetime(df_brent['fecha'], errors='coerce')
    df_brent = df_brent.dropna(subset=['fecha'])
    df_brent = df_brent[(df_brent['fecha'] >= START_DATE) & (df_brent['fecha'] <= END_DATE)]
    df_brent['fecha'] = df_brent['fecha'].dt.to_period('M').dt.to_timestamp()
    df_brent = df_brent.sort_values('fecha').reset_index(drop=True)
    results['brent'] = df_brent
    print(f"   OK – {len(df_brent)} obs ({df_brent['fecha'].min().strftime('%Y-%m')} a {df_brent['fecha'].max().strftime('%Y-%m')})")
    print(f"   Fuente: Nivel en USD por barril (precio spot mensual, FOB)")
except Exception as e:
    print(f"   ERROR: {e}")
    results['brent'] = None

# ─────────────────────────────────────────────────────────────────────────────
# 2. IMACEC – Banco Central de Chile (Boletín Estadístico)
#    Archivo: BC001_g.xls
# ─────────────────────────────────────────────────────────────────────────────
print("\n2. Descargando IMACEC (BCCh Boletín)...")
try:
    url_imacec = BOLETIN_BASE + "BC001_g.xls"
    r = requests.get(url_imacec, headers=HEADERS, timeout=60)
    r.raise_for_status()
    xls = pd.ExcelFile(BytesIO(r.content))
    print(f"   Hojas disponibles: {xls.sheet_names}")

    # Probar con la primera hoja útil
    for sheet in xls.sheet_names:
        try:
            df_raw = pd.read_excel(BytesIO(r.content), sheet_name=sheet, header=None)
            print(f"   Hoja '{sheet}': {df_raw.shape}")
            print(df_raw.head(10).to_string())
            break
        except Exception as se:
            print(f"   Error en hoja '{sheet}': {se}")
except Exception as e:
    print(f"   ERROR descargando: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# 3. TPM – Banco Central de Chile (Boletín Estadístico)
#    Archivo: BB032_2_g.xls
# ─────────────────────────────────────────────────────────────────────────────
print("\n3. Descargando TPM (BCCh Boletín)...")
try:
    url_tpm = BOLETIN_BASE + "BB032_2_g.xls"
    r = requests.get(url_tpm, headers=HEADERS, timeout=60)
    r.raise_for_status()
    xls = pd.ExcelFile(BytesIO(r.content))
    print(f"   Hojas disponibles: {xls.sheet_names}")
    for sheet in xls.sheet_names:
        try:
            df_raw = pd.read_excel(BytesIO(r.content), sheet_name=sheet, header=None)
            print(f"   Hoja '{sheet}': {df_raw.shape}")
            print(df_raw.head(10).to_string())
            break
        except Exception as se:
            print(f"   Error: {se}")
except Exception as e:
    print(f"   ERROR descargando: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# 4. M1 – Banco Central de Chile (Boletín Estadístico)
#    Archivo: BB003.xls
# ─────────────────────────────────────────────────────────────────────────────
print("\n4. Descargando M1 (BCCh Boletín)...")
try:
    url_m1 = BOLETIN_BASE + "BB003.xls"
    r = requests.get(url_m1, headers=HEADERS, timeout=60)
    r.raise_for_status()
    xls = pd.ExcelFile(BytesIO(r.content))
    print(f"   Hojas disponibles: {xls.sheet_names}")
    for sheet in xls.sheet_names:
        try:
            df_raw = pd.read_excel(BytesIO(r.content), sheet_name=sheet, header=None)
            print(f"   Hoja '{sheet}': {df_raw.shape}")
            print(df_raw.head(10).to_string())
            break
        except Exception as se:
            print(f"   Error: {se}")
except Exception as e:
    print(f"   ERROR descargando: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# 5. IPC – Banco Central de Chile (serie empalmada)
# ─────────────────────────────────────────────────────────────────────────────
print("\n5. Buscando IPC empalmado en BCCh...")
# El BCCh BDE tiene IPC empalmado; intentamos API pública
# Series IDs encontrados: IPC_EMP_2018, IPC_EMP_2023
# API REST del BCCh (requiere credenciales, pero probamos)
# Alternativa: usar la API de FRED o World Bank

# Intentar FRED (CHLCPIALLMINMEI = Chile CPI All Items Monthly)
print("   Intentando FRED API...")
try:
    url_fred_ipc = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=CHLCPIALLMINMEI"
    r = requests.get(url_fred_ipc, headers=HEADERS, timeout=30)
    r.raise_for_status()
    from io import StringIO
    df_ipc = pd.read_csv(StringIO(r.text))
    df_ipc.columns = ['fecha', 'ipc_indice']
    df_ipc['fecha'] = pd.to_datetime(df_ipc['fecha'], errors='coerce')
    df_ipc = df_ipc.dropna()
    df_ipc = df_ipc[(df_ipc['fecha'] >= START_DATE) & (df_ipc['fecha'] <= END_DATE)]
    df_ipc = df_ipc.sort_values('fecha').reset_index(drop=True)
    results['ipc'] = df_ipc
    print(f"   OK (FRED) – {len(df_ipc)} obs ({df_ipc['fecha'].min().strftime('%Y-%m')} a {df_ipc['fecha'].max().strftime('%Y-%m')})")
    print(f"   Fuente: Índice (base variable OCDE, mensual)")
except Exception as e:
    print(f"   FRED falló: {e}")
    results['ipc'] = None

# ─────────────────────────────────────────────────────────────────────────────
# 6. EMBI (Riesgo País) – BCRP Peru (tienen datos de Chile también)
# ─────────────────────────────────────────────────────────────────────────────
print("\n6. Buscando EMBI Chile (BCRP Peru API)...")
try:
    # API pública del Banco Central del Perú - tiene EMBIG para varios países
    # Intentar con la URL de series mensuales del BCRP
    url_bcrp = "https://estadisticas.bcrp.gob.pe/estadisticas/series/mensuales/resultados/PN01262PM/html"
    r = requests.get(url_bcrp, headers=HEADERS, timeout=30)
    print(f"   Status BCRP: {r.status_code}")
    if r.status_code == 200:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(r.text, 'html.parser')
        print(soup.text[:500])
except Exception as e:
    print(f"   BCRP falló: {e}")

print("\n" + "=" * 60)
print("Exploración completada. Revisa los resultados arriba.")
print("Ahora ajustaremos el parseo según la estructura real de los archivos.")
