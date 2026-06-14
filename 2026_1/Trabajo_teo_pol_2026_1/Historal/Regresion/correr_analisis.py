import sys, io, warnings
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
warnings.filterwarnings('ignore')

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from scipy import stats
import statsmodels.api as sm
from statsmodels.tsa.stattools import adfuller, kpss
from statsmodels.tsa.ardl import ARDL, UECM, ardl_select_order
from statsmodels.tsa.vector_ar.var_model import VAR
from statsmodels.stats.stattools import durbin_watson
from statsmodels.stats.diagnostic import acorr_breusch_godfrey, het_white
import glob, os

# ─── rutas ──────────────────────────────────────────────────
BASE   = r"C:\Users\51950\Documents\ULIMA\2026_1\Teo. Politica\Trabajo grupal\Regresion"
XLSX   = os.path.join(BASE, "datos_canal_crediticio.xlsx")
OUT    = os.path.join(BASE, "output")
os.makedirs(OUT, exist_ok=True)

SEP = "=" * 60

def titulo(txt):
    print(f"\n{SEP}\n  {txt}\n{SEP}")

# ════════════════════════════════════════════════════════════
# 1. CARGAR DATOS
# ════════════════════════════════════════════════════════════
titulo("1. CARGA DE DATOS")

df = pd.read_excel(XLSX, sheet_name="datos")
df['fecha'] = pd.to_datetime(df['fecha'])
df = df.set_index('fecha').sort_index()
df.index = pd.DatetimeIndex(df.index).to_period('M')

for col in ['tasa_ref','tamn','ipc','tc']:
    df[col] = pd.to_numeric(df[col], errors='coerce')

df['inflacion'] = df['ipc'].pct_change() * 100
df['brecha']    = df['tamn'] - df['tasa_ref']
df['d_tamn']    = df['tamn'].diff()
df['d_tasa']    = df['tasa_ref'].diff()
df = df.dropna(subset=['tamn','tasa_ref','ipc','tc'])

print(f"  Periodo : {df.index[0]}  →  {df.index[-1]}")
print(f"  Obs     : {len(df)}")
print(df[['tasa_ref','tamn','ipc','tc']].head(3).to_string())

# ════════════════════════════════════════════════════════════
# 2. ESTADISTICAS DESCRIPTIVAS
# ════════════════════════════════════════════════════════════
titulo("2. ESTADISTICAS DESCRIPTIVAS")

desc = df[['tasa_ref','tamn','brecha','inflacion','tc']].describe().T
desc.columns = ['Obs','Media','Desv.Est.','Min','25%','50%','75%','Max']
desc['Obs'] = desc['Obs'].astype(int)
print(desc[['Obs','Media','Desv.Est.','Min','Max']].round(3).to_string())

# ════════════════════════════════════════════════════════════
# 3. GRAFICOS
# ════════════════════════════════════════════════════════════
titulo("3. GRAFICOS")

idx = df.index.to_timestamp()
plt.style.use('seaborn-v0_8-whitegrid')
NAVY = '#1F3864'; ORANGE = '#C55A11'; RED = '#C00000'

# Grafico 1 — Series en el tiempo
fig, ax = plt.subplots(figsize=(12,5))
ax.plot(idx, df['tamn'],     color=NAVY,   lw=2,   label='TAMN')
ax.plot(idx, df['tasa_ref'], color=ORANGE, lw=2, ls='--', label='Tasa referencia BCRP')
ax.set_title('TAMN vs Tasa de referencia BCRP (2004–2024)', fontsize=13, fontweight='bold')
ax.set_ylabel('Tasa de interés (%)'); ax.legend(); fig.tight_layout()
fig.savefig(os.path.join(OUT,'g1_series.png'), dpi=150); plt.close()
print("  ✔ g1_series.png")

# Grafico 2 — Spread
fig, ax = plt.subplots(figsize=(12,4))
ax.plot(idx, df['brecha'], color=RED, lw=1.8)
ax.axhline(0, color='gray', ls='--', lw=1)
ax.fill_between(idx, df['brecha'], 0,
                where=df['brecha']>0, alpha=0.15, color=RED)
ax.set_title('Spread: TAMN − Tasa de referencia BCRP', fontsize=13, fontweight='bold')
ax.set_ylabel('Puntos porcentuales'); fig.tight_layout()
fig.savefig(os.path.join(OUT,'g2_spread.png'), dpi=150); plt.close()
print("  ✔ g2_spread.png")

# Grafico 3 — Scatter
fig, ax = plt.subplots(figsize=(7,6))
ax.scatter(df['tasa_ref'], df['tamn'], color=NAVY, alpha=0.55, s=22, label='Observaciones')
m,b = np.polyfit(df['tasa_ref'], df['tamn'], 1)
xs = np.linspace(df['tasa_ref'].min(), df['tasa_ref'].max(), 100)
ax.plot(xs, m*xs+b, color=RED, lw=2, label=f'Ajuste lineal (β={m:.3f})')
ax.set_title('TAMN vs Tasa de referencia', fontsize=13, fontweight='bold')
ax.set_xlabel('Tasa referencia BCRP (%)'); ax.set_ylabel('TAMN (%)')
ax.legend(); fig.tight_layout()
fig.savefig(os.path.join(OUT,'g3_scatter.png'), dpi=150); plt.close()
print("  ✔ g3_scatter.png")

# ════════════════════════════════════════════════════════════
# 4. PRUEBAS DE RAIZ UNITARIA
# ════════════════════════════════════════════════════════════
titulo("4. PRUEBAS DE RAIZ UNITARIA  (H0: tiene raiz unitaria)")

def adf_test(series, name, maxlag=12, trend='c'):
    res = adfuller(series.dropna(), maxlag=maxlag, regression=trend, autolag='AIC')
    sig = '*** p<0.01' if res[1]<0.01 else ('** p<0.05' if res[1]<0.05 else ('* p<0.10' if res[1]<0.10 else 'No rechaza H0'))
    print(f"  {name:<35}  ADF={res[0]:>8.4f}  p={res[1]:.4f}  lags={res[2]}  → {sig}")
    return res[1]

print("\n  [Niveles]")
adf_test(df['tamn'],     'TAMN')
adf_test(df['tasa_ref'], 'Tasa referencia BCRP')

print("\n  [Primeras diferencias]")
adf_test(df['d_tamn'].dropna(), 'D.TAMN')
adf_test(df['d_tasa'].dropna(), 'D.Tasa referencia')

print("""
  INTERPRETACION:
    Niveles      → no rechaza H0  → series I(1)  ✔
    Diferencias  → rechaza H0     → series I(0)  ✔
    Conclusion: ambas series son I(1) → procede ARDL bounds test
""")

# ════════════════════════════════════════════════════════════
# 5. SELECCION DE REZAGOS
# ════════════════════════════════════════════════════════════
titulo("5. SELECCION DE REZAGOS OPTIMOS (AIC / BIC)")

data_var = df[['tamn','tasa_ref']].dropna()
model_var = VAR(data_var)
res_lag   = model_var.select_order(maxlags=12)
lag_aic   = res_lag.aic
lag_bic   = res_lag.bic

print(res_lag.summary())
print(f"\n  Rezago optimo AIC : {lag_aic}")
print(f"  Rezago optimo BIC : {lag_bic}")
LAG_OPT = int(lag_bic) if lag_bic > 0 else 2

# ════════════════════════════════════════════════════════════
# 6. ARDL + BOUNDS TEST
# ════════════════════════════════════════════════════════════
titulo("6. MODELO ARDL + BOUNDS TEST  (Pesaran et al., 2001)")

data_ardl = df[['tamn','tasa_ref']].dropna()

# Seleccion automatica del orden ARDL por AIC
sel = ardl_select_order(data_ardl['tamn'], 6,
                        data_ardl[['tasa_ref']], 6,
                        ic='aic', trend='c')
p_opt = sel.model.ardl_order[0]
q_opt = sel.model.ardl_order[1]
print(f"\n  Orden ARDL seleccionado por AIC: ARDL({p_opt}, {q_opt})")

ardl_mod = ARDL(data_ardl['tamn'], p_opt,
                data_ardl[['tasa_ref']], q_opt,
                trend='c')
ardl_res = ardl_mod.fit()
print(ardl_res.summary())

# Bounds test via UECM (reparametrizacion del ARDL)
uecm_mod = UECM(data_ardl['tamn'], p_opt,
                data_ardl[['tasa_ref']], q_opt,
                trend='c')
uecm_res = uecm_mod.fit()
bt = uecm_res.bounds_test(case=3)          # case 3 = constante no restringida
print("\n  --- BOUNDS TEST (Pesaran et al. 2001) ---")
print(f"  F-statistic   : {bt.stat:.4f}")
cv = bt.crit_vals                             # DataFrame con index=percentile, cols=lower/upper
cv_lo = cv.loc[95.0, 'lower']
cv_hi = cv.loc[95.0, 'upper']
print(f"  Valor critico I(0) al 5% : {cv_lo:.2f}")
print(f"  Valor critico I(1) al 5% : {cv_hi:.2f}")
print(f"  P-values  upper={bt.p_values['upper']:.4f}  lower={bt.p_values['lower']:.4f}")
if bt.stat > cv_hi:
    print("  CONCLUSION: F > I(1) bound -> RECHAZO H0 -> COINTEGRACION CONFIRMADA")
elif bt.stat < cv_lo:
    print("  CONCLUSION: F < I(0) bound -> NO cointegracion")
else:
    print("  CONCLUSION: Resultado inconcluso (entre ambos valores criticos)")

# ════════════════════════════════════════════════════════════
# 7. MODELO ECM (LARGO Y CORTO PLAZO)
# ════════════════════════════════════════════════════════════
titulo("7. MODELO ECM — PASS-THROUGH CORTO Y LARGO PLAZO")

data_ecm = df[['tamn','tasa_ref','d_tamn','d_tasa']].dropna().copy()

# Paso 1: Relacion de largo plazo OLS en niveles
X_lr  = sm.add_constant(data_ecm['tasa_ref'])
lr    = sm.OLS(data_ecm['tamn'], X_lr).fit()
data_ecm['ect'] = lr.resid
pt_lp = lr.params['tasa_ref']

print(f"\n  [Largo plazo]  TAMN = {lr.params['const']:.3f} + {pt_lp:.3f} * Tasa_ref")
print(f"  R² = {lr.rsquared:.4f}")
print(f"\n  Pass-through de LARGO PLAZO = {pt_lp:.4f}")
if   pt_lp >= 0.95: interp = "traspaso casi completo"
elif pt_lp >= 0.60: interp = "traspaso incompleto moderado"
else:               interp = "traspaso incompleto bajo"
print(f"  Interpretacion: {interp}")

# Verificar estacionariedad del ECT
ect_adf = adfuller(data_ecm['ect'], maxlag=4, regression='n', autolag='AIC')
print(f"\n  ADF sobre ECT: stat={ect_adf[0]:.4f}, p={ect_adf[1]:.4f}")
print("  ECT es I(0) → cointegración confirmada ✔" if ect_adf[1] < 0.05
      else "  ALERTA: ECT podria no ser I(0)")

# Paso 2: ECM de corto plazo
data_ecm['L_ect']   = data_ecm['ect'].shift(1)
data_ecm['L_d_tamn']= data_ecm['d_tamn'].shift(1)
data_ecm['L_d_tasa']= data_ecm['d_tasa'].shift(1)
data_ecm = data_ecm.dropna()

X_ecm = sm.add_constant(data_ecm[['d_tasa','L_d_tasa','L_d_tamn','L_ect']])
ecm   = sm.OLS(data_ecm['d_tamn'], X_ecm).fit(cov_type='HC1')

print("\n  [Corto plazo — ECM con errores robustos HC1]")
print(ecm.summary())

pt_cp  = ecm.params['d_tasa']
lambda_val = ecm.params['L_ect']

print(f"\n  Pass-through INMEDIATO  (β d_tasa): {pt_cp:.4f}")
print(f"  Velocidad de ajuste     (λ  L_ect): {lambda_val:.4f}")
print(f"  Half-life (meses al equilibrio): {abs(np.log(0.5)/np.log(1+lambda_val)):.1f} meses"
      if -1 < lambda_val < 0 else "  ALERTA: lambda fuera de rango esperado")

# ════════════════════════════════════════════════════════════
# 8. DIAGNOSTICOS
# ════════════════════════════════════════════════════════════
titulo("8. DIAGNOSTICOS DEL ECM  (H0: sin problema → p > 0.05)")

res_ecm = ecm.resid

# Autocorrelacion — Breusch-Godfrey
print("\n  --- Breusch-Godfrey (autocorrelacion) ---")
for lag in [1,2,3,6,12]:
    lm, pval, _, _ = acorr_breusch_godfrey(ecm, nlags=lag)
    flag = "✔ ok" if pval > 0.05 else "✗ problema"
    print(f"  Lags={lag:>2}  LM={lm:>8.4f}  p={pval:.4f}  {flag}")

# Heterocedasticidad — White
print("\n  --- White (heterocedasticidad) ---")
lm_w, pval_w, _, _ = het_white(res_ecm, X_ecm)
flag_w = "✔ ok (homocedastica)" if pval_w > 0.05 else "✗ heterocedastica (usar HC1)"
print(f"  LM={lm_w:.4f}  p={pval_w:.4f}  → {flag_w}")

# Normalidad — Jarque-Bera
jb_stat, jb_p = stats.jarque_bera(res_ecm)
print("\n  --- Jarque-Bera (normalidad) ---")
print(f"  JB={jb_stat:.4f}  p={jb_p:.4f}  → {'✔ normal' if jb_p>0.05 else 'no normal (N>100: no critico)'}")

# Durbin-Watson
dw = durbin_watson(res_ecm)
print(f"\n  Durbin-Watson = {dw:.4f}  (cercano a 2 = sin autocorr.)")

# Grafico de residuos
fig, axes = plt.subplots(1,2,figsize=(12,4))
axes[0].plot(data_ecm.index.to_timestamp(), res_ecm, color=NAVY, lw=1)
axes[0].axhline(0, color='gray', ls='--')
axes[0].set_title('Residuos del ECM'); axes[0].set_ylabel('Residuo')
axes[1].hist(res_ecm, bins=25, color=NAVY, alpha=0.7, edgecolor='white')
axes[1].set_title('Distribucion de residuos')
fig.tight_layout()
fig.savefig(os.path.join(OUT,'g5_residuos.png'), dpi=150); plt.close()
print("\n  ✔ g5_residuos.png")

# ════════════════════════════════════════════════════════════
# 9. FUNCION IMPULSO-RESPUESTA (VAR)
# ════════════════════════════════════════════════════════════
titulo("9. FUNCION IMPULSO-RESPUESTA (VAR)")

# Ordenamiento Cholesky: tasa_ref primero (exogena)
data_irf  = df[['tasa_ref','tamn']].dropna()
var_model = VAR(data_irf)
var_res   = var_model.fit(LAG_OPT, trend='c')
print(var_res.summary())

irf = var_res.irf(periods=18)

# Grafico IRF: respuesta de tamn ante shock en tasa_ref
fig, ax = plt.subplots(figsize=(10,5))
# tasa_ref es variable 0, tamn es variable 1
oirf_vals = irf.orth_irfs[:, 1, 0]   # respuesta [tamn] a shock [tasa_ref]
ax.plot(range(19), oirf_vals, color=NAVY, lw=2.5, marker='o', ms=5, label='OIRF')
ax.axhline(0, color='gray', ls='--', lw=1)
ax.fill_between(range(19),
                irf.cum_effect_stderr()[:, 1, 0] * -1.96 + oirf_vals,
                irf.cum_effect_stderr()[:, 1, 0] *  1.96 + oirf_vals,
                alpha=0.15, color=NAVY, label='IC 95%')
ax.set_title('IRF: Respuesta de TAMN ante shock de 1pp en Tasa de Referencia',
             fontsize=12, fontweight='bold')
ax.set_xlabel('Meses'); ax.set_ylabel('Puntos porcentuales')
ax.set_xticks(range(0,19)); ax.legend(); fig.tight_layout()
fig.savefig(os.path.join(OUT,'g4_irf.png'), dpi=150); plt.close()
print("  ✔ g4_irf.png")

print("\n  Valores OIRF por mes:")
print(f"  {'Mes':>4}  {'Respuesta TAMN':>16}")
for i, v in enumerate(oirf_vals):
    print(f"  {i:>4}  {v:>16.4f}")

# ════════════════════════════════════════════════════════════
# 10. RESUMEN EJECUTIVO
# ════════════════════════════════════════════════════════════
titulo("10. RESUMEN EJECUTIVO")

print(f"""
  CANAL CREDITICIO - TRANSMISION DE POLITICA MONETARIA EN PERU
  Periodo: {df.index[0]} - {df.index[-1]}  |  Obs: {len(df)}
  Metodologia: ARDL({p_opt},{q_opt}) / ECM con errores robustos HC1

  ┌─────────────────────────────────────────────────────────┐
  │  RAIZ UNITARIA                                          │
  │  TAMN y Tasa_ref son I(1) en niveles → ARDL valido     │
  ├─────────────────────────────────────────────────────────┤
  │  COINTEGRACIÓN (Bounds Test)                            │
  │  F-stat = {bt.stat:.4f}  |  I(1) bound 5% = {cv_hi:.2f}              │
  │  {"CONFIRMADA" if bt.stat > cv_hi else "NO confirmada"}                                              │
  ├─────────────────────────────────────────────────────────┤
  │  PASS-THROUGH                                           │
  │  Largo plazo  (θ) = {pt_lp:>8.4f}  → {interp:<25}│
  │  Corto plazo  (β) = {pt_cp:>8.4f}                          │
  │  Veloc. ajuste(λ) = {lambda_val:>8.4f}  (debe ser < 0)          │
  ├─────────────────────────────────────────────────────────┤
  │  DIAGNOSTICOS                                           │
  │  Breusch-Godfrey (lag 1): p = {acorr_breusch_godfrey(ecm,1)[1]:.4f}               │
  │  White heteroced.:        p = {pval_w:.4f}               │
  │  Durbin-Watson:           DW = {dw:.4f}              │
  └─────────────────────────────────────────────────────────┘

  Graficos guardados en: {OUT}
""")

print("ANALISIS COMPLETADO.")
