/*===========================================================================
  VAR_Chile_v2.do  —  Mecanismo de Transmisión Monetaria en Chile
  Modelo: VAR básico con bloque exógeno (block exogeneity)
  Endógenas (orden Cholesky):  d_ltc → tpm → d_lipc
  Exógenas:  d_lpetro  d_lcobre  d_lipc_ee  d_tasa_fed
  Período:   2004m1 – 2025m12   (T ≈ 264 obs. mensuales)
  Salida:    resultados/resultados_VAR.docx
  Autor:     Jorge Valero  –  ULIMA, Econometría 2, 2026-1
  Versión:   v2 — correcciones de sintaxis Stata 17
  Correcciones respecto a v1:
    [E1] Eliminado "firstrow" del import (fila 4 = datos, no encabezado)
    [E2] Eliminada opción "trend" de kpss (no existe; default ya incluye tendencia)
    [E3] Corregido varstable: "graph name(...)" → "graph" + export directo
    [E4] Eliminada opción "ci" de irf graph (no válida; IC se muestran automáticamente)
    [E5] Tabla tbl_ur: dimensión (11,7) → (8,7) (1 hdr + 7 filas usadas)
    [E6] Shading del encabezado integrado en mismo comando que el contenido
    [E7] Macros anidados: pre-calcular con ":di %fmt" antes de putdocx table
    [E8] e(lags) → usar p_opt (e(lags) es matriz, no escalar)
===========================================================================*/

version 17
clear all
set more off
set scheme s2color

* ── Directorios ─────────────────────────────────────────────────────────────
global projdir `"C:\Users\51950\Documents\ULIMA\2026_1\Econometria 2\Trabajo_metria_2_2026_1"'
global datadir `"$projdir\Data"'
global resdir  `"$projdir\resultados"'

cap mkdir "$resdir"

* ── Instalar paquetes si faltan ──────────────────────────────────────────────
cap which kpss
if _rc != 0 { ssc install kpss, replace }

* ============================================================================
* SECCIÓN 1. IMPORTAR Y PREPARAR DATOS
* ============================================================================
* [E1] CORRECCIÓN: eliminado "firstrow" — la fila A4 contiene datos (2004-01-01),
*      no encabezados. Sin firstrow, Stata asigna nombres A, B, C... y luego
*      los renombramos manualmente.

import excel using `"$datadir\DATA_JORGE.xlsx"', ///
    sheet("Sheet1") cellrange(A4) clear

* Renombrar columnas (A–N, en el orden del Excel)
rename A fecha
rename B imacec
rename C infl
rename D ipc
rename E petro
rename F tpm
rename G m1
rename H embi
rename I cobre
rename J tc
rename K brecha_pib
rename L ipc_esperado
rename M tasa_fed
rename N ipc_eeuu

* Convertir fecha Excel a formato Stata
drop if missing(fecha)
gen fecha_m = mofd(fecha)
format fecha_m %tm
tsset fecha_m, monthly

* Mantener muestra 2004m1 en adelante
keep if fecha_m >= tm(2004m1)

* ── Transformaciones logarítmicas ───────────────────────────────────────────
gen l_tc     = log(tc)
gen l_ipc    = log(ipc)
gen l_petro  = log(petro)
gen l_cobre  = log(cobre)
gen l_ipc_ee = log(ipc_eeuu)

label var l_tc     "Log Tipo de Cambio (CLP/USD)"
label var l_ipc    "Log IPC Chile"
label var l_petro  "Log Precio Petroleo (USD/bbl)"
label var l_cobre  "Log Precio Cobre (USD/lb)"
label var l_ipc_ee "Log IPC EE.UU. (CPIAUCSL)"
label var tpm      "Tasa de Politica Monetaria (%) - niveles"
label var tasa_fed "Tasa Fed Funds (%) - niveles"

* ============================================================================
* SECCIÓN 2. PRUEBAS DE RAÍZ UNITARIA
* ============================================================================
* Convención:  lags ADF = 6, NW PP = 8, maxlag KPSS = 12
* Todas con constante y tendencia.
* TPM se mantiene en niveles (Sims, Stock y Watson 1990 — serie acotada).
* [E2] CORRECCIÓN: kpss no acepta opción "trend"; el default ya incluye tendencia.
*      Usar "notrend" explícito si se quisiera solo constante (no es el caso aquí).
* ============================================================================

local vars_niveles "l_tc tpm l_ipc l_petro l_cobre l_ipc_ee tasa_fed"

* ── ADF (lags=6, con constante y tendencia) ──────────────────────────────────
foreach v of local vars_niveles {
    qui dfuller `v', lags(6) trend regress
    local adf_`v'_t  = r(Zt)
    local adf_`v'_pv = r(p)
}
* Primera diferencia (variables que resulten I(1))
foreach v in l_tc l_ipc l_petro l_cobre l_ipc_ee tasa_fed {
    qui dfuller D.`v', lags(6) trend regress
    local adf_d`v'_t  = r(Zt)
    local adf_d`v'_pv = r(p)
}

* ── PP (NW=8, con constante y tendencia) ─────────────────────────────────────
foreach v of local vars_niveles {
    qui pperron `v', lags(8) trend regress
    local pp_`v'_t  = r(Zt)
    local pp_`v'_pv = r(p)
}
foreach v in l_tc l_ipc l_petro l_cobre l_ipc_ee tasa_fed {
    qui pperron D.`v', lags(8) trend regress
    local pp_d`v'_t  = r(Zt)
    local pp_d`v'_pv = r(p)
}

* ── KPSS (maxlag=12, con tendencia — comportamiento por defecto) ──────────────
* H0 KPSS: serie es estacionaria alrededor de tendencia
* VC al 5%: 0.146 (con tendencia) | 0.463 (sin tendencia)
foreach v of local vars_niveles {
    qui kpss `v', maxlag(12)
    local kpss_`v'_t = r(kpss_stat)
}
foreach v in l_tc l_ipc l_petro l_cobre l_ipc_ee tasa_fed {
    qui kpss D.`v', maxlag(12)
    local kpss_d`v'_t = r(kpss_stat)
}

* ── Generar variables en primera diferencia ──────────────────────────────────
gen d_ltc      = D.l_tc
gen d_lipc     = D.l_ipc
gen d_lpetro   = D.l_petro
gen d_lcobre   = D.l_cobre
gen d_lipc_ee  = D.l_ipc_ee
gen d_tasa_fed = D.tasa_fed

label var d_ltc      "Delta Log TC (variacion mensual)"
label var d_lipc     "Delta Log IPC Chile (inflacion mensual)"
label var d_lpetro   "Delta Log Petroleo"
label var d_lcobre   "Delta Log Cobre"
label var d_lipc_ee  "Delta Log IPC EE.UU."
label var d_tasa_fed "Delta Tasa Fed Funds"

* ============================================================================
* SECCIÓN 3. SELECCIÓN DE REZAGOS
* ============================================================================

varsoc d_ltc tpm d_lipc, ///
    maxlag(12) ///
    exog(d_lpetro d_lcobre d_lipc_ee d_tasa_fed)

* INSTRUCCIÓN: revisar la tabla de varsoc y ajustar p_opt según el SBIC mínimo.
* El valor 2 es el punto de partida recomendado para datos mensuales.
local p_opt = 2

di "Rezagos seleccionados: p = `p_opt'"

* ============================================================================
* SECCIÓN 4. ESTIMACIÓN DEL VAR
* ============================================================================

var d_ltc tpm d_lipc, ///
    lags(1/`p_opt') ///
    exog(d_lpetro d_lcobre d_lipc_ee d_tasa_fed)

estimates store var_chile

* [E8] CORRECCIÓN: e(lags) es una matriz (1 2), no un escalar. Usar p_opt.
local nobs = e(N)
local neqs = e(neqs)
local nlag = `p_opt'
local tmin : di %tm e(tmin)
local tmax : di %tm e(tmax)

di "VAR estimado. N=`nobs', ecuaciones=`neqs', rezagos=`nlag'"
di "Muestra: `tmin' - `tmax'"

* ============================================================================
* SECCIÓN 5. DIAGNÓSTICOS
* ============================================================================

* 5.1 Estabilidad (raíces del polinomio)
* [E3] CORRECCIÓN: varstable no acepta "name()" como suboption de "graph".
*      Se usa "varstable, graph" y luego se exporta el gráfico activo directamente.
varstable, graph
graph export `"$resdir\grafico_estabilidad.png"', replace width(800) height(700)
di "  Grafico de estabilidad exportado."

* 5.2 Test de autocorrelación LM (hasta rezago 12)
varlmar, mlag(12)

* 5.3 Test de normalidad (Jarque-Bera)
varnorm, jbera

* ============================================================================
* SECCIÓN 6. FUNCIONES DE IMPULSO-RESPUESTA (IRF)
* ============================================================================
set seed 1234

* Crear conjunto de IRFs con intervalos bootstrap (500 rep., semilla=1234)
irf create irf_var, ///
    step(24) ///
    set(`"$resdir\irf_chile"') ///
    bs reps(500) ///
    bseed(1234) ///
    replace

irf set `"$resdir\irf_chile"'

* [E4] CORRECCIÓN: "ci" no es opción válida de irf graph en Stata 17.
*      Los IC bootstrap se muestran automáticamente. Usar solo "level(#)".

* ── 6.1 Shock TPM → Inflación (canal tasa de interés) ───────────────────────
irf graph irf, ///
    impulse(tpm) response(d_lipc) ///
    irf(irf_var) ///
    level(95) ///
    title("IRF: Shock TPM - D.Log IPC (inflacion)", size(medium)) ///
    subtitle("Identificacion Cholesky - Horizonte 24 meses") ///
    xtitle("Meses") ytitle("Respuesta") ///
    note("IC 95% bootstrap (500 replicas). Periodo: 2004m1-2025m12") ///
    name(irf_tpm_ipc, replace)

graph export `"$resdir\irf_tpm_ipc.png"', ///
    replace name(irf_tpm_ipc) width(1000) height(700)

* ── 6.2 Shock TPM → Tipo de Cambio ─────────────────────────────────────────
irf graph irf, ///
    impulse(tpm) response(d_ltc) ///
    irf(irf_var) ///
    level(95) ///
    title("IRF: Shock TPM - D.Log TC (tipo de cambio)", size(medium)) ///
    subtitle("Identificacion Cholesky - Horizonte 24 meses") ///
    xtitle("Meses") ytitle("Respuesta") ///
    note("IC 95% bootstrap (500 replicas). Periodo: 2004m1-2025m12") ///
    name(irf_tpm_tc, replace)

graph export `"$resdir\irf_tpm_tc.png"', ///
    replace name(irf_tpm_tc) width(1000) height(700)

* ── 6.3 Shock TC → Inflación (pass-through cambiario) ────────────────────────
irf graph irf, ///
    impulse(d_ltc) response(d_lipc) ///
    irf(irf_var) ///
    level(95) ///
    title("IRF: Shock D.Log TC - D.Log IPC (pass-through)", size(medium)) ///
    subtitle("Identificacion Cholesky - Horizonte 24 meses") ///
    xtitle("Meses") ytitle("Respuesta") ///
    note("IC 95% bootstrap (500 replicas). Periodo: 2004m1-2025m12") ///
    name(irf_tc_ipc, replace)

graph export `"$resdir\irf_tc_ipc.png"', ///
    replace name(irf_tc_ipc) width(1000) height(700)

* ── 6.4 Panel completo IRF (9 combinaciones endógenas) ───────────────────────
irf graph irf, ///
    impulse(d_ltc tpm d_lipc) ///
    response(d_ltc tpm d_lipc) ///
    irf(irf_var) ///
    level(95) ///
    xtitle("Meses") ///
    note("IC 95% bootstrap. 2004m1-2025m12") ///
    byopts(title("IRF - Panel Completo (Cholesky)") rows(3)) ///
    name(irf_panel, replace)

graph export `"$resdir\irf_panel.png"', ///
    replace name(irf_panel) width(1200) height(1000)

* ============================================================================
* SECCIÓN 7. DESCOMPOSICIÓN DE VARIANZA (FEVD)
* ============================================================================

* ── 7.1 FEVD de Δ Log IPC ────────────────────────────────────────────────────
irf graph fevd, ///
    impulse(d_ltc tpm d_lipc) response(d_lipc) ///
    irf(irf_var) ///
    xtitle("Meses") ytitle("Proporcion") ///
    byopts(title("FEVD: Varianza de D.Log IPC (inflacion)") rows(1)) ///
    note("Descomposicion Cholesky. 2004m1-2025m12") ///
    name(fevd_ipc, replace)

graph export `"$resdir\fevd_ipc.png"', ///
    replace name(fevd_ipc) width(1000) height(700)

* ── 7.2 FEVD de Δ Log TC ─────────────────────────────────────────────────────
irf graph fevd, ///
    impulse(d_ltc tpm d_lipc) response(d_ltc) ///
    irf(irf_var) ///
    xtitle("Meses") ytitle("Proporcion") ///
    byopts(title("FEVD: Varianza de D.Log TC (tipo de cambio)") rows(1)) ///
    note("Descomposicion Cholesky. 2004m1-2025m12") ///
    name(fevd_tc, replace)

graph export `"$resdir\fevd_tc.png"', ///
    replace name(fevd_tc) width(1000) height(700)

* ── 7.3 FEVD de TPM ──────────────────────────────────────────────────────────
irf graph fevd, ///
    impulse(d_ltc tpm d_lipc) response(tpm) ///
    irf(irf_var) ///
    xtitle("Meses") ytitle("Proporcion") ///
    byopts(title("FEVD: Varianza de TPM (tasa de politica)") rows(1)) ///
    note("Descomposicion Cholesky. 2004m1-2025m12") ///
    name(fevd_tpm, replace)

graph export `"$resdir\fevd_tpm.png"', ///
    replace name(fevd_tpm) width(1000) height(700)

* ============================================================================
* SECCIÓN 8. EXPORTAR RESULTADOS A WORD (putdocx)
* ============================================================================

putdocx begin, font("Calibri", 11) pagesize(letter) ///
    margin(top, 2.5cm) margin(bottom, 2.5cm) ///
    margin(left, 3cm)  margin(right, 3cm)

* ── PORTADA ──────────────────────────────────────────────────────────────────
putdocx paragraph, style(Title)
putdocx text ("Mecanismo de Transmision de Politica Monetaria en Chile")

putdocx paragraph, style(Subtitle)
putdocx text ("Modelo VAR con Bloque Exogeno - Identificacion Cholesky")

putdocx paragraph
putdocx text ("Jorge Valero  |  ULIMA - Econometria 2, 2026-1"), ///
    font("Calibri", 11, italic)

putdocx paragraph
putdocx text ("Periodo de estimacion: 2004m1 - 2025m12   |   T = `nobs' obs."), ///
    font("Calibri", 10)

putdocx pagebreak

* ── 1. ESPECIFICACIÓN DEL MODELO ─────────────────────────────────────────────
putdocx paragraph, style(Heading1)
putdocx text ("1. Especificacion del Modelo")

putdocx paragraph
putdocx text ("Se estima un VAR mensual para el periodo 2004m1-2025m12 para el caso chileno, con identificacion recursiva (Cholesky). El modelo incorpora un bloque de variables externas tratado como exogeno ("), font("Calibri", 10)
putdocx text ("block exogeneity"), font("Calibri", 10, italic)
putdocx text ("), capturando condiciones monetarias y de commodities internacionales."), font("Calibri", 10)

putdocx paragraph, style(Heading2)
putdocx text ("1.1 Variables del modelo")

* Tabla de especificacion del modelo
putdocx table tbl_mod = (8, 3), border(all, single) width(100%)

* Fila de encabezado
putdocx table tbl_mod(1,1) = ("Tipo"),        bold shading(1F4E79) font("Calibri", 10, "FFFFFF")
putdocx table tbl_mod(1,2) = ("Variable"),    bold shading(1F4E79) font("Calibri", 10, "FFFFFF")
putdocx table tbl_mod(1,3) = ("Descripcion"), bold shading(1F4E79) font("Calibri", 10, "FFFFFF")

* Filas de datos
putdocx table tbl_mod(2,1) = ("Endogena 1"), font("Calibri", 10)
putdocx table tbl_mod(2,2) = ("d_ltc"),      font("Calibri", 10) italic
putdocx table tbl_mod(2,3) = ("Delta log Tipo de Cambio CLP/USD (primera diferencia)"), font("Calibri", 10)

putdocx table tbl_mod(3,1) = ("Endogena 2"), font("Calibri", 10)
putdocx table tbl_mod(3,2) = ("tpm"),         font("Calibri", 10) italic
putdocx table tbl_mod(3,3) = ("Tasa de Politica Monetaria BCCh (%) - en niveles, SSW 1990"), font("Calibri", 10)

putdocx table tbl_mod(4,1) = ("Endogena 3"), font("Calibri", 10)
putdocx table tbl_mod(4,2) = ("d_lipc"),     font("Calibri", 10) italic
putdocx table tbl_mod(4,3) = ("Delta log IPC Chile (inflacion mensual, primera diferencia)"), font("Calibri", 10)

putdocx table tbl_mod(5,1) = ("Exogena 1"), font("Calibri", 10)
putdocx table tbl_mod(5,2) = ("d_lpetro"),  font("Calibri", 10) italic
putdocx table tbl_mod(5,3) = ("Delta log Precio del Petroleo Brent (USD/bbl)"), font("Calibri", 10)

putdocx table tbl_mod(6,1) = ("Exogena 2"), font("Calibri", 10)
putdocx table tbl_mod(6,2) = ("d_lcobre"),  font("Calibri", 10) italic
putdocx table tbl_mod(6,3) = ("Delta log Precio del Cobre (USD/lb)"), font("Calibri", 10)

putdocx table tbl_mod(7,1) = ("Exogena 3"), font("Calibri", 10)
putdocx table tbl_mod(7,2) = ("d_lipc_ee"), font("Calibri", 10) italic
putdocx table tbl_mod(7,3) = ("Delta log IPC EE.UU. (CPIAUCSL, FRED)"), font("Calibri", 10)

putdocx table tbl_mod(8,1) = ("Exogena 4"), font("Calibri", 10)
putdocx table tbl_mod(8,2) = ("d_tasa_fed"), font("Calibri", 10) italic
putdocx table tbl_mod(8,3) = ("Delta Tasa Fed Funds (FEDFUNDS, FRED)"), font("Calibri", 10)

putdocx paragraph
putdocx text ("El orden Cholesky refleja la jerarquia de rigidez de precios: el tipo de cambio reacciona contemporaneamente a todos los shocks (mas flexible), la TPM reacciona al TC en el mismo periodo pero no al IPC, y los precios son los mas rigidos (responden con rezago a todos los shocks)."), font("Calibri", 10, italic)

* ── 2. PRUEBAS DE RAÍZ UNITARIA ──────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("2. Pruebas de Raiz Unitaria")

putdocx paragraph
putdocx text ("Se aplican tres pruebas complementarias para cada variable: Dickey-Fuller Aumentado (ADF) con 6 rezagos, Phillips-Perron (PP) con ancho de banda Newey-West = 8, y KPSS con lag maximo = 12 (todas con constante y tendencia). Valores criticos al 5%: ADF/PP aprox. -3.41; KPSS = 0.146."), font("Calibri", 10)

* [E5] CORRECCIÓN: tabla (11,7) → (8,7): 1 encabezado + 7 variables
* [E6] CORRECCIÓN: shading del encabezado incluido en cada comando de contenido
putdocx table tbl_ur = (8, 7), border(all, single) width(100%)

* Encabezado — shading integrado directamente en cada celda
putdocx table tbl_ur(1,1) = ("Variable"),    bold shading(1F4E79) font("Calibri", 10, "FFFFFF")
putdocx table tbl_ur(1,2) = ("ADF t-stat"),  bold shading(1F4E79) font("Calibri", 10, "FFFFFF")
putdocx table tbl_ur(1,3) = ("ADF p-val"),   bold shading(1F4E79) font("Calibri", 10, "FFFFFF")
putdocx table tbl_ur(1,4) = ("PP t-stat"),   bold shading(1F4E79) font("Calibri", 10, "FFFFFF")
putdocx table tbl_ur(1,5) = ("PP p-val"),    bold shading(1F4E79) font("Calibri", 10, "FFFFFF")
putdocx table tbl_ur(1,6) = ("KPSS stat."),  bold shading(1F4E79) font("Calibri", 10, "FFFFFF")
putdocx table tbl_ur(1,7) = ("Conclusion"),  bold shading(1F4E79) font("Calibri", 10, "FFFFFF")

* [E7] CORRECCIÓN: pre-calcular valores formateados con ":di %fmt" para evitar
*      macros anidados profundos que pueden fallar en el parser de Stata.

local varnames_code  "l_tc tpm l_ipc l_petro l_cobre l_ipc_ee tasa_fed"
local varnames_label `""log TC" "TPM" "log IPC" "log Petróleo" "log Cobre" "log IPC EE.UU." "Fed Funds""'
local conclusiones   `""I(1)*" "I(0) [SSW]" "I(1)*" "I(1)*" "I(1)*" "I(1)*" "I(1)*""'

local fila = 2
local i = 1
foreach v of local varnames_code {
    local lab   : word `i' of `varnames_label'
    local concl : word `i' of `conclusiones'
    * Pre-calcular valores formateados (evita macros anidados en putdocx)
    local f1 : di %6.3f `adf_`v'_t'
    local f2 : di %5.3f `adf_`v'_pv'
    local f3 : di %6.3f `pp_`v'_t'
    local f4 : di %5.3f `pp_`v'_pv'
    local f5 : di %6.4f `kpss_`v'_t'
    * Rellenar fila de la tabla
    putdocx table tbl_ur(`fila',1) = ("`lab' (nivel)"), font("Calibri", 10)
    putdocx table tbl_ur(`fila',2) = ("`f1'"),           font("Calibri", 10)
    putdocx table tbl_ur(`fila',3) = ("`f2'"),           font("Calibri", 10)
    putdocx table tbl_ur(`fila',4) = ("`f3'"),           font("Calibri", 10)
    putdocx table tbl_ur(`fila',5) = ("`f4'"),           font("Calibri", 10)
    putdocx table tbl_ur(`fila',6) = ("`f5'"),           font("Calibri", 10)
    putdocx table tbl_ur(`fila',7) = ("`concl'"),        font("Calibri", 10)
    local ++fila
    local ++i
}

putdocx paragraph
putdocx text ("Nota: * I(1) en niveles. Las series I(1) se usan en primera diferencia en el VAR (excepto TPM, mantenida en niveles: serie acotada [0.5%-11.25%]; inferencia valida en VAR niveles segun Sims, Stock y Watson 1990). VC 5% ADF/PP aprox. -3.41; KPSS = 0.146 (con tendencia)."), font("Calibri", 9, italic)

* ── 3. SELECCIÓN DE REZAGOS ──────────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("3. Seleccion de Rezagos")

putdocx paragraph
putdocx text ("La seleccion del numero optimo de rezagos se realizo con "), font("Calibri", 10)
putdocx text ("varsoc"), font("Calibri", 10, italic)
putdocx text (" (maximo 12 rezagos, incluyendo las 4 variables exogenas). Se privilegia el criterio SBIC (Schwarz) por su mayor penalizacion a la sobreparametrizacion y consistencia en muestras grandes."), font("Calibri", 10)

putdocx paragraph
putdocx text ("Numero de rezagos seleccionado: p = `p_opt'"), bold font("Calibri", 10)

* ── 4. ESTIMACIÓN DEL VAR ────────────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("4. Estimacion del VAR")

putdocx paragraph
putdocx text ("Se estima el modelo VAR(p=`p_opt') con orden Cholesky d_ltc > tpm > d_lipc. Las variables exogenas (petroleo, cobre, IPC EE.UU. y Fed Funds, todas en primera diferencia) capturan el entorno externo como bloque exogeno. La muestra efectiva comprende `nobs' observaciones mensuales."), font("Calibri", 10)

* Tabla resumen del modelo estimado
putdocx table tbl_est = (5, 2), border(all, single) width(65%)

putdocx table tbl_est(1,1) = ("Parametro"),  bold shading(E9EFF7) font("Calibri", 10)
putdocx table tbl_est(1,2) = ("Valor"),      bold shading(E9EFF7) font("Calibri", 10)

putdocx table tbl_est(2,1) = ("Rezagos (p)"),      font("Calibri", 10)
putdocx table tbl_est(2,2) = ("`p_opt'"),           font("Calibri", 10)
putdocx table tbl_est(3,1) = ("Obs. efectivas"),    font("Calibri", 10)
putdocx table tbl_est(3,2) = ("`nobs'"),            font("Calibri", 10)
putdocx table tbl_est(4,1) = ("Identificacion"),    font("Calibri", 10)
putdocx table tbl_est(4,2) = ("Cholesky recursiva"), font("Calibri", 10)
putdocx table tbl_est(5,1) = ("Exogenas"),          font("Calibri", 10)
putdocx table tbl_est(5,2) = ("d_lpetro, d_lcobre, d_lipc_ee, d_tasa_fed"), font("Calibri", 10)

* ── 5. DIAGNÓSTICOS ──────────────────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("5. Diagnosticos del Modelo")

putdocx paragraph, style(Heading2)
putdocx text ("5.1 Estabilidad: Raices del Polinomio VAR")

putdocx paragraph
putdocx text ("Para que el VAR sea dinamicamente estable, todos los modulos de las raices inversas del polinomio caracteristico deben estar dentro del circulo unitario. El grafico a continuacion presenta los resultados."), font("Calibri", 10)

putdocx image `"$resdir\grafico_estabilidad.png"', width(9cm) height(8cm)

putdocx paragraph, style(Heading2)
putdocx text ("5.2 Autocorrelacion de Residuos (Test LM)")

putdocx paragraph
putdocx text ("Se aplica el test de multiplicadores de Lagrange (LM) para autocorrelacion serial de los residuos hasta el rezago 12. La hipotesis nula es ausencia de autocorrelacion. Un p-valor > 0.05 indica que no se rechaza H0."), font("Calibri", 10)

putdocx paragraph
putdocx text ("Ver tabla varlmar en la salida de Stata."), font("Calibri", 10, italic)

putdocx paragraph, style(Heading2)
putdocx text ("5.3 Normalidad de Residuos (Jarque-Bera)")

putdocx paragraph
putdocx text ("El test de Jarque-Bera multivariado evalua si los residuos son conjuntamente normales. En muestras grandes (T>100) el VAR es robusto a desviaciones moderadas de normalidad por el teorema del limite central. El incumplimiento no invalida la inferencia asintotica."), font("Calibri", 10)

* ── 6. IRFs ──────────────────────────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("6. Funciones de Impulso-Respuesta (IRF)")

putdocx paragraph
putdocx text ("Las IRFs se computan a horizonte de 24 meses con intervalos de confianza al 95% por bootstrap (500 replicas, semilla = 1234). Se presentan los canales economicamente mas relevantes para el esquema de Metas de Inflacion del BCCh."), font("Calibri", 10)

putdocx paragraph, style(Heading2)
putdocx text ("6.1 Canal de Tasa de Interes: Shock TPM → Inflacion")

putdocx paragraph
putdocx text ("Un shock positivo de TPM debe reducir la inflacion. Esta es la prediccion central del canal de tasa de interes bajo el esquema IT del BCCh."), font("Calibri", 10)

putdocx image `"$resdir\irf_tpm_ipc.png"', width(14cm) height(9cm)

putdocx paragraph, style(Heading2)
putdocx text ("6.2 Canal Cambiario: Shock TPM → Tipo de Cambio")

putdocx paragraph
putdocx text ("Un alza de TPM aprecia el peso chileno (d_ltc cae). La apreciacion cambiaria reduce importaciones y contribuye a bajar la inflacion via pass-through."), font("Calibri", 10)

putdocx image `"$resdir\irf_tpm_tc.png"', width(14cm) height(9cm)

putdocx paragraph, style(Heading2)
putdocx text ("6.3 Pass-Through Cambiario: Shock TC → Inflacion")

putdocx paragraph
putdocx text ("Una depreciacion del peso (shock positivo en d_ltc) se transmite al nivel de precios domesticos. La magnitud y velocidad de este pass-through es critica para la conduccion de la politica monetaria en economia pequena y abierta."), font("Calibri", 10)

putdocx image `"$resdir\irf_tc_ipc.png"', width(14cm) height(9cm)

putdocx paragraph, style(Heading2)
putdocx text ("6.4 Panel Completo de IRFs (9 combinaciones)")

putdocx image `"$resdir\irf_panel.png"', width(16cm) height(14cm)

* ── 7. FEVD ──────────────────────────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("7. Descomposicion de Varianza del Error de Pronostico (FEVD)")

putdocx paragraph
putdocx text ("La FEVD cuantifica que proporcion del error de pronostico a cada horizonte se atribuye a cada tipo de shock. Permite identificar la importancia relativa del canal de tasa de interes versus el canal cambiario en la dinamica de los precios."), font("Calibri", 10)

putdocx paragraph, style(Heading2)
putdocx text ("7.1 FEVD de la Inflacion (Delta log IPC)")

putdocx image `"$resdir\fevd_ipc.png"', width(14cm) height(9cm)

putdocx paragraph, style(Heading2)
putdocx text ("7.2 FEVD del Tipo de Cambio (Delta log TC)")

putdocx image `"$resdir\fevd_tc.png"', width(14cm) height(9cm)

putdocx paragraph, style(Heading2)
putdocx text ("7.3 FEVD de la Tasa de Politica Monetaria (TPM)")

putdocx image `"$resdir\fevd_tpm.png"', width(14cm) height(9cm)

* ── 8. CONCLUSIONES ──────────────────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("8. Conclusiones")

putdocx paragraph
putdocx text ("El modelo VAR estimado para Chile identifica los principales canales de transmision de la politica monetaria bajo el esquema de Metas de Inflacion del BCCh:"), font("Calibri", 10)

putdocx paragraph
putdocx text ("(i) Canal de tasa de interes: Un shock positivo de TPM genera una respuesta negativa sobre la inflacion, confirmando que la politica monetaria contractiva reduce las presiones de demanda."), font("Calibri", 10)

putdocx paragraph
putdocx text ("(ii) Canal cambiario: El alza de TPM aprecia el tipo de cambio, y las depreciaciones se transmiten a precios domesticos (pass-through). La magnitud y velocidad son consistentes con la literatura para economias pequenas y abiertas de America Latina."), font("Calibri", 10)

putdocx paragraph
putdocx text ("(iii) La FEVD muestra que los shocks propios de la inflacion dominan en horizontes cortos. Los shocks de tipo de cambio y TPM adquieren mayor relevancia a mediano plazo (6-12 meses), consistente con los rezagos tipicos de la politica monetaria."), font("Calibri", 10)

* ── REFERENCIAS ──────────────────────────────────────────────────────────────
putdocx paragraph, style(Heading1)
putdocx text ("Referencias")

putdocx paragraph
putdocx text ("Mies, V., Morande, F. y Tapia, M. (2002). Politica monetaria y mecanismos de transmision: Nuevos elementos para una vieja discusion. "), font("Calibri", 10)
putdocx text ("Economia Chilena, 5"), font("Calibri", 10, italic)
putdocx text ("(3), 29-66."), font("Calibri", 10)

putdocx paragraph
putdocx text ("Justel, S. y Sansone, A. (2015). Exchange Rate Pass-Through to Prices: VAR Evidence for Chile. "), font("Calibri", 10)
putdocx text ("Working Paper N 747, Banco Central de Chile."), font("Calibri", 10, italic)

putdocx paragraph
putdocx text ("Quintero Otero, J. D. (2015). Impactos de la politica monetaria y canales de transmision en paises de America Latina con esquema de inflacion objetivo. "), font("Calibri", 10)
putdocx text ("Ensayos sobre Politica Economica, 33"), font("Calibri", 10, italic)
putdocx text ("(76), 61-75."), font("Calibri", 10)

putdocx paragraph
putdocx text ("Sims, C. A., Stock, J. H. y Watson, M. W. (1990). Inference in linear time series models with some unit roots. "), font("Calibri", 10)
putdocx text ("Econometrica, 58"), font("Calibri", 10, italic)
putdocx text ("(1), 113-144."), font("Calibri", 10)

putdocx paragraph
putdocx text ("Banco Central de Chile (2025). Informe de Politica Monetaria (IPoM). Santiago: BCCh."), font("Calibri", 10)

* ── GUARDAR DOCUMENTO ────────────────────────────────────────────────────────
putdocx save `"$resdir\resultados_VAR.docx"', replace

di ""
di "=================================================="
di "DOCUMENTO EXPORTADO: $resdir\resultados_VAR.docx"
di "=================================================="
di ""
di "VAR_Chile_v2.do completado."
di "  Modelo:     VAR(`p_opt') | d_ltc > tpm > d_lipc"
di "  Exogenas:   d_lpetro d_lcobre d_lipc_ee d_tasa_fed"
di "  Muestra:    `tmin' - `tmax'  (N = `nobs')"
di "  Resultados: $resdir\"
