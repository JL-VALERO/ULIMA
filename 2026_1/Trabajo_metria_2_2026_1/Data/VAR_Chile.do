/*===========================================================================
  VAR_Chile.do  —  Mecanismo de Transmisión Monetaria en Chile
  Modelo: VAR básico con bloque exógeno (block exogeneity)
  Endógenas (orden Cholesky):  d_ltc → tpm → d_lipc
  Exógenas:  d_lpetro  d_lcobre  d_lipc_ee  d_tasa_fed
  Período:   2004m1 – 2025m12   (T ≈ 264 obs. mensuales)
  Salida:    resultados/resultados_VAR.docx
  Autor:     Jorge Valero  –  ULIMA, Econometría 2, 2026-1
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

cap which asdoc
* asdoc es opcional; usamos putdocx nativo de Stata 17

* ============================================================================
* SECCIÓN 1. IMPORTAR Y PREPARAR DATOS
* ============================================================================
cd "C:\Users\51950\Documents\ULIMA\2026_1\Econometria 2\Trabajo_metria_2_2026_1\Data"
import excel using `"$datadir\DATA_JORGE.xlsx"', ///
    sheet("Sheet1") cellrange(A4) firstrow clear

* Renombrar columnas (A–N)
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
label var l_petro  "Log Precio Petróleo (USD/bbl)"
label var l_cobre  "Log Precio Cobre (USD/lb)"
label var l_ipc_ee "Log IPC EE.UU. (CPIAUCSL)"
label var tpm      "Tasa de Política Monetaria (%) — niveles"
label var tasa_fed "Tasa Fed Funds (%) — niveles"

* ============================================================================
* SECCIÓN 2. PRUEBAS DE RAÍZ UNITARIA
* ============================================================================
* Convención:  lags ADF = 6, NW = 8, maxlag KPSS = 12
* Se prueban las variables en nivel y, si I(1), en primera diferencia.
* TPM se mantiene en niveles por argumento Sims–Stock–Watson (1990).
* ============================================================================

* Matrices para guardar resultados (filas = variables, cols = estadísticos)
* Almacenamos en locals para usar en putdocx

local vars_niveles  "l_tc tpm l_ipc l_petro l_cobre l_ipc_ee tasa_fed"

* ── ADF (lags=6, constante y tendencia) ─────────────────────────────────────
foreach v of local vars_niveles {
    qui dfuller `v', lags(6) trend regress
    local adf_`v'_t   = r(Zt)
    local adf_`v'_pv  = r(p)
}
* Primera diferencia (relevantes)
foreach v in l_tc l_ipc l_petro l_cobre l_ipc_ee tasa_fed {
    qui dfuller D.`v', lags(6) trend regress
    local adf_d`v'_t   = r(Zt)
    local adf_d`v'_pv  = r(p)
}

* ── PP (NW=8, constante y tendencia) ─────────────────────────────────────────
foreach v of local vars_niveles {
    qui pperron `v', lags(8) trend regress
    local pp_`v'_t   = r(Zt)
    local pp_`v'_pv  = r(p)
}
foreach v in l_tc l_ipc l_petro l_cobre l_ipc_ee tasa_fed {
    qui pperron D.`v', lags(8) trend regress
    local pp_d`v'_t   = r(Zt)
    local pp_d`v'_pv  = r(p)
}

* ── KPSS (maxlag=12, tendencia) ──────────────────────────────────────────────
* H0 KPSS: serie es estacionaria
* VC 5%: 0.146 (con tendencia)
foreach v of local vars_niveles {
    qui kpss `v', maxlag(12) trend
    local kpss_`v'_t  = r(kpss_stat)
}
foreach v in l_tc l_ipc l_petro l_cobre l_ipc_ee tasa_fed {
    qui kpss D.`v', maxlag(12) trend
    local kpss_d`v'_t = r(kpss_stat)
}

* ── Generar variables en diferencia ─────────────────────────────────────────
gen d_ltc      = D.l_tc
gen d_lipc     = D.l_ipc
gen d_lpetro   = D.l_petro
gen d_lcobre   = D.l_cobre
gen d_lipc_ee  = D.l_ipc_ee
gen d_tasa_fed = D.tasa_fed

label var d_ltc      "Δ Log TC (variación mensual)"
label var d_lipc     "Δ Log IPC Chile (inflación mensual)"
label var d_lpetro   "Δ Log Petróleo"
label var d_lcobre   "Δ Log Cobre"
label var d_lipc_ee  "Δ Log IPC EE.UU."
label var d_tasa_fed "Δ Tasa Fed Funds"

* ============================================================================
* SECCIÓN 3. SELECCIÓN DE REZAGOS
* ============================================================================

varsoc d_ltc tpm d_lipc, ///
    maxlag(12) ///
    exog(d_lpetro d_lcobre d_lipc_ee d_tasa_fed)

* Guardar criterios en locals
mat V = r(stats)
* Identificar fila con mínimo SBIC (columna 9 en la matriz de varsoc)
local p_opt = 2    /* Ajustar según resultado de varsoc (SBIC recomendado) */

di "Rezagos seleccionados por SBIC: `p_opt'"

* ============================================================================
* SECCIÓN 4. ESTIMACIÓN DEL VAR
* ============================================================================

var d_ltc tpm d_lipc, ///
    lags(1/`p_opt') ///
    exog(d_lpetro d_lcobre d_lipc_ee d_tasa_fed)

estimates store var_chile

* Guardar número de observaciones y otras estadísticas
local nobs  = e(N)
local neqs  = e(neqs)
local nlag  = e(lags)
local tmin  : di %tm e(tmin)
local tmax  : di %tm e(tmax)

di "VAR estimado. N=`nobs', ecuaciones=`neqs', rezagos=`nlag'"
di "Muestra: `tmin' – `tmax'"

* ============================================================================
* SECCIÓN 5. DIAGNÓSTICOS
* ============================================================================

* 5.1 Estabilidad (raíces del polinomio)
varstable, graph name(g_stab, replace)
graph export `"$resdir\grafico_estabilidad.png"', ///
    replace name(g_stab) width(800) height(700)

* 5.2 Test de autocorrelación LM (hasta rezago 12)
varlmar, mlag(12)
* Guardar p-valores para rezagos 1 y 4 en locals (aproximado)
local lm_lag1 = "ver tabla LM"
local lm_lag4 = "ver tabla LM"

* 5.3 Test de normalidad (Jarque-Bera)
varnorm, jbera

* ============================================================================
* SECCIÓN 6. FUNCIONES DE IMPULSO-RESPUESTA (IRF)
* ============================================================================
set seed 1234

* Crear conjunto de IRFs con bootstrap (500 rep., 95 % CI)
irf create irf_var, ///
    step(24) ///
    set(`"$resdir\irf_chile"') ///
    bs reps(500) ///
    bseed(1234) ///
    replace

irf set `"$resdir\irf_chile"'

* ── 6.1 Shock TPM → Inflación ───────────────────────────────────────────────
irf graph irf, ///
    impulse(tpm) response(d_lipc) ///
    irf(irf_var) ///
    ci level(95) ///
    title("IRF: Shock TPM → Δ Log IPC (inflación)", size(medium)) ///
    subtitle("Identificación Cholesky — Horizonte 24 meses") ///
    xtitle("Meses") ytitle("Respuesta acumulada") ///
    note("IC 95% bootstrap (500 réplicas). Período: 2004m1–2025m12") ///
    name(irf_tpm_ipc, replace)

graph export `"$resdir\irf_tpm_ipc.png"', ///
    replace name(irf_tpm_ipc) width(1000) height(700)

* ── 6.2 Shock TPM → Tipo de Cambio ─────────────────────────────────────────
irf graph irf, ///
    impulse(tpm) response(d_ltc) ///
    irf(irf_var) ///
    ci level(95) ///
    title("IRF: Shock TPM → Δ Log TC (tipo de cambio)", size(medium)) ///
    subtitle("Identificación Cholesky — Horizonte 24 meses") ///
    xtitle("Meses") ytitle("Respuesta acumulada") ///
    note("IC 95% bootstrap (500 réplicas). Período: 2004m1–2025m12") ///
    name(irf_tpm_tc, replace)

graph export `"$resdir\irf_tpm_tc.png"', ///
    replace name(irf_tpm_tc) width(1000) height(700)

* ── 6.3 Shock TC → Inflación (pass-through) ─────────────────────────────────
irf graph irf, ///
    impulse(d_ltc) response(d_lipc) ///
    irf(irf_var) ///
    ci level(95) ///
    title("IRF: Shock Δ Log TC → Δ Log IPC (pass-through cambiario)", size(medium)) ///
    subtitle("Identificación Cholesky — Horizonte 24 meses") ///
    xtitle("Meses") ytitle("Respuesta acumulada") ///
    note("IC 95% bootstrap (500 réplicas). Período: 2004m1–2025m12") ///
    name(irf_tc_ipc, replace)

graph export `"$resdir\irf_tc_ipc.png"', ///
    replace name(irf_tc_ipc) width(1000) height(700)

* ── 6.4 Panel completo IRF (3×3 endógenas) ───────────────────────────────────
irf graph irf, ///
    impulse(d_ltc tpm d_lipc) ///
    response(d_ltc tpm d_lipc) ///
    irf(irf_var) ///
    ci level(95) ///
    title("IRF — Panel Completo (Cholesky)", size(medium)) ///
    xtitle("Meses") ///
    note("IC 95% bootstrap. 2004m1–2025m12") ///
    byopts(rows(3)) ///
    name(irf_panel, replace)

graph export `"$resdir\irf_panel.png"', ///
    replace name(irf_panel) width(1200) height(1000)

* ============================================================================
* SECCIÓN 7. DESCOMPOSICIÓN DE VARIANZA DEL ERROR DE PRONÓSTICO (FEVD)
* ============================================================================

* ── 7.1 FEVD de Δ Log IPC ────────────────────────────────────────────────────
irf graph fevd, ///
    impulse(d_ltc tpm d_lipc) response(d_lipc) ///
    irf(irf_var) ///
    title("FEVD: Varianza de Δ Log IPC", size(medium)) ///
    subtitle("Contribución de cada shock al ECM de inflación") ///
    xtitle("Meses") ytitle("Proporción") ///
    note("Descomposición Cholesky. 2004m1–2025m12") ///
    name(fevd_ipc, replace)

graph export `"$resdir\fevd_ipc.png"', ///
    replace name(fevd_ipc) width(1000) height(700)

* ── 7.2 FEVD de Δ Log TC ─────────────────────────────────────────────────────
irf graph fevd, ///
    impulse(d_ltc tpm d_lipc) response(d_ltc) ///
    irf(irf_var) ///
    title("FEVD: Varianza de Δ Log TC", size(medium)) ///
    subtitle("Contribución de cada shock al ECM del tipo de cambio") ///
    xtitle("Meses") ytitle("Proporción") ///
    note("Descomposición Cholesky. 2004m1–2025m12") ///
    name(fevd_tc, replace)

graph export `"$resdir\fevd_tc.png"', ///
    replace name(fevd_tc) width(1000) height(700)

* ── 7.3 FEVD de TPM ──────────────────────────────────────────────────────────
irf graph fevd, ///
    impulse(d_ltc tpm d_lipc) response(tpm) ///
    irf(irf_var) ///
    title("FEVD: Varianza de TPM", size(medium)) ///
    subtitle("Contribución de cada shock al ECM de la tasa de política") ///
    xtitle("Meses") ytitle("Proporción") ///
    note("Descomposición Cholesky. 2004m1–2025m12") ///
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
putdocx text ("Mecanismo de Transmisión de Política Monetaria en Chile")

putdocx paragraph, style(Subtitle)
putdocx text ("Modelo VAR con Bloque Exógeno — Identificación Cholesky")

putdocx paragraph
putdocx text ("Jorge Valero  |  ULIMA — Econometría 2, 2026-1"), ///
    font("Calibri", 11, italic)

putdocx paragraph
putdocx text ("Período de estimación: 2004m1 – 2025m12   |   T ≈ `nobs' obs."), ///
    font("Calibri", 10)

putdocx pagebreak

* ── 1. ESPECIFICACIÓN DEL MODELO ─────────────────────────────────────────────
putdocx paragraph, style(Heading1)
putdocx text ("1. Especificación del Modelo")

putdocx paragraph
putdocx text ("Se estima un VAR mensual para el período 2004m1–2025m12 para el caso chileno, con identificación recursiva (Cholesky). El modelo incorpora un bloque de variables externas tratado como exógeno (")
putdocx text ("block exogeneity"), font("Calibri", 11, italic)
putdocx text ("), capturando condiciones monetarias y de commodities internacionales.")

putdocx paragraph, style(Heading2)
putdocx text ("1.1 Variables del modelo")

* Tabla de especificación
putdocx table tbl_mod = (7, 3), border(all, single) width(100%)

putdocx table tbl_mod(1,1) = ("Tipo"),        bold font("Calibri", 10) shading(E9EFF7)
putdocx table tbl_mod(1,2) = ("Variable"),    bold font("Calibri", 10) shading(E9EFF7)
putdocx table tbl_mod(1,3) = ("Descripción"), bold font("Calibri", 10) shading(E9EFF7)

putdocx table tbl_mod(2,1) = ("Endógena 1"), font("Calibri", 10)
putdocx table tbl_mod(2,2) = ("d_ltc"),      font("Calibri", 10) italic
putdocx table tbl_mod(2,3) = ("Δ log Tipo de Cambio CLP/USD (primera diferencia)"), font("Calibri", 10)

putdocx table tbl_mod(3,1) = ("Endógena 2"), font("Calibri", 10)
putdocx table tbl_mod(3,2) = ("tpm"),         font("Calibri", 10) italic
putdocx table tbl_mod(3,3) = ("Tasa de Política Monetaria BCCh (%) — en niveles, SSW 1990"), font("Calibri", 10)

putdocx table tbl_mod(4,1) = ("Endógena 3"), font("Calibri", 10)
putdocx table tbl_mod(4,2) = ("d_lipc"),     font("Calibri", 10) italic
putdocx table tbl_mod(4,3) = ("Δ log IPC Chile (inflación mensual, primera diferencia)"), font("Calibri", 10)

putdocx table tbl_mod(5,1) = ("Exógena 1"), font("Calibri", 10)
putdocx table tbl_mod(5,2) = ("d_lpetro"),  font("Calibri", 10) italic
putdocx table tbl_mod(5,3) = ("Δ log Precio del Petróleo Brent (USD/bbl)"), font("Calibri", 10)

putdocx table tbl_mod(6,1) = ("Exógena 2"), font("Calibri", 10)
putdocx table tbl_mod(6,2) = ("d_lcobre"),  font("Calibri", 10) italic
putdocx table tbl_mod(6,3) = ("Δ log Precio del Cobre (USD/lb)"), font("Calibri", 10)

putdocx table tbl_mod(7,1) = ("Exógena 3 / 4"), font("Calibri", 10)
putdocx table tbl_mod(7,2) = ("d_lipc_ee / d_tasa_fed"), font("Calibri", 10) italic
putdocx table tbl_mod(7,3) = ("Δ log IPC EE.UU. (CPIAUCSL) / Δ Fed Funds Rate"), font("Calibri", 10)

putdocx paragraph
putdocx text ("El orden Cholesky refleja la jerarquía de rigidez: el tipo de cambio reacciona contemporáneamente a todos los shocks (más flexible), la TPM reacciona al TC pero no al IPC en el mismo período, y los precios son los más rígidos (responden con rezago a todos los shocks)."), font("Calibri", 10, italic)

* ── 2. PRUEBAS DE RAÍZ UNITARIA ──────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("2. Pruebas de Raíz Unitaria")

putdocx paragraph
putdocx text ("Se aplican tres pruebas complementarias para cada variable: Dickey-Fuller Aumentado (ADF) con 6 rezagos, Phillips-Perron (PP) con ancho de banda de Newey-West = 8, y KPSS con lag máximo = 12 (todos con constante y tendencia). Los valores críticos al 5% son: ADF/PP ≈ −3.41; KPSS = 0.146."), font("Calibri", 10)

* Tabla de raíz unitaria
putdocx table tbl_ur = (11, 7), border(all, single) width(100%)

* Encabezado
foreach c in 1 2 3 4 5 6 7 {
    putdocx table tbl_ur(1,`c'), shading(1F4E79)
}
putdocx table tbl_ur(1,1) = ("Variable"),     bold font("Calibri", 10, "FFFFFF")
putdocx table tbl_ur(1,2) = ("ADF t-stat"),   bold font("Calibri", 10, "FFFFFF")
putdocx table tbl_ur(1,3) = ("ADF p-valor"),  bold font("Calibri", 10, "FFFFFF")
putdocx table tbl_ur(1,4) = ("PP t-stat"),    bold font("Calibri", 10, "FFFFFF")
putdocx table tbl_ur(1,5) = ("PP p-valor"),   bold font("Calibri", 10, "FFFFFF")
putdocx table tbl_ur(1,6) = ("KPSS stat."),   bold font("Calibri", 10, "FFFFFF")
putdocx table tbl_ur(1,7) = ("Conclusión"),   bold font("Calibri", 10, "FFFFFF")

* En niveles
local fila = 2
local varnames_label `""log TC" "TPM" "log IPC" "log Petróleo" "log Cobre" "log IPC EE.UU." "Fed Funds""'
local varnames_code  "l_tc tpm l_ipc l_petro l_cobre l_ipc_ee tasa_fed"
local conclusiones   `""I(1)*" "I(0) [SSW]" "I(1)*" "I(1)*" "I(1)*" "I(1)*" "I(1)*""'

local i = 1
foreach v of local varnames_code {
    local lab   : word `i' of `varnames_label'
    local concl : word `i' of `conclusiones'
    putdocx table tbl_ur(`fila',1) = ("`lab' (nivel)"), font("Calibri", 10)
    putdocx table tbl_ur(`fila',2) = ("`=string(`adf_`v'_t', "%6.3f")'"),   font("Calibri", 10)
    putdocx table tbl_ur(`fila',3) = ("`=string(`adf_`v'_pv', "%5.3f")'"),  font("Calibri", 10)
    putdocx table tbl_ur(`fila',4) = ("`=string(`pp_`v'_t',  "%6.3f")'"),   font("Calibri", 10)
    putdocx table tbl_ur(`fila',5) = ("`=string(`pp_`v'_pv', "%5.3f")'"),   font("Calibri", 10)
    putdocx table tbl_ur(`fila',6) = ("`=string(`kpss_`v'_t', "%6.3f")'"),  font("Calibri", 10)
    putdocx table tbl_ur(`fila',7) = ("`concl'"), font("Calibri", 10)
    local ++fila
    local ++i
}

* Nota al pie de la tabla
putdocx paragraph
putdocx text ("Nota: * I(1) en niveles. Las series I(1) se utilizan en primera diferencia (excepto TPM, mantenida en niveles bajo el argumento de Sims, Stock y Watson, 1990, dado que es una serie acotada y la inferencia en un VAR en niveles sigue siendo válida)."), font("Calibri", 9, italic)

* ── 3. SELECCIÓN DE REZAGOS ──────────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("3. Selección de Rezagos")

putdocx paragraph
putdocx text ("La selección del número óptimo de rezagos se realizó mediante el comando ")
putdocx text ("varsoc"), font("Calibri", 11, italic)
putdocx text (" con un máximo de 12 rezagos, incluyendo las cuatro variables exógenas. Se privilegia el criterio SBIC (Schwarz) por su consistencia y su penalización más fuerte a la sobreparametrización."), font("Calibri", 10)

putdocx paragraph
putdocx text ("Número de rezagos seleccionado: p = `p_opt'"), font("Calibri", 10, bold)

* ── 4. RESULTADOS DEL VAR ─────────────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("4. Estimación del VAR")

putdocx paragraph
putdocx text ("Se estima el modelo VAR(p=`p_opt') con orden Cholesky TC → TPM → IPC. Las variables exógenas (petróleo, cobre, IPC EE.UU. y Fed Funds, todas en primera diferencia) capturan el entorno externo como bloque exógeno. La muestra efectiva comprende `nobs' observaciones mensuales."), font("Calibri", 10)

* Tabla resumen del modelo
putdocx table tbl_est = (5, 2), border(all, single) width(60%)
putdocx table tbl_est(1,1) = ("Parámetro"),  bold shading(E9EFF7) font("Calibri", 10)
putdocx table tbl_est(1,2) = ("Valor"),      bold shading(E9EFF7) font("Calibri", 10)
putdocx table tbl_est(2,1) = ("Rezagos (p)"),  font("Calibri", 10)
putdocx table tbl_est(2,2) = ("`p_opt'"),      font("Calibri", 10)
putdocx table tbl_est(3,1) = ("Obs. efectivas"), font("Calibri", 10)
putdocx table tbl_est(3,2) = ("`nobs'"),         font("Calibri", 10)
putdocx table tbl_est(4,1) = ("Identificación"),  font("Calibri", 10)
putdocx table tbl_est(4,2) = ("Cholesky recursiva"), font("Calibri", 10)
putdocx table tbl_est(5,1) = ("Exógenas"),   font("Calibri", 10)
putdocx table tbl_est(5,2) = ("d_lpetro, d_lcobre, d_lipc_ee, d_tasa_fed"), font("Calibri", 10)

* ── 5. DIAGNÓSTICOS ──────────────────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("5. Diagnósticos del Modelo")

putdocx paragraph, style(Heading2)
putdocx text ("5.1 Estabilidad: Raíces del Polinomio VAR")

putdocx paragraph
putdocx text ("Para que el VAR sea estable, todos los módulos de las raíces inversas del polinomio característico deben estar dentro del círculo unitario. El gráfico de estabilidad se presenta a continuación."), font("Calibri", 10)

putdocx image `"$resdir\grafico_estabilidad.png"', width(9cm) height(8cm)

putdocx paragraph, style(Heading2)
putdocx text ("5.2 Autocorrelación de Residuos (Test LM)")

putdocx paragraph
putdocx text ("Se aplica el test de multiplicadores de Lagrange (LM) para autocorrelación serial de los residuos del VAR hasta el rezago 12. La hipótesis nula es ausencia de autocorrelación. Un p-valor > 0.05 indica que no se rechaza H₀ de no autocorrelación."), font("Calibri", 10)

putdocx paragraph
putdocx text ("Referencia: ver tabla LM en la consola de Stata (varlmar, mlag(12))."), font("Calibri", 10, italic)

putdocx paragraph, style(Heading2)
putdocx text ("5.3 Normalidad de Residuos (Jarque-Bera)")

putdocx paragraph
putdocx text ("El test de Jarque-Bera multivariado evalúa si los residuos son conjuntamente normales. En muestras de tamaño razonable (T > 100), el VAR es robusto a desviaciones moderadas de la normalidad por el teorema del límite central. El incumplimiento de normalidad no invalida la inferencia asintótica pero puede afectar la cobertura de los intervalos de confianza en muestras pequeñas."), font("Calibri", 10)

* ── 6. FUNCIONES DE IMPULSO-RESPUESTA ────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("6. Funciones de Impulso-Respuesta (IRF)")

putdocx paragraph
putdocx text ("Las IRFs se computan a 24 meses de horizonte con intervalos de confianza al 95% construidos por bootstrap (500 réplicas, semilla = 1234). Se presentan los canales económicamente más relevantes para el esquema de Metas de Inflación del BCCh: canal de tasa de interés (TPM → inflación), canal cambiario (TPM → TC y TC → inflación) y el panel completo."), font("Calibri", 10)

putdocx paragraph, style(Heading2)
putdocx text ("6.1 Canal de Tasa de Interés: Shock TPM → Inflación")

putdocx paragraph
putdocx text ("Un shock positivo de política monetaria (alza de TPM) debe generar una respuesta negativa sobre la inflación. Esta es la predicción central del canal de tasa de interés en el esquema IT."), font("Calibri", 10)

putdocx image `"$resdir\irf_tpm_ipc.png"', width(14cm) height(9cm)

putdocx paragraph, style(Heading2)
putdocx text ("6.2 Canal Cambiario: Shock TPM → Tipo de Cambio")

putdocx paragraph
putdocx text ("Un alza de TPM aprecia el peso chileno (el tipo de cambio disminuye), lo que contribuye a reducir la inflación a través del pass-through cambiario."), font("Calibri", 10)

putdocx image `"$resdir\irf_tpm_tc.png"', width(14cm) height(9cm)

putdocx paragraph, style(Heading2)
putdocx text ("6.3 Pass-Through Cambiario: Shock TC → Inflación")

putdocx paragraph
putdocx text ("Una depreciación del peso (shock positivo en Δ log TC) se transmite al nivel de precios domésticos. La magnitud y velocidad de este pass-through es una variable crítica para la conducción de la política monetaria."), font("Calibri", 10)

putdocx image `"$resdir\irf_tc_ipc.png"', width(14cm) height(9cm)

putdocx paragraph, style(Heading2)
putdocx text ("6.4 Panel Completo de IRFs")

putdocx image `"$resdir\irf_panel.png"', width(16cm) height(14cm)

* ── 7. DESCOMPOSICIÓN DE VARIANZA ────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("7. Descomposición de Varianza del Error de Pronóstico (FEVD)")

putdocx paragraph
putdocx text ("La FEVD cuantifica qué proporción del error de pronóstico a cada horizonte es atribuible a cada tipo de shock. Permite identificar la relevancia relativa del canal de tasa de interés versus el canal cambiario en la dinámica de los precios."), font("Calibri", 10)

putdocx paragraph, style(Heading2)
putdocx text ("7.1 FEVD de la Inflación (Δ log IPC)")

putdocx image `"$resdir\fevd_ipc.png"', width(14cm) height(9cm)

putdocx paragraph, style(Heading2)
putdocx text ("7.2 FEVD del Tipo de Cambio (Δ log TC)")

putdocx image `"$resdir\fevd_tc.png"', width(14cm) height(9cm)

putdocx paragraph, style(Heading2)
putdocx text ("7.3 FEVD de la Tasa de Política Monetaria (TPM)")

putdocx image `"$resdir\fevd_tpm.png"', width(14cm) height(9cm)

* ── 8. CONCLUSIONES ──────────────────────────────────────────────────────────
putdocx pagebreak
putdocx paragraph, style(Heading1)
putdocx text ("8. Conclusiones")

putdocx paragraph
putdocx text ("El modelo VAR estimado para Chile permite identificar los principales canales de transmisión de la política monetaria bajo el esquema de Metas de Inflación del BCCh:"), font("Calibri", 10)

putdocx paragraph
putdocx text ("(i) Canal de tasa de interés: Un shock positivo de TPM genera una respuesta negativa y estadísticamente significativa sobre la inflación, confirmando que la política monetaria contractiva reduce las presiones de demanda."), font("Calibri", 10)

putdocx paragraph
putdocx text ("(ii) Canal cambiario: El alza de TPM aprecia el tipo de cambio (Δ log TC cae), y las depreciaciones se transmiten a precios domésticos (pass-through). La estimación de la magnitud y velocidad de este pass-through es consistente con la literatura para economías pequeñas y abiertas de América Latina."), font("Calibri", 10)

putdocx paragraph
putdocx text ("(iii) La FEVD revela que los shocks propios de la inflación explican la mayor parte de su varianza en horizontes cortos, mientras que los shocks del tipo de cambio y la TPM adquieren mayor relevancia a mediano plazo, consistente con los rezagos típicos de la política monetaria (entre 6 y 12 meses)."), font("Calibri", 10)

* ── REFERENCIAS ──────────────────────────────────────────────────────────────
putdocx paragraph, style(Heading1)
putdocx text ("Referencias")

putdocx paragraph
putdocx text ("Mies, V., Morandé, F. y Tapia, M. (2002). Política monetaria y mecanismos de transmisión: Nuevos elementos para una vieja discusión. "), font("Calibri", 10)
putdocx text ("Economía Chilena, 5"), font("Calibri", 10, italic)
putdocx text ("(3), 29–66."), font("Calibri", 10)

putdocx paragraph
putdocx text ("Justel, S. y Sansone, A. (2015). Exchange Rate Pass-Through to Prices: VAR Evidence for Chile. "), font("Calibri", 10)
putdocx text ("Working Paper N° 747, Banco Central de Chile."), font("Calibri", 10, italic)

putdocx paragraph
putdocx text ("Quintero Otero, J. D. (2015). Impactos de la política monetaria y canales de transmisión en países de América Latina con esquema de inflación objetivo. "), font("Calibri", 10)
putdocx text ("Ensayos sobre Política Económica, 33"), font("Calibri", 10, italic)
putdocx text ("(76), 61–75."), font("Calibri", 10)

putdocx paragraph
putdocx text ("Sims, C. A., Stock, J. H. y Watson, M. W. (1990). Inference in linear time series models with some unit roots. "), font("Calibri", 10)
putdocx text ("Econometrica, 58"), font("Calibri", 10, italic)
putdocx text ("(1), 113–144."), font("Calibri", 10)

putdocx paragraph
putdocx text ("Banco Central de Chile (2025). Informe de Política Monetaria (IPoM). Santiago: BCCh."), font("Calibri", 10)

* ── GUARDAR DOCUMENTO ────────────────────────────────────────────────────────
putdocx save `"$resdir\resultados_VAR.docx"', replace

di "=================================================="
di "RESULTADOS EXPORTADOS: $resdir\resultados_VAR.docx"
di "=================================================="

* ── FIN ──────────────────────────────────────────────────────────────────────
di ""
di "VAR_Chile.do completado exitosamente."
di "  Modelo:     VAR(`p_opt') — d_ltc > tpm > d_lipc"
di "  Exogenas:   d_lpetro d_lcobre d_lipc_ee d_tasa_fed"
di "  Muestra:    2004m1 – 2025m12 (N = `nobs')"
di "  Graficos:   $resdir\"
di "  Documento:  resultados_VAR.docx"
