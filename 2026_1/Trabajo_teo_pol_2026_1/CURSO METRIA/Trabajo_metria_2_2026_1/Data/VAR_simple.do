/*===========================================================================
  VAR_simple.do  —  Transmision Monetaria en Chile (bloque domestico)
  VAR mensual 2004m1-2025m12  |  Stata 17
  Cholesky: d_ltc → tpm → d_lipc  (TC mas flexible, IPC mas rigido)
  Exogenas: d_lpetro  d_lcobre  d_lipc_ee  d_tasa_fed
  IRF ortogonalizadas (oirf): bandas IC 90% y 95% simultaneas
  Autor: Jorge Valero  |  ULIMA Econometria 2, 2026-1
===========================================================================*/

version 17
clear all
set more off

* Rutas
global datadir `"C:\Users\51950\Documents\ULIMA\2026_1\Econometria 2\Trabajo_metria_2_2026_1\Data"'
global resdir  `"C:\Users\51950\Documents\ULIMA\2026_1\Econometria 2\Trabajo_metria_2_2026_1\resultados"'
cap mkdir "$resdir"


*=============================================================================
* SECCIÓN 1 — DEFINICIÓN DE VARIABLES
*=============================================================================

* Fila 4 = primer dato (2004-01-01). Sin "firstrow": esa fila es dato, no encabezado.
import excel using `"$datadir\DATA_JORGE.xlsx"', sheet("Sheet1") cellrange(A4) clear

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

* Estructura temporal mensual (2004m1 en adelante)
drop if missing(fecha)
gen fecha_m = mofd(fecha)
format fecha_m %tm
tsset fecha_m, monthly
keep if fecha_m >= tm(2004m1)

* Transformaciones logaritmicas
gen l_tc     = log(tc)
gen l_ipc    = log(ipc)
gen l_petro  = log(petro)
gen l_cobre  = log(cobre)
gen l_ipc_ee = log(ipc_eeuu)

* Etiquetas
label var tc        "Tipo de Cambio nominal CLP/USD"
label var tpm       "Tasa de Politica Monetaria BCCh (%)"
label var ipc       "IPC Chile (base 2009 = 100)"
label var petro     "Precio Petroleo Brent (USD/barril)"
label var cobre     "Precio Cobre (USD/libra)"
label var tasa_fed  "Fed Funds Rate EE.UU. (%)"
label var ipc_eeuu  "IPC EE.UU. CPIAUCSL (base 1982-84 = 100)"
label var l_tc      "Log Tipo de Cambio CLP/USD"
label var l_ipc     "Log IPC Chile"
label var l_petro   "Log Precio Petroleo"
label var l_cobre   "Log Precio Cobre"
label var l_ipc_ee  "Log IPC EE.UU."

di _newline "=== ESTADISTICAS DESCRIPTIVAS ==="
summarize tc tpm ipc petro cobre tasa_fed ipc_eeuu


*=============================================================================
* SECCIÓN 2 — PRUEBAS DE ESTACIONARIEDAD
*=============================================================================
* H0 ADF/PP: la serie TIENE raiz unitaria (no estacionaria).
* Rechazar H0: |t| > 3.41 o p < 0.05 → serie I(0).
* Todas las pruebas usan constante y tendencia.

* --- ADF en niveles (lags = 4) -----------------------------------------------
di _newline "=== ADF EN NIVELES — H0: raiz unitaria ==="
dfuller l_tc,     lags(4) trend   // TC: se espera I(1)
dfuller tpm,      lags(4) trend   // TPM: posiblemente I(0), serie acotada
dfuller l_ipc,    lags(4) trend   // IPC: se espera I(1)
dfuller l_petro,  lags(4) trend   // Petroleo: se espera I(1)
dfuller l_cobre,  lags(4) trend   // Cobre: se espera I(1)
dfuller l_ipc_ee, lags(4) trend   // IPC EE.UU.: se espera I(1)
dfuller tasa_fed, lags(4) trend   // Fed Funds: se espera I(1)

* --- Phillips-Perron en niveles (NW bandwidth = 8) ---------------------------
di _newline "=== PHILLIPS-PERRON EN NIVELES ==="
pperron l_tc,     lags(8) trend
pperron tpm,      lags(8) trend
pperron l_ipc,    lags(8) trend
pperron l_petro,  lags(8) trend
pperron l_cobre,  lags(8) trend
pperron l_ipc_ee, lags(8) trend
pperron tasa_fed, lags(8) trend

* --- Primera diferencia: variables I(1) -------------------------------------
* TPM se mantiene en niveles (SSW 1990): serie acotada [0.5%-11.25%],
* la inferencia en VAR en niveles es valida segun Sims, Stock y Watson (1990).
gen d_ltc      = D.l_tc
gen d_lipc     = D.l_ipc
gen d_lpetro   = D.l_petro
gen d_lcobre   = D.l_cobre
gen d_lipc_ee  = D.l_ipc_ee
gen d_tasa_fed = D.tasa_fed

label var d_ltc      "D.log TC: variacion mensual tipo de cambio"
label var d_lipc     "D.log IPC Chile: inflacion mensual"
label var d_lpetro   "D.log Precio Petroleo"
label var d_lcobre   "D.log Precio Cobre"
label var d_lipc_ee  "D.log IPC EE.UU."
label var d_tasa_fed "D.Tasa Fed Funds"

* --- ADF en primera diferencia (confirmar I(0)) ------------------------------
di _newline "=== ADF EN PRIMERA DIFERENCIA — confirmar estacionariedad ==="
dfuller d_ltc,      lags(4) trend
dfuller d_lipc,     lags(4) trend
dfuller d_lpetro,   lags(4) trend
dfuller d_lcobre,   lags(4) trend
dfuller d_lipc_ee,  lags(4) trend
dfuller d_tasa_fed, lags(4) trend

* --- Resumen de decision de estacionariedad ---------------------------------
di _newline "=== RESUMEN: FORMA DE CADA VARIABLE EN EL VAR ==="
di " Variable       | Orden | Rol en el VAR"
di " ---------------|-------|------------------------------------"
di " log TC         | I(1)  | d_ltc      endogena (1ra posicion)"
di " TPM            | I(0)* | tpm        endogena (2da posicion)"
di " log IPC Chile  | I(1)  | d_lipc     endogena (3ra posicion)"
di " log Petroleo   | I(1)  | d_lpetro   exogena"
di " log Cobre      | I(1)  | d_lcobre   exogena"
di " log IPC EE.UU. | I(1)  | d_lipc_ee  exogena"
di " Tasa Fed       | I(1)  | d_tasa_fed exogena"
di " * TPM en niveles: SSW (1990)"


*=============================================================================
* SECCIÓN 3 — ESTIMACIÓN DEL VAR
*=============================================================================
* Ordenamiento Cholesky: d_ltc → tpm → d_lipc
*   d_ltc  (1ro): TC reacciona contemporaneamente a todos los shocks
*   tpm    (2do): BCCh reacciona al TC en t, pero no al IPC en t
*   d_lipc (3ro): IPC es el mas rigido (ajusta con mayor rezago)
*
* ─── SELECCION DE REZAGOS ───────────────────────────────────────────────────
* Criterio principal: SBIC (Schwarz). Criterios secundarios: AIC, HQIC.
* Si las IRF no tienen el signo teorico correcto con p = 2, prueba otro valor.
*   Signos esperados:
*     TPM(+) → d_lipc: NEGATIVO  (politica contractiva reduce inflacion)
*     TPM(+) → d_ltc : NEGATIVO  (suba de tasa aprecia el peso: TC baja)
*     d_ltc(+) → d_lipc: POSITIVO (depreciacion genera inflacion importada)

di _newline "=== SELECCION DE REZAGOS (SBIC = columna BIC) ==="
varsoc d_ltc tpm d_lipc, maxlag(12) exog(d_lpetro d_lcobre d_lipc_ee d_tasa_fed)

* AJUSTAR p segun el SBIC minimo de la tabla de arriba.
* Opciones a probar si los signos no son correctos: 1  2  3  4  6
local p = 2

* --- Estimacion del VAR ------------------------------------------------------
di _newline "=== VAR(`p') | Cholesky: d_ltc > tpm > d_lipc ==="
var d_ltc tpm d_lipc, lags(1/`p') exog(d_lpetro d_lcobre d_lipc_ee d_tasa_fed)
estimates store var_chile

local nobs = e(N)
local tmin : di %tm e(tmin)
local tmax : di %tm e(tmax)
di "  Muestra: `tmin' - `tmax'  |  N = `nobs' observaciones"

* --- Test de autocorrelacion (H0: sin autocorrelacion) -----------------------
di _newline "=== TEST LM — H0: residuos sin autocorrelacion ==="
di "(p-valor > 0.05 en todos los rezagos indica no autocorrelacion)"
varlmar, mlag(12)

* --- Test de estabilidad (raices dentro del circulo unitario) ----------------
di _newline "=== ESTABILIDAD DEL VAR ==="
di "(todas las raices inversas deben estar DENTRO del circulo unitario)"
varstable, graph
graph export `"$resdir\estabilidad_VAR.png"', replace width(700) height(650)
di "  -> estabilidad_VAR.png"


*=============================================================================
* SECCIÓN 4 — IMPULSO RESPUESTA (IRF)
*=============================================================================
* Se usan IRF ORTOGONALIZADAS (oirf): respuesta a un shock de 1 desvio
* estandar del shock estructural, identificado via Cholesky.
* Usar oirf (NO irf) es obligatorio para interpretar las IRF con la
* identificacion de Cholesky.
*
* Bandas de confianza: 90% (banda mas oscura) y 95% (banda mas clara).
* Se derivan del error estandar asintotico almacenado en el archivo .irf.
* Linea de cero: referencia para evaluar significancia estadistica.
*
* Si alguna IRF no tiene el signo teorico:
*   (1) ajusta p en Seccion 3 y re-ejecuta el script completo
*   (2) verifica que tpm, d_ltc, d_lipc esten bien transformadas
*   (3) verifica el orden Cholesky (d_ltc primero, d_lipc ultimo)

* --- Crear archivo IRF (IC asintóticos, horizonte 24 meses) -----------------
irf create irf_var, step(24) set(`"$resdir\irf_var"') replace

* El archivo .irf es un dataset de Stata. Variables clave:
*   irfname  impulse  response  step  oirf  upper_oirf  lower_oirf
* Si el script falla en la carga, ejecuta primero:
*   describe using "$resdir\irf_var.irf"

* =============================================================================
* IRF 1: Shock TPM → Inflacion (D.log IPC)
* SIGNO ESPERADO: NEGATIVO — politica contractiva reduce la inflacion
* =============================================================================
preserve
use `"$resdir\irf_var.irf"', clear
keep if irfname == "irf_var" & impulse == "tpm" & response == "d_lipc"
sort step

* Extraer IC 95% (almacenado) y calcular IC 90% desde el error estandar
cap gen u95 = upper_oirf
if _rc != 0 { gen u95 = upper }
cap gen l95 = lower_oirf
if _rc != 0 { gen l95 = lower }
gen se  = (u95 - oirf) / invnormal(0.975)   // SE desde IC 95% → z = 1.96
gen u90 = oirf + invnormal(0.95) * se        // IC 90% superior → z = 1.645
gen l90 = oirf - invnormal(0.95) * se        // IC 90% inferior
gen zero = 0

twoway ///
    (rarea u95 l95 step, fcolor(navy%10) lwidth(none))        ///  banda 95%
    (rarea u90 l90 step, fcolor(navy%25) lwidth(none))        ///  banda 90%
    (line  oirf step,    lcolor(navy) lwidth(medthick))       ///  OIRF
    (line  zero step,    lcolor(black) lpattern(dash) lwidth(thin)), /// y=0
    xlabel(0(4)24, labsize(small)) ///
    xtitle("Meses", size(small)) ///
    ytitle("Respuesta (OIRF)", size(small)) ///
    title("Shock TPM (+1 DE) → Inflacion [D.log IPC]", size(medsmall)) ///
    subtitle("Signo esperado: NEGATIVO (politica contractiva baja inflacion)", ///
             size(small) color(maroon)) ///
    note("IC 90% (banda oscura) y IC 95% (banda clara). Cholesky. `tmin'-`tmax'.", ///
         size(vsmall)) ///
    legend(order(2 "IC 90%" 1 "IC 95%" 3 "OIRF" 4 "y = 0") ///
           rows(1) size(vsmall)) ///
    name(irf_tpm_ipc, replace)
restore
graph export `"$resdir\irf_tpm_ipc.png"', name(irf_tpm_ipc) replace width(950) height(650)
di "  -> irf_tpm_ipc.png  (signo esperado: NEGATIVO)"

* =============================================================================
* IRF 2: Shock TPM → Tipo de Cambio (D.log TC)
* SIGNO ESPERADO: NEGATIVO — suba de tasa atrae capital y aprecia el peso
* Recuerda: d_ltc = D.log(CLP/USD). Baja = apreciacion del peso.
* =============================================================================
preserve
use `"$resdir\irf_var.irf"', clear
keep if irfname == "irf_var" & impulse == "tpm" & response == "d_ltc"
sort step

cap gen u95 = upper_oirf
if _rc != 0 { gen u95 = upper }
cap gen l95 = lower_oirf
if _rc != 0 { gen l95 = lower }
gen se  = (u95 - oirf) / invnormal(0.975)
gen u90 = oirf + invnormal(0.95) * se
gen l90 = oirf - invnormal(0.95) * se
gen zero = 0

twoway ///
    (rarea u95 l95 step, fcolor(dkgreen%10) lwidth(none)) ///
    (rarea u90 l90 step, fcolor(dkgreen%25) lwidth(none)) ///
    (line  oirf step,    lcolor(dkgreen) lwidth(medthick)) ///
    (line  zero step,    lcolor(black) lpattern(dash) lwidth(thin)), ///
    xlabel(0(4)24, labsize(small)) ///
    xtitle("Meses", size(small)) ///
    ytitle("Respuesta (OIRF)", size(small)) ///
    title("Shock TPM (+1 DE) → Tipo de Cambio [D.log TC]", size(medsmall)) ///
    subtitle("Signo esperado: NEGATIVO (apreciacion del peso: TC disminuye)", ///
             size(small) color(maroon)) ///
    note("IC 90% (banda oscura) y IC 95% (banda clara). Cholesky. `tmin'-`tmax'.", ///
         size(vsmall)) ///
    legend(order(2 "IC 90%" 1 "IC 95%" 3 "OIRF" 4 "y = 0") ///
           rows(1) size(vsmall)) ///
    name(irf_tpm_tc, replace)
restore
graph export `"$resdir\irf_tpm_tc.png"', name(irf_tpm_tc) replace width(950) height(650)
di "  -> irf_tpm_tc.png  (signo esperado: NEGATIVO)"

* =============================================================================
* IRF 3: Shock TC → Inflacion (D.log IPC) — Pass-through cambiario
* SIGNO ESPERADO: POSITIVO — depreciacion encarece importaciones y sube IPC
* =============================================================================
preserve
use `"$resdir\irf_var.irf"', clear
keep if irfname == "irf_var" & impulse == "d_ltc" & response == "d_lipc"
sort step

cap gen u95 = upper_oirf
if _rc != 0 { gen u95 = upper }
cap gen l95 = lower_oirf
if _rc != 0 { gen l95 = lower }
gen se  = (u95 - oirf) / invnormal(0.975)
gen u90 = oirf + invnormal(0.95) * se
gen l90 = oirf - invnormal(0.95) * se
gen zero = 0

twoway ///
    (rarea u95 l95 step, fcolor(maroon%10) lwidth(none)) ///
    (rarea u90 l90 step, fcolor(maroon%25) lwidth(none)) ///
    (line  oirf step,    lcolor(maroon) lwidth(medthick)) ///
    (line  zero step,    lcolor(black) lpattern(dash) lwidth(thin)), ///
    xlabel(0(4)24, labsize(small)) ///
    xtitle("Meses", size(small)) ///
    ytitle("Respuesta (OIRF)", size(small)) ///
    title("Shock D.log TC (+1 DE) → Inflacion [D.log IPC]", size(medsmall)) ///
    subtitle("Signo esperado: POSITIVO (depreciacion genera inflacion importada)", ///
             size(small) color(maroon)) ///
    note("IC 90% (banda oscura) y IC 95% (banda clara). Cholesky. `tmin'-`tmax'.", ///
         size(vsmall)) ///
    legend(order(2 "IC 90%" 1 "IC 95%" 3 "OIRF" 4 "y = 0") ///
           rows(1) size(vsmall)) ///
    name(irf_tc_ipc, replace)
restore
graph export `"$resdir\irf_tc_ipc.png"', name(irf_tc_ipc) replace width(950) height(650)
di "  -> irf_tc_ipc.png  (signo esperado: POSITIVO)"

* --- Resumen final -----------------------------------------------------------
di _newline "============================================================"
di "VAR_simple.do completado. Archivos en: $resdir"
di "  estabilidad_VAR.png"
di "  irf_tpm_ipc.png | irf_tpm_tc.png | irf_tc_ipc.png"
di "============================================================"
di "VERIFICACION DE SIGNOS (teoria economica):"
di "  irf_tpm_ipc -> debe ser NEGATIVO (al menos 1-12 meses)"
di "  irf_tpm_tc  -> debe ser NEGATIVO (apreciacion cambiaria)"
di "  irf_tc_ipc  -> debe ser POSITIVO (pass-through)"
di ""
di "Si algun signo es incorrecto: cambia 'local p = X' en Seccion 3"
di "y re-ejecuta el script completo. Valores a probar: 1  2  3  4  6"
di "============================================================"
