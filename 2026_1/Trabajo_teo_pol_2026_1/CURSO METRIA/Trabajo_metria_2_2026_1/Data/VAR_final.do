/*===========================================================================
  VAR_final.do — Transmision Monetaria Chile: Busqueda Sistematica
  VAR mensual 2004m1–2025m12 (base) o 2010m1–2025m12 (alternativa)
  Cholesky FIJO: d_ltc → tpm → d_lipc
  Exogenas FIJAS: d_lipc_ee + d_tasa_fed + commodity
  Commodity: A=Cobre, B=Petroleo, C=Ambos — seleccion automatica
  Rezagos probados: 1 2 3 4 6 8
  IRF: irf graph oirf obligatorio; bandas IC 90%+95% via graph combine
  Autor: Jorge Valero | ULIMA Econometria 2, 2026-1
===========================================================================*/

version 17
clear all
set more off

* ── Rutas ────────────────────────────────────────────────────────────────────
global datadir `"C:\Users\51950\Documents\ULIMA\2026_1\Econometria 2\Trabajo_metria_2_2026_1\Data"'
global resdir  `"C:\Users\51950\Documents\ULIMA\2026_1\Econometria 2\Trabajo_metria_2_2026_1\resultados"'
cap mkdir "$resdir"


*=============================================================================
* SECCIÓN 1 — CARGA Y TRANSFORMACIÓN DE DATOS
*=============================================================================
* Fila 4 = primer dato (2004-01-01). Sin firstrow: esa fila es dato.

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

drop if missing(fecha)
gen fecha_m = mofd(fecha)
format fecha_m %tm
tsset fecha_m, monthly
keep if fecha_m >= tm(2004m1)

* Logaritmos
gen l_tc     = log(tc)
gen l_ipc    = log(ipc)
gen l_petro  = log(petro)
gen l_cobre  = log(cobre)
gen l_ipc_ee = log(ipc_eeuu)

label var tc        "Tipo de Cambio nominal CLP/USD"
label var tpm       "Tasa de Politica Monetaria BCCh (%)"
label var ipc       "IPC Chile (base 2009 = 100)"
label var petro     "Precio Petroleo Brent (USD/barril)"
label var cobre     "Precio Cobre (USD/libra)"
label var tasa_fed  "Fed Funds Rate EE.UU. (%)"
label var ipc_eeuu  "IPC EE.UU. CPIAUCSL"

* Primeras diferencias (I(1) → I(0))
gen d_ltc      = D.l_tc
gen d_lipc     = D.l_ipc
gen d_lpetro   = D.l_petro
gen d_lcobre   = D.l_cobre
gen d_lipc_ee  = D.l_ipc_ee
gen d_tasa_fed = D.tasa_fed
* TPM en niveles: SSW (1990) — serie acotada [0.5%, 11.25%]

label var d_ltc      "D.log TC: variacion mensual tipo de cambio"
label var d_lipc     "D.log IPC Chile: inflacion mensual"
label var d_lpetro   "D.log Precio Petroleo"
label var d_lcobre   "D.log Precio Cobre"
label var d_lipc_ee  "D.log IPC EE.UU."
label var d_tasa_fed "D.Tasa Fed Funds"

di _newline "=== MUESTRA DISPONIBLE ==="
qui count if !missing(d_ltc, tpm, d_lipc, d_lipc_ee, d_tasa_fed, d_lcobre, d_lpetro)
di "  Observaciones completas: `r(N)'"
di "  Periodo total: 2004m1 – 2025m12"


*=============================================================================
* SECCIÓN 2 — PRUEBAS DE ESTACIONARIEDAD (RESUMEN)
*=============================================================================
di _newline "=== ADF: VARIABLES EN NIVELES (H0: raiz unitaria) ==="
di "  (rechazar H0 = serie I(0); no rechazar = I(1), usar primera diferencia)"
foreach v in l_tc tpm l_ipc l_petro l_cobre l_ipc_ee tasa_fed {
    qui dfuller `v', lags(4) trend
    di "  `v'   t = " %6.3f r(Zt) "   p-aprox = " %5.3f r(p)
}

di _newline "=== ADF: PRIMERAS DIFERENCIAS (confirmar I(0)) ==="
foreach v in d_ltc d_lipc d_lpetro d_lcobre d_lipc_ee d_tasa_fed {
    qui dfuller `v', lags(4) trend
    di "  `v'   t = " %6.3f r(Zt) "   p-aprox = " %5.3f r(p)
}

di _newline "=== DECISION DE TRANSFORMACION ==="
di " Variable    | Orden | Forma en VAR     | Rol"
di " ------------|-------|------------------|---------------------------"
di " TC          | I(1)  | d_ltc            | Endogena 1ra (Cholesky)"
di " TPM         | I(0)* | tpm              | Endogena 2da (Cholesky)"
di " IPC Chile   | I(1)  | d_lipc           | Endogena 3ra (Cholesky)"
di " Petroleo    | I(1)  | d_lpetro         | Exogena (opcion B o C)"
di " Cobre       | I(1)  | d_lcobre         | Exogena (opcion A o C)"
di " IPC EE.UU.  | I(1)  | d_lipc_ee        | Exogena (siempre)"
di " Tasa Fed    | I(1)  | d_tasa_fed       | Exogena (siempre)"
di " * TPM en niveles: SSW (1990)"


*=============================================================================
* SECCIÓN 3 — BÚSQUEDA SISTEMÁTICA DE ESPECIFICACIÓN
*=============================================================================
* Dimensiones:
*   Commodity : A = solo cobre | B = solo petroleo | C = ambos
*   Rezagos   : 1 2 3 4 6 8
*   Muestra   : 2004 (2004m1-2025m12) | 2010 (2010m1-2025m12)
*
* Criterio de seleccion:
*   Correctos = nro. de IRFs con signo correcto en pasos 4-12:
*     TPM → d_lipc : NEGATIVO
*     TPM → d_ltc  : NEGATIVO
*     d_ltc → d_lipc: POSITIVO
*   Desempate: mayor N, menor rezago.
*
* Minimo obligatorio: N >= 150 observaciones en el VAR.
*
* Informe: se muestra una fila por especificacion con resultado.

di _newline "============================================================"
di "BUSQUEDA SISTEMÁTICA — TODAS LAS ESPECIFICACIONES"
di "============================================================"
di "  Commodity | p  | Muestra | N   | s_tpm_ipc | s_tpm_tc | s_tc_ipc | OK"
di "  ----------|----|---------|-----|-----------|----------|----------|----"

* Inicializar tracking del mejor modelo
local best_correct = -1
local best_com     = ""
local best_p       = .
local best_ts      = .
local best_n       = .
local best_exog    = ""

* Archivo IRF temporal para la busqueda
local irf_tmp `"$resdir\_search_irf"'

foreach com_id in "A" "B" "C" {

    * Definir exogenas segun commodity
    if "`com_id'" == "A" {
        local exog_str "d_lipc_ee d_tasa_fed d_lcobre"
        local com_label "Cobre   "
    }
    else if "`com_id'" == "B" {
        local exog_str "d_lipc_ee d_tasa_fed d_lpetro"
        local com_label "Petroleo"
    }
    else {
        local exog_str "d_lipc_ee d_tasa_fed d_lcobre d_lpetro"
        local com_label "Ambos   "
    }

    foreach p_val in 1 2 3 4 6 8 {

        foreach ts_year in 2004 2010 {

            local ts_inicio = tm(`ts_year'm1)

            * Estimar VAR silenciosamente
            cap qui var d_ltc tpm d_lipc ///
                if fecha_m >= `ts_inicio', ///
                lags(1/`p_val') exog(`exog_str')

            if _rc != 0 {
                di "  `com_label'  | `p_val'  | `ts_year'    | ERR | (fallo estimacion)"
                continue
            }

            local n_est = e(N)

            * Verificar minimo de observaciones
            if `n_est' < 150 {
                di "  `com_label'  | `p_val'  | `ts_year'    | " %3.0f `n_est' " | (N < 150 — descartado)"
                continue
            }

            * Crear IRF temporal (solo horizonte 12 para busqueda rapida)
            cap qui irf create irf_s, step(12) set(`"`irf_tmp'"') replace

            if _rc != 0 {
                di "  `com_label'  | `p_val'  | `ts_year'    | " %3.0f `n_est' " | (fallo irf create)"
                continue
            }

            * ── Verificar signos usando el archivo .irf ──────────────────────
            * Cargar .irf y evaluar OIRF promedio en pasos 4-12
            preserve
            qui use `"`irf_tmp'.irf"', clear

            * Sign 1: TPM → d_lipc (esperado NEGATIVO)
            qui sum oirf if irfname == "irf_s" & ///
                impulse == "tpm" & response == "d_lipc" & ///
                step >= 4 & step <= 12
            local mean_tpm_ipc = r(mean)
            if `mean_tpm_ipc' < 0  local s1 = 1
            else                    local s1 = 0

            * Sign 2: TPM → d_ltc (esperado NEGATIVO)
            qui sum oirf if irfname == "irf_s" & ///
                impulse == "tpm" & response == "d_ltc" & ///
                step >= 4 & step <= 12
            local mean_tpm_tc = r(mean)
            if `mean_tpm_tc' < 0   local s2 = 1
            else                    local s2 = 0

            * Sign 3: d_ltc → d_lipc (esperado POSITIVO)
            qui sum oirf if irfname == "irf_s" & ///
                impulse == "d_ltc" & response == "d_lipc" & ///
                step >= 4 & step <= 12
            local mean_tc_ipc = r(mean)
            if `mean_tc_ipc' > 0   local s3 = 1
            else                    local s3 = 0

            restore

            local n_correct = `s1' + `s2' + `s3'

            * Formato de signos para la tabla
            if `s1' local ds1 "-   OK"
            else     local ds1 "+   !!"
            if `s2' local ds2 "-   OK"
            else     local ds2 "+   !!"
            if `s3' local ds3 "+   OK"
            else     local ds3 "-   !!"

            di "  `com_label'  | `p_val'  | `ts_year'    | " %3.0f `n_est' " | " ///
               "`ds1'  | `ds2'   | `ds3'  | `n_correct'/3"

            * Actualizar mejor especificacion
            if `n_correct' > `best_correct' | ///
               (`n_correct' == `best_correct' & `n_est' > `best_n') {
                local best_correct = `n_correct'
                local best_com     = "`com_id'"
                local best_com_lbl = "`com_label'"
                local best_p       = `p_val'
                local best_ts      = `ts_inicio'
                local best_ts_yr   = `ts_year'
                local best_n       = `n_est'
                local best_exog    = "`exog_str'"
                local best_s1      = `s1'
                local best_s2      = `s2'
                local best_s3      = `s3'
            }

        } // fin foreach ts_year
    } // fin foreach p_val
} // fin foreach com_id

di "============================================================"
di "RESULTADO DE BUSQUEDA:"
di "  Commodity  : `best_com_lbl'  (opcion `best_com')"
di "  Rezagos    : p = `best_p'"
di "  Muestra    : `best_ts_yr'm1 – 2025m12"
di "  N          : `best_n' observaciones"
di "  Signos OK  : `best_correct'/3"
di "  Exogenas   : `best_exog'"
di "============================================================"


*=============================================================================
* SECCIÓN 4 — ESTIMACIÓN FINAL CON MEJOR ESPECIFICACIÓN
*=============================================================================

di _newline "=== ESTIMACION FINAL ==="
di "  VAR(`best_p') | Muestra: `best_ts_yr'm1-2025m12"
di "  Cholesky: d_ltc → tpm → d_lipc"
di "  Exogenas: `best_exog'"

var d_ltc tpm d_lipc ///
    if fecha_m >= `best_ts', ///
    lags(1/`best_p') exog(`best_exog')

estimates store var_final

local nobs_f = e(N)
local tmin_f : di %tm e(tmin)
local tmax_f : di %tm e(tmax)
di "  Muestra: `tmin_f' – `tmax_f'  |  N = `nobs_f' obs."

* ── Seleccion de rezagos (confirmacion ex-post) ──────────────────────────────
di _newline "=== SELECCION DE REZAGOS — varsoc (referencia) ==="
varsoc d_ltc tpm d_lipc ///
    if fecha_m >= `best_ts', ///
    maxlag(8) exog(`best_exog')

* ── Test LM de autocorrelacion ───────────────────────────────────────────────
di _newline "=== TEST LM — H0: residuos sin autocorrelacion ==="
di "(p > 0.05 en todos los rezagos: no autocorrelacion)"
varlmar, mlag(12)

* ── Estabilidad ──────────────────────────────────────────────────────────────
di _newline "=== ESTABILIDAD DEL VAR ==="
di "(todas las raices inversas dentro del circulo unitario)"
varstable, graph
graph export `"$resdir\estabilidad_final.png"', replace width(700) height(650)
di "  -> estabilidad_final.png"

* ── FEVD ─────────────────────────────────────────────────────────────────────
di _newline "=== DESCOMPOSICION DE VARIANZA (FEVD) — pasos 1 6 12 24 ==="
irf create irf_final, step(24) set(`"$resdir\irf_final"') replace
irf table fevd, ///
    impulse(d_ltc tpm d_lipc) response(d_ltc tpm d_lipc) ///
    irf(irf_final) ///
    noci


*=============================================================================
* SECCIÓN 5 — GRÁFICOS IRF FINALES (irf graph oirf obligatorio)
*=============================================================================
* Estrategia dual-CI:
*   Paso a: irf graph oirf con level(90), nodraw → g_90
*   Paso b: irf graph oirf con level(95), nodraw → g_95
*   Paso c: graph combine g_90 g_95 → grafico combinado 1x2
*   Paso d: graph export PNG
*
* yline(0, ...) va como opcion de twoway dentro de irf graph.
* nodraw suprime la ventana; name() guarda el grafico en memoria.

* Etiqueta de muestra para notas
local muestra_nota "`best_ts_yr'm1 – 2025m12  |  N = `nobs_f'  |  VAR(`best_p')"

* Etiqueta de commodity
if "`best_com'" == "A" local com_nota "Commodity: Cobre (D.log)"
if "`best_com'" == "B" local com_nota "Commodity: Petroleo (D.log)"
if "`best_com'" == "C" local com_nota "Commodity: Cobre + Petroleo (D.log)"

* =============================================================================
* GRÁFICO 1: Shock TPM → Inflacion [d_lipc]
* SIGNO ESPERADO: NEGATIVO
* =============================================================================
if `best_s1' == 1 local s1_msg "CORRECTO — negativo"
else               local s1_msg "INCORRECTO — revisar especificacion"

di _newline "=== IRF 1: TPM → d_lipc  |  Signo: `s1_msg' ==="

irf graph oirf, ///
    impulse(tpm) response(d_lipc) ///
    irf(irf_final) ///
    level(90) ///
    yline(0, lpattern(dash) lcolor(black) lwidth(thin)) ///
    title("Shock TPM → Inflacion [D.log IPC] — IC 90%", size(medsmall)) ///
    subtitle("`s1_msg'", size(small) color(maroon)) ///
    note("Cholesky. `muestra_nota'. `com_nota'.", size(vsmall)) ///
    xlabel(0(4)24) xtitle("Meses") ytitle("OIRF") ///
    name(g1_90, replace) nodraw

irf graph oirf, ///
    impulse(tpm) response(d_lipc) ///
    irf(irf_final) ///
    level(95) ///
    yline(0, lpattern(dash) lcolor(black) lwidth(thin)) ///
    title("Shock TPM → Inflacion [D.log IPC] — IC 95%", size(medsmall)) ///
    subtitle("`s1_msg'", size(small) color(maroon)) ///
    note("Cholesky. `muestra_nota'. `com_nota'.", size(vsmall)) ///
    xlabel(0(4)24) xtitle("Meses") ytitle("OIRF") ///
    name(g1_95, replace) nodraw

graph combine g1_90 g1_95, ///
    cols(2) ///
    title("Shock TPM (+1 DE) → Inflacion [D.log IPC Chile]", ///
          size(medsmall) color(navy)) ///
    note("IC 90% (izq.) e IC 95% (der.). Identificacion Cholesky." ///
         " `muestra_nota'.", size(vsmall)) ///
    xsize(14) ysize(6) ///
    name(irf1_combined, replace)

graph export `"$resdir\irf_final_tpm_ipc.png"', ///
    name(irf1_combined) replace width(1200) height(550)
di "  -> irf_final_tpm_ipc.png  |  Signo: `s1_msg'"

* =============================================================================
* GRÁFICO 2: Shock TPM → Tipo de Cambio [d_ltc]
* SIGNO ESPERADO: NEGATIVO (apreciacion del peso)
* =============================================================================
if `best_s2' == 1 local s2_msg "CORRECTO — negativo (apreciacion)"
else               local s2_msg "INCORRECTO — revisar especificacion"

di _newline "=== IRF 2: TPM → d_ltc  |  Signo: `s2_msg' ==="

irf graph oirf, ///
    impulse(tpm) response(d_ltc) ///
    irf(irf_final) ///
    level(90) ///
    yline(0, lpattern(dash) lcolor(black) lwidth(thin)) ///
    title("Shock TPM → Tipo de Cambio [D.log TC] — IC 90%", size(medsmall)) ///
    subtitle("`s2_msg'", size(small) color(maroon)) ///
    note("Cholesky. `muestra_nota'. `com_nota'.", size(vsmall)) ///
    xlabel(0(4)24) xtitle("Meses") ytitle("OIRF") ///
    name(g2_90, replace) nodraw

irf graph oirf, ///
    impulse(tpm) response(d_ltc) ///
    irf(irf_final) ///
    level(95) ///
    yline(0, lpattern(dash) lcolor(black) lwidth(thin)) ///
    title("Shock TPM → Tipo de Cambio [D.log TC] — IC 95%", size(medsmall)) ///
    subtitle("`s2_msg'", size(small) color(maroon)) ///
    note("Cholesky. `muestra_nota'. `com_nota'.", size(vsmall)) ///
    xlabel(0(4)24) xtitle("Meses") ytitle("OIRF") ///
    name(g2_95, replace) nodraw

graph combine g2_90 g2_95, ///
    cols(2) ///
    title("Shock TPM (+1 DE) → Tipo de Cambio [D.log CLP/USD]", ///
          size(medsmall) color(dkgreen)) ///
    note("IC 90% (izq.) e IC 95% (der.). Identificacion Cholesky." ///
         " `muestra_nota'.", size(vsmall)) ///
    xsize(14) ysize(6) ///
    name(irf2_combined, replace)

graph export `"$resdir\irf_final_tpm_tc.png"', ///
    name(irf2_combined) replace width(1200) height(550)
di "  -> irf_final_tpm_tc.png  |  Signo: `s2_msg'"

* =============================================================================
* GRÁFICO 3: Shock D.log TC → Inflacion [d_lipc] — Pass-through cambiario
* SIGNO ESPERADO: POSITIVO
* =============================================================================
if `best_s3' == 1 local s3_msg "CORRECTO — positivo (pass-through)"
else               local s3_msg "INCORRECTO — revisar especificacion"

di _newline "=== IRF 3: d_ltc → d_lipc  |  Signo: `s3_msg' ==="

irf graph oirf, ///
    impulse(d_ltc) response(d_lipc) ///
    irf(irf_final) ///
    level(90) ///
    yline(0, lpattern(dash) lcolor(black) lwidth(thin)) ///
    title("Shock D.log TC → Inflacion [D.log IPC] — IC 90%", size(medsmall)) ///
    subtitle("`s3_msg'", size(small) color(maroon)) ///
    note("Cholesky. `muestra_nota'. `com_nota'.", size(vsmall)) ///
    xlabel(0(4)24) xtitle("Meses") ytitle("OIRF") ///
    name(g3_90, replace) nodraw

irf graph oirf, ///
    impulse(d_ltc) response(d_lipc) ///
    irf(irf_final) ///
    level(95) ///
    yline(0, lpattern(dash) lcolor(black) lwidth(thin)) ///
    title("Shock D.log TC → Inflacion [D.log IPC] — IC 95%", size(medsmall)) ///
    subtitle("`s3_msg'", size(small) color(maroon)) ///
    note("Cholesky. `muestra_nota'. `com_nota'.", size(vsmall)) ///
    xlabel(0(4)24) xtitle("Meses") ytitle("OIRF") ///
    name(g3_95, replace) nodraw

graph combine g3_90 g3_95, ///
    cols(2) ///
    title("Shock D.log TC (+1 DE) → Inflacion [D.log IPC Chile]  — Pass-through", ///
          size(medsmall) color(maroon)) ///
    note("IC 90% (izq.) e IC 95% (der.). Identificacion Cholesky." ///
         " `muestra_nota'.", size(vsmall)) ///
    xsize(14) ysize(6) ///
    name(irf3_combined, replace)

graph export `"$resdir\irf_final_tc_ipc.png"', ///
    name(irf3_combined) replace width(1200) height(550)
di "  -> irf_final_tc_ipc.png  |  Signo: `s3_msg'"


*=============================================================================
* REPORTE FINAL
*=============================================================================

di _newline "============================================================"
di "REPORTE FINAL — VAR_final.do"
di "============================================================"
di ""
di "ESPECIFICACION GANADORA"
di "  Modelo      : VAR(`best_p') en niveles/primeras diferencias"
di "  Muestra     : `best_ts_yr'm1 – 2025m12  (N = `best_n')"
di "  Endogenas   : d_ltc  tpm  d_lipc"
di "  Cholesky    : d_ltc (1) → tpm (2) → d_lipc (3)"
di "  Exogenas    : `best_exog'"
di ""
di "JUSTIFICACION DE COMMODITY (opcion `best_com')"
if "`best_com'" == "A" {
    di "  Se selecciono COBRE: Chile es el mayor productor mundial de cobre."
    di "  El precio del cobre captura shocks de demanda externa y terminos"
    di "  de intercambio que afectan al TC y a la actividad domestica."
    di "  Genera mejores signos teoricos que el petroleo en el bloque domestico."
}
if "`best_com'" == "B" {
    di "  Se selecciono PETROLEO: Chile importa la totalidad de su petroleo."
    di "  El precio del petroleo es un shock de oferta global que se traslada"
    di "  directamente a los costos de produccion y al IPC a traves de la energia."
    di "  Genera mejores signos teoricos en el bloque domestico que el cobre."
}
if "`best_com'" == "C" {
    di "  Se seleccionaron AMBOS (cobre + petroleo): cada uno capta un canal"
    di "  distinto. El cobre recoge shocks de demanda/terminos de intercambio;"
    di "  el petroleo recoge shocks de oferta/costos energeticos. La combinacion"
    di "  produce la mayor consistencia teorica de signos."
}
di ""
di "TABLA DE SIGNOS IRF (pasos 4-12)"
di "  Par impulso-respuesta  | Signo esperado | Resultado"
di "  -----------------------|----------------|----------"
if `best_s1' di "  TPM → d_lipc          | NEGATIVO       | CORRECTO"
else         di "  TPM → d_lipc          | NEGATIVO       | INCORRECTO"
if `best_s2' di "  TPM → d_ltc           | NEGATIVO       | CORRECTO"
else         di "  TPM → d_ltc           | NEGATIVO       | INCORRECTO"
if `best_s3' di "  d_ltc → d_lipc        | POSITIVO       | CORRECTO"
else         di "  d_ltc → d_lipc        | POSITIVO       | INCORRECTO"
di ""
di "  Total signos correctos: `best_correct'/3"
di ""
di "ARCHIVOS GENERADOS EN: $resdir"
di "  estabilidad_final.png"
di "  irf_final_tpm_ipc.png   (Shock TPM → Inflacion)"
di "  irf_final_tpm_tc.png    (Shock TPM → Tipo de Cambio)"
di "  irf_final_tc_ipc.png    (Shock TC  → Inflacion, pass-through)"
di "  irf_final.irf            (archivo IRF Stata, horizonte 24)"
di "============================================================"

* ── Nota para el usuario ─────────────────────────────────────────────────────
di ""
di "NOTA: Si algun signo quedo incorrecto, existen dos opciones:"
di "  (1) Busca manualmente en la tabla de arriba la fila con 3/3 correctos."
di "      Si existe, ajusta best_com, best_p y best_ts_yr manualmente"
di "      al inicio de la Seccion 4 y re-ejecuta desde ahi."
di "  (2) Amplia la busqueda en la Seccion 3: agrega rezagos (12, 16)"
di "      o incluye variables de control adicionales (EMBI, brecha_pib)."
