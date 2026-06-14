/*==============================================================
  VAR ENGINE v2.0  |  Stata 17  |  Jorge Valero ULIMA
  Operado por Claude desde terminal.

  Sin pausas internas. Todas las decisiones configuradas
  en Sección 0 antes de ejecutar.

  OUTPUTS  →  output/run_YYYYMMDD/
    log_var.txt              todos los resultados
    diag_estabilidad.png     círculo unitario
    irf_[imp]_[resp].png     un PNG por par impulso-respuesta
    fevd_[resp].png          un PNG por variable respuesta
    irf_main.irf             archivo IRF Stata
    scripts/var_simple.do    estilo codigos.do (~40 líneas)
    scripts/var_completo.do  copia de este engine
==============================================================*/

version 17
clear all
set more off
set scheme s2color

/* ════════════════════════════════════════════════════════════
   SECCIÓN 0  CONFIGURACIÓN
   ════════════════════════════════════════════════════════════ */

* ── Archivo ──────────────────────────────────────────────────
global datafile   `"input/datos_var.xlsx"'
global hoja       "datos"
global cellrange  "A1"        // headers en fila 1, fecha estándar Excel

* ── Tiempo ───────────────────────────────────────────────────
* datemode "auto"   = fechas estándar Excel (usa mofd(datevar))
* datemode "manual" = fechas en texto → genera índice tm(startdate)+_n-1
global datemode   "auto"
global startdate  "2004m1"
global datevar    "fecha"     // nombre de la columna de fecha
global freq       "monthly"
global inicio     "2004m1"
global fin        "2021m12"   // IPC solo disponible hasta Dic 2021

* ── Variables (nombres exactos del Excel) ────────────────────
global rawvars  "tasa_ref tamn ipc tc"

* ── Transformación por variable ──────────────────────────────
* Opciones: level | log | diff | diff3 | logdiff | pct | pct3
global trans_tasa_ref  "level"   // tasa política: niveles (argumento SSW)
global trans_tamn      "level"   // tasa activa:   niveles (argumento SSW)
global trans_ipc       "pct"     // inflación mensual = (ipc/L.ipc-1)*100
global trans_tc        "level"   // tipo de cambio: niveles

* ── Especificación VAR ───────────────────────────────────────
* Orden Cholesky: más exógena → más endógena
global chol_order  "tasa_ref tc tamn pct_ipc"
global exogenas    ""
global p_use       2            // canal_var.do referencia usa p=2; ajustar tras varsoc
global p_max       6

* ── IRF ─────────────────────────────────────────────────────
global irf_step    24
global irf_ic      90

* ── Rutas ────────────────────────────────────────────────────
global resdir      `"output"'

/* ════════════════════════════════════════════════════════════
   SETUP — carpetas y log
   ════════════════════════════════════════════════════════════ */

cap mkdir "$resdir"
local dd = string(day(today()),   "%02.0f")
local mm = string(month(today()), "%02.0f")
local yy = string(year(today()))
global outdir  `"$resdir/run_`yy'`mm'`dd'"'
global scrpdir `"$outdir/scripts"'
cap mkdir "$outdir"
cap mkdir "$scrpdir"

log using `"$outdir/log_var.smcl"', replace smcl name(varlog)


/* ════════════════════════════════════════════════════════════
   SECCIÓN 1  IMPORTAR + DIAGNÓSTICO DE DATOS
   ════════════════════════════════════════════════════════════ */

di _newline(2) as result ///
  "══════════════════════════════════════════" _n ///
  "  SECCIÓN 1: IMPORTAR Y DIAGNOSTICAR"     _n ///
  "══════════════════════════════════════════"

import excel using "$datafile", ///
    sheet("$hoja") cellrange("$cellrange") firstrow clear

* ── Configurar serie de tiempo ───────────────────────────────
if "$datemode" == "manual" {
    gen __t = tm($startdate) + _n - 1
    format __t %tm
    tsset __t, monthly
    keep if __t >= tm($inicio) & __t <= tm($fin)
}
else {
    gen __t = mofd($datevar)
    format __t %tm
    tsset __t, monthly
    keep if __t >= tm($inicio) & __t <= tm($fin)
}

qui count
local total_obs = r(N)
di as result "  Observaciones: `total_obs'  |  Período: $inicio – $fin"

* ── Diagnóstico de calidad de datos ─────────────────────────
di _newline as text "── Diagnóstico de datos ──"
di as text "  Variable          Obs válidas  Missings  %Miss   Min         Max"
di as text "  ──────────────────────────────────────────────────────────────────"

local any_miss_alert = 0

foreach v of global rawvars {
    cap confirm variable `v'
    if _rc != 0 {
        di as error "  ERROR: la variable '`v'' no existe en el Excel."
        di as error "  Variables disponibles:"
        describe, simple
        log close varlog
        exit 1
    }

    qui count if missing(`v')
    local nmiss   = r(N)
    local nvalid  = `total_obs' - `nmiss'
    local pctmiss = round(`nmiss' / `total_obs' * 100, 0.1)
    qui sum `v'
    local vmin = r(min)
    local vmax = r(max)

    if `pctmiss' > 20 {
        di as error "  " %-18s "`v'" %6.0f `nvalid' "      " %5.0f `nmiss' ///
            "    " %5.1f `pctmiss' "%   " %10.3g `vmin' "  " %10.3g `vmax' ///
            "  ← ALERTA >20%"
        local any_miss_alert = 1
    }
    else if `pctmiss' > 0 {
        di as text "  " %-18s "`v'" %6.0f `nvalid' "      " %5.0f `nmiss' ///
            "    " %5.1f `pctmiss' "%   " %10.3g `vmin' "  " %10.3g `vmax' ///
            "  ← missings"
    }
    else {
        di as result "  " %-18s "`v'" %6.0f `nvalid' "      " %5.0f `nmiss' ///
            "    " %5.1f `pctmiss' "%   " %10.3g `vmin' "  " %10.3g `vmax'
    }
}

if `any_miss_alert' {
    di as error _n "  AVISO: variables con >20% missings → reportado a Claude."
}

qui tsreport
if r(N_gaps) > 0 {
    di as error _n "  ALERTA: `=r(N_gaps)' gap(s) en la serie → reportado a Claude."
}
else {
    di as result _n "  Continuidad temporal: sin gaps ✓"
}

di _newline as text "── Estadísticas descriptivas ──"
summarize $rawvars


/* ════════════════════════════════════════════════════════════
   SECCIÓN 2  RAÍCES UNITARIAS — NIVELES
   ADF (4 lags, c+t)  +  PP (NW=8, c+t)
   ════════════════════════════════════════════════════════════ */

di _newline(2) as result ///
  "══════════════════════════════════════════"         _n ///
  "  SECCIÓN 2: RAÍCES UNITARIAS (NIVELES)"           _n ///
  "══════════════════════════════════════════"         _n ///
  "  H0: raíz unitaria.  p < 0.05 → rechazar → I(0)" _n

di as text "  Variable       ADF-t    ADF-p    PP-t     PP-p     Trans. config."
di as text "  ─────────────  ──────   ──────   ──────   ──────   ─────────────"

foreach v of global rawvars {
    qui dfuller `v', lags(4) trend
    local t_adf = r(Zt)
    local p_adf = r(p)
    qui pperron `v', lags(8) trend
    local t_pp  = r(Zt)
    local p_pp  = r(p)

    di as text "  " %-14s "`v'" ///
       %7.3f `t_adf' "  " %6.3f `p_adf' "  " ///
       %7.3f `t_pp'  "  " %6.3f `p_pp'  "  " ///
       "→ ${trans_`v'}"
}


/* ════════════════════════════════════════════════════════════
   SECCIÓN 3  APLICAR TRANSFORMACIONES
   ════════════════════════════════════════════════════════════ */

di _newline(2) as result ///
  "══════════════════════════════════════════" _n ///
  "  SECCIÓN 3: TRANSFORMACIONES"            _n ///
  "══════════════════════════════════════════"

global finalvars ""

foreach v of global rawvars {
    local tr "${trans_`v'}"

    if "`tr'" == "level" {
        global finalvars "$finalvars `v'"
        global fname_`v' "`v'"
        di as text "  `v'  →  nivel  [`v']"
    }
    else if "`tr'" == "log" {
        cap drop l_`v'
        gen l_`v' = log(`v')
        label var l_`v' "log(`v')"
        global finalvars "$finalvars l_`v'"
        global fname_`v' "l_`v'"
        di as text "  `v'  →  log(`v')  [l_`v']"
    }
    else if "`tr'" == "diff" {
        cap drop d_`v'
        gen d_`v' = D.`v'
        label var d_`v' "D.`v'"
        global finalvars "$finalvars d_`v'"
        global fname_`v' "d_`v'"
        di as text "  `v'  →  D.`v'  [d_`v']"
    }
    else if "`tr'" == "diff3" {
        cap drop d3_`v'
        gen d3_`v' = `v' - L3.`v'
        label var d3_`v' "D3.`v' (var. trim.)"
        global finalvars "$finalvars d3_`v'"
        global fname_`v' "d3_`v'"
        di as text "  `v'  →  `v' - L3.`v'  [d3_`v']"
    }
    else if "`tr'" == "logdiff" {
        cap drop dl_`v'
        gen dl_`v' = D.log(`v')
        label var dl_`v' "D.log(`v')"
        global finalvars "$finalvars dl_`v'"
        global fname_`v' "dl_`v'"
        di as text "  `v'  →  D.log(`v')  [dl_`v']"
    }
    else if "`tr'" == "pct" {
        cap drop pct_`v'
        gen pct_`v' = (`v' / L.`v' - 1) * 100
        label var pct_`v' "Var.% `v'"
        global finalvars "$finalvars pct_`v'"
        global fname_`v' "pct_`v'"
        di as text "  `v'  →  var.% mensual  [pct_`v']"
    }
    else if "`tr'" == "pct3" {
        cap drop pct3_`v'
        gen pct3_`v' = (`v' / L3.`v' - 1) * 100
        label var pct3_`v' "Var.% trim. `v'"
        global finalvars "$finalvars pct3_`v'"
        global fname_`v' "pct3_`v'"
        di as text "  `v'  →  (`v'/L3.`v'-1)*100  [pct3_`v']"
    }
    else {
        di as error "  ERROR: transformación '`tr'' desconocida para `v'."
        di as error "  Opciones: level | log | diff | diff3 | logdiff | pct | pct3"
        log close varlog
        exit 1
    }
}

di as result _n "  Variables finales: $finalvars"


/* ════════════════════════════════════════════════════════════
   SECCIÓN 4  CONFIRMAR I(0) — VARIABLES TRANSFORMADAS
   ════════════════════════════════════════════════════════════ */

di _newline(2) as result ///
  "══════════════════════════════════════════════════════" _n ///
  "  SECCIÓN 4: CONFIRMAR I(0) — VARS. TRANSFORMADAS"   _n ///
  "══════════════════════════════════════════════════════"

local still_nonstat = 0
foreach v of global finalvars {
    qui pperron `v', lags(8) trend
    local pv = r(p)
    if `pv' < 0.05 {
        di as result "  %-24s p = " %5.3f `pv' "  I(0) ✓"
    }
    else {
        di as error  "  %-24s p = " %5.3f `pv' "  ALERTA: sigue no-estacionaria"
        local still_nonstat = 1
    }
}

if `still_nonstat' {
    di as error _n "  AVISO: variable(s) no-estacionarias → reportado a Claude."
}


/* ════════════════════════════════════════════════════════════
   SECCIÓN 5  SELECCIÓN DE REZAGOS (varsoc)
   p_use configurado en Sección 0. varsoc corre para registro.
   ════════════════════════════════════════════════════════════ */

di _newline(2) as result ///
  "══════════════════════════════════════════" _n ///
  "  SECCIÓN 5: SELECCIÓN DE REZAGOS"        _n ///
  "══════════════════════════════════════════"

if "$exogenas" != "" {
    varsoc $finalvars, maxlag($p_max) exog($exogenas)
}
else {
    varsoc $finalvars, maxlag($p_max)
}

di as result _n "  Rezagos en estimación: p = $p_use  (configurado en Sección 0)"


/* ════════════════════════════════════════════════════════════
   SECCIÓN 6  ESTIMAR VAR + IRF CREATE
   ════════════════════════════════════════════════════════════ */

di _newline(2) as result ///
  "══════════════════════════════════════════" _n ///
  "  SECCIÓN 6: ESTIMANDO VAR($p_use)"       _n ///
  "══════════════════════════════════════════"
di as text "  Endógenas : $chol_order"
di as text "  Exógenas  : $exogenas"
di as text "  Cholesky  : identificación recursiva"

if "$exogenas" != "" {
    var $chol_order, lags(1/$p_use) exog($exogenas)
}
else {
    var $chol_order, lags(1/$p_use)
}

estimates store var_main

local nobs   = e(N)
local neq    = e(k_eq)
local nparam = `neq' * `neq' * $p_use
di as result _n "  Obs: `nobs'  |  Ecuaciones: `neq'  |  Parámetros: `nparam'"

if `nparam' > `nobs' / 5 {
    di as error "  AVISO: muchos parámetros para el tamaño muestral (T/k²p < 5)."
}

irf create irf_main, step($irf_step) set(`"$outdir/irf_main"') replace


/* ════════════════════════════════════════════════════════════
   SECCIÓN 7  DIAGNÓSTICOS COMPLETOS
   ════════════════════════════════════════════════════════════ */

di _newline(2) as result ///
  "══════════════════════════════════════════" _n ///
  "  SECCIÓN 7: DIAGNÓSTICOS"                _n ///
  "══════════════════════════════════════════"

* 7a. Estabilidad
di _newline as result "── 7a. Estabilidad (raíces características) ──"
varstable, graph dlabel
graph export `"$outdir/diag_estabilidad.png"', replace width(700) height(650)
di as text "  → diag_estabilidad.png"

* 7b. Autocorrelación LM
di _newline as result "── 7b. Test LM de autocorrelación ──"
di as text "  H0: sin autocorrelación serial. Deseable: p > 0.05."
varlmar, mlag(12)

* 7c. Normalidad (informativa)
di _newline as result "── 7c. Normalidad Jarque-Bera (informativa) ──"
varnorm, jbera

* 7d. Significancia de rezagos
di _newline as result "── 7d. Significancia de rezagos — Wald test ──"
varwle

* 7e. Causalidad de Granger
di _newline as result "── 7e. Causalidad de Granger ──"
vargranger


/* ════════════════════════════════════════════════════════════
   SECCIÓN 8  IRF — ORTOGONALIZADAS (todos los pares)
   ════════════════════════════════════════════════════════════ */

di _newline(2) as result ///
  "══════════════════════════════════════════" _n ///
  "  SECCIÓN 8: IMPULSO-RESPUESTA (OIRF)"    _n ///
  "══════════════════════════════════════════"

local vars_list "$chol_order"

foreach imp of local vars_list {
    foreach resp of local vars_list {

        irf graph oirf,                                                              ///
            impulse(`imp') response(`resp') irf(irf_main)                            ///
            level($irf_ic)                                                           ///
            yline(0, lpattern(dash) lcolor(gs8) lwidth(thin))                       ///
            xlabel(0(4)$irf_step) xtitle("Períodos") ytitle("OIRF")                ///
            title("Shock `imp' → `resp'", size(medsmall) color(navy))               ///
            note("Cholesky. IC $irf_ic%. p=$p_use. $inicio–$fin", size(vsmall))    ///
            name(g_irf, replace)

        graph export `"$outdir/irf_`imp'_`resp'.png"',                              ///
            replace width(900) height(580)
        di as text "  → irf_`imp'_`resp'.png"
    }
}


/* ════════════════════════════════════════════════════════════
   SECCIÓN 9  FEVD — DESCOMPOSICIÓN DE VARIANZA
   ════════════════════════════════════════════════════════════ */

di _newline(2) as result ///
  "══════════════════════════════════════════" _n ///
  "  SECCIÓN 9: DESCOMPOSICIÓN DE VARIANZA"  _n ///
  "══════════════════════════════════════════"

irf table fevd, irf(irf_main) ///
    impulse($chol_order) response($chol_order) noci

foreach resp of local vars_list {
    irf graph fevd,                                                               ///
        response(`resp') irf(irf_main)                                            ///
        title("FEVD: `resp'", size(medsmall) color(maroon))                      ///
        note("Cholesky. p=$p_use. $inicio–$fin", size(vsmall))                   ///
        name(g_fevd, replace)
    graph export `"$outdir/fevd_`resp'.png"',                                    ///
        replace width(900) height(580)
    di as text "  → fevd_`resp'.png"
}


/* ════════════════════════════════════════════════════════════
   SECCIÓN 10  GENERAR var_simple.do
   Estilo codigos.do: lineal, sin automatización, ~40 líneas.
   ════════════════════════════════════════════════════════════ */

di _newline(2) as result ///
  "══════════════════════════════════════════" _n ///
  "  SECCIÓN 10: GENERANDO var_simple.do"    _n ///
  "══════════════════════════════════════════"

file open fsimple using `"$scrpdir/var_simple.do"', write replace text

file write fsimple "/*=============================================" _n
file write fsimple "  VAR Simple"                                     _n
file write fsimple "  Especificación: VAR($p_use) | $inicio–$fin"    _n
file write fsimple "  Cholesky: $chol_order"                         _n
file write fsimple "  Generado por: var_completo.do"                 _n
file write fsimple "=============================================*/"  _n _n

file write fsimple "*** Configurar serie de tiempo ***"               _n
file write fsimple "gen time = tm($startdate) + _n - 1"              _n
file write fsimple "format time %tm"                                  _n
file write fsimple "tsset time, monthly"                              _n _n

file write fsimple "*** Prueba de estacionariedad (niveles) ***"     _n
foreach v of global rawvars {
    file write fsimple "pperron `v'" _n
}
file write fsimple _n

file write fsimple "*** Crear variables ***"                          _n
foreach v of global rawvars {
    local tr "${trans_`v'}"
    if "`tr'" == "level"   file write fsimple "* `v': se usa en niveles" _n
    if "`tr'" == "log"     file write fsimple "gen l_`v' = log(`v')" _n
    if "`tr'" == "diff"    file write fsimple "gen d_`v' = D.`v'" _n
    if "`tr'" == "diff3"   file write fsimple "gen d3_`v' = `v' - L3.`v'" _n
    if "`tr'" == "logdiff" file write fsimple "gen dl_`v' = D.log(`v')" _n
    if "`tr'" == "pct"     file write fsimple "gen pct_`v' = (`v'/L.`v'-1)*100" _n
    if "`tr'" == "pct3"    file write fsimple "gen pct3_`v' = (`v'/L3.`v'-1)*100" _n
}
file write fsimple _n

file write fsimple "*** Prueba de estacionariedad (vars. transformadas) ***" _n
foreach v of global finalvars {
    file write fsimple "pperron `v'" _n
}
file write fsimple _n

file write fsimple "*** Selección de rezagos ***"                     _n
if "$exogenas" != "" {
    file write fsimple "varsoc $chol_order, maxlag($p_max) exog($exogenas)" _n _n
}
else {
    file write fsimple "varsoc $chol_order, maxlag($p_max)" _n _n
}

file write fsimple "*** Estimación VAR ***"                           _n
if "$exogenas" != "" {
    file write fsimple "var $chol_order, lags(1/$p_use) exog($exogenas)" _n _n
}
else {
    file write fsimple "var $chol_order, lags(1/$p_use)" _n _n
}

file write fsimple "*** Diagnósticos ***"                             _n
file write fsimple "varwle"                                           _n
file write fsimple "varlmar, mlag(6)"                                 _n
file write fsimple "varstable, graph dlabel"                          _n
file write fsimple "vargranger"                                       _n _n

file write fsimple "*** IRF ***"                                      _n
file write fsimple `"irf create irf_run, step($irf_step) set("irf_run") replace"' _n
foreach imp of local vars_list {
    foreach resp of local vars_list {
        file write fsimple ///
            "irf graph oirf, impulse(`imp') response(`resp') irf(irf_run)" _n
    }
}
file write fsimple _n

file write fsimple "*** Descomposición de varianza ***"               _n
file write fsimple "irf table fevd, noci"                             _n

file close fsimple
di as result "  → var_simple.do generado"


/* ════════════════════════════════════════════════════════════
   SECCIÓN 11  CERRAR, COPIAR Y RESUMEN FINAL
   ════════════════════════════════════════════════════════════ */

cap copy "var_auto.do" `"$scrpdir/var_completo.do"', replace

log close varlog
translate `"$outdir/log_var.smcl"' `"$outdir/log_var.txt"', replace linesize(120)

di as result ""
di as result "  ══════════════════════════════════════════════"
di as result "  ANÁLISIS COMPLETO  |  $outdir"
di as result "  ══════════════════════════════════════════════"
di as text   "  VAR($p_use)  |  $inicio–$fin"
di as text   "  Cholesky: $chol_order"
di as result "  ── Archivos generados ──────────────────────"
di as text   "  log_var.txt"
di as text   "  diag_estabilidad.png"
di as text   "  irf_[impulso]_[respuesta].png"
di as text   "  fevd_[variable].png"
di as text   "  irf_main.irf"
di as text   "  scripts/var_simple.do"
di as text   "  scripts/var_completo.do"
di as result "  ══════════════════════════════════════════════"

exit, STATA
