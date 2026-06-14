/* ============================================================
   VAR CON CHOQUES ESTRUCTURALES - CANAL CREDITICIO PERU
   Shocks evaluados:
     D_gfc      : Crisis Financiera Global   (sep-2008 a jun-2009)
     D_pandemic : Pandemia COVID-19           (mar-2020 a dic-2021)
     D_highinfl : Alta inflacion post-COVID   (ene-2022 a dic-2023)
   Teoria y Politica Monetaria | ULIMA 2026-1
   Stata 17
   ============================================================ */

clear all
set more off
set linesize 120

cap ssc install estout, replace

global ruta  "C:\Users\51950\Documents\ULIMA\2026_1\Teo. Politica\Trabajo grupal\Regresion_VAR_Structural"
global datos "$ruta\datos_var_structural.xlsx"
global out   "$ruta\output"
cap mkdir "$out"


/* ----------------------------------------------------------
   1. DATOS Y VARIABLES ESTRUCTURALES
---------------------------------------------------------- */

import excel "$datos", sheet("datos") firstrow clear
gen    mes = mofd(fecha)
format mes %tm
tsset  mes, monthly
drop   fecha
destring tasa_ref tamn ipc tc, replace force

gen inflacion = (ipc / L.ipc - 1) * 100
label var tasa_ref  "Tasa ref. BCRP (%)"
label var tamn      "TAMN (%)"
label var inflacion "Inflacion mensual (%)"
label var tc        "Tipo de cambio (S/$)"

* Dummies de choques estructurales
gen D_gfc      = (mes >= tm(2008m9)  & mes <= tm(2009m6))
gen D_pandemic = (mes >= tm(2020m3)  & mes <= tm(2021m12))
gen D_highinfl = (mes >= tm(2022m1)  & mes <= tm(2023m12))

label var D_gfc      "Crisis Financiera Global 2008-2009"
label var D_pandemic "Pandemia COVID-19 2020-2021"
label var D_highinfl "Alta inflacion post-COVID 2022-2023"

display "Obs: " _N
display "Periodos de quiebre estructural:"
tabulate D_gfc
tabulate D_pandemic
tabulate D_highinfl


/* ----------------------------------------------------------
   2. GRAFICO: SERIES CON SOMBREADO DE CHOQUES
---------------------------------------------------------- */

* Crear variable para zona GFC y pandemia (para sombreado)
twoway (rarea D_gfc D_gfc mes, bcolor(red%15) lwidth(none))          ///
       (rarea D_pandemic D_pandemic mes, bcolor(blue%15) lwidth(none)) ///
       (rarea D_highinfl D_highinfl mes, bcolor(orange%15) lwidth(none)) ///
       (line tasa_ref mes, lcolor(cranberry) lwidth(medthick))        ///
       (line tamn     mes, lcolor(navy)      lwidth(medthick)),       ///
    title("Tasa BCRP y TAMN con choques estructurales")              ///
    ytitle("Tasa (%)") xtitle("")                                     ///
    legend(label(1 "GFC 2008-09") label(2 "COVID 2020-21")           ///
           label(3 "Inf. alta 2022-23") label(4 "Tasa BCRP")        ///
           label(5 "TAMN") rows(2)) scheme(s1color)
graph export "$out\g1_series_choques.png", replace width(1400)

twoway (rarea D_gfc D_gfc mes, bcolor(red%15) lwidth(none))             ///
       (rarea D_pandemic D_pandemic mes, bcolor(blue%15) lwidth(none))   ///
       (rarea D_highinfl D_highinfl mes, bcolor(orange%15) lwidth(none)) ///
       (line inflacion mes, lcolor(forest_green) lwidth(medthick))       ///
       (line tc        mes, lcolor(dkorange)     lwidth(medthick)),      ///
    title("Inflacion y Tipo de cambio con choques estructurales")        ///
    ytitle("") xtitle("")                                                ///
    legend(label(1 "GFC") label(2 "COVID") label(3 "Inf. alta")        ///
           label(4 "Inflacion mensual %") label(5 "Tipo de cambio S/$") rows(2)) ///
    scheme(s1color)
graph export "$out\g2_inf_tc_choques.png", replace width(1400)

display "Graficos de series guardados."


/* ----------------------------------------------------------
   3. PRUEBA DE QUIEBRE ESTRUCTURAL (Chow) EN ECUACIONES OLS
   H0: sin quiebre en la fecha especificada
   Rechazar H0 -> hay cambio estructural
---------------------------------------------------------- */

display _newline "=== TESTS DE QUIEBRE ESTRUCTURAL (CHOW) ==="

* Quiebre en GFC (sep-2008)
display _newline "--- Chow: TAMN ~ tasa_ref  |  Quiebre = sep-2008 ---"
gen post_gfc = (mes >= tm(2008m9))
reg tamn tasa_ref post_gfc c.tasa_ref#c.post_gfc
display "Coef. interaccion = cambio en pass-through durante/post GFC"

* Quiebre en pandemia (mar-2020)
display _newline "--- Chow: TAMN ~ tasa_ref  |  Quiebre = mar-2020 ---"
gen post_covid = (mes >= tm(2020m3))
reg tamn tasa_ref post_covid c.tasa_ref#c.post_covid
display "Coef. interaccion = cambio en pass-through durante/post COVID"

* Diferencias en medias de las variables durante cada shock
display _newline "--- Estadisticas por periodo ---"
tabstat tasa_ref tamn inflacion tc, by(D_gfc)      stat(mean sd) col(stat) nototal
tabstat tasa_ref tamn inflacion tc, by(D_pandemic)  stat(mean sd) col(stat) nototal
tabstat tasa_ref tamn inflacion tc, by(D_highinfl)  stat(mean sd) col(stat) nototal


/* ----------------------------------------------------------
   4. VAR BASELINE (sin dummies - referencia)
   Mismo modelo que Regresion_VAR para comparacion
---------------------------------------------------------- */

display _newline "=== VAR BASELINE (sin dummies de choque) ==="

var tasa_ref tc tamn inflacion, lags(1/2)
estimates store var_base

varstable
display _newline "Estabilidad VAR baseline (arriba)"

irf create irf_base, step(24) set("$out\irf_base") replace

irf table oirf, impulse(tasa_ref) response(tamn) noci
display "(IRF baseline: tasa_ref -> tamn)"


/* ----------------------------------------------------------
   5. VAR CON DUMMIES DE CHOQUE ESTRUCTURAL (modelo principal)
   Los dummies entran como variables exogenas
   Capturan desplazamientos del intercepto durante cada shock
---------------------------------------------------------- */

display _newline "=== VAR CON DUMMIES DE CHOQUE ESTRUCTURAL ==="

var tasa_ref tc tamn inflacion, lags(1/2) ///
    exog(D_gfc D_pandemic D_highinfl)

estimates store var_structural

display _newline "--- Estabilidad del VAR estructural ---"
varstable
varstable, graph
graph export "$out\g3_estabilidad_structural.png", replace width(800)

display _newline "--- Diagnosticos de residuos ---"
varlmar, mlag(12)
varnorm

display _newline "--- Causalidad de Granger (VAR con dummies) ---"
vargranger


/* ----------------------------------------------------------
   6. IRF COMPARACION: baseline vs estructural
   Misma metodologia, diferente especificacion
---------------------------------------------------------- */

display _newline "=== IRF: VAR CON DUMMIES (modelo principal) ==="

irf create irf_struct, step(24) set("$out\irf_structural") replace

* IRF tasa_ref -> tamn
irf graph oirf, impulse(tasa_ref) response(tamn) ///
    yline(0, lpattern(dash) lcolor(gs10)) ///
    title("IRF: TAMN ante shock en Tasa Ref. (VAR con dummies estructurales)") ///
    ytitle("Puntos porcentuales") xtitle("Meses") scheme(s1color)
graph export "$out\g4_irf_tamn_structural.png", replace width(1200)

* IRF tasa_ref -> inflacion
irf graph oirf, impulse(tasa_ref) response(inflacion) ///
    yline(0, lpattern(dash) lcolor(gs10)) ///
    title("IRF: Inflacion ante shock en Tasa Ref. (VAR con dummies)") ///
    ytitle("Puntos porcentuales") xtitle("Meses") scheme(s1color)
graph export "$out\g5_irf_inf_structural.png", replace width(1200)

* IRF tasa_ref -> tc
irf graph oirf, impulse(tasa_ref) response(tc) ///
    yline(0, lpattern(dash) lcolor(gs10)) ///
    title("IRF: Tipo de Cambio ante shock en Tasa Ref. (VAR con dummies)") ///
    ytitle("S/$") xtitle("Meses") scheme(s1color)
graph export "$out\g6_irf_tc_structural.png", replace width(1200)

* IRF combinada 3 canales
irf graph oirf, impulse(tasa_ref) response(tc tamn inflacion) ///
    yline(0, lpattern(dash) lcolor(gs10)) ///
    title("Canales de transmision (VAR con dummies de choques)") ///
    xtitle("Meses") scheme(s1color)
graph export "$out\g7_irf_combinada_structural.png", replace width(1400)

* Tablas numericas IRF estructural
display _newline "=== TABLA IRF: tasa_ref -> tamn (VAR estructural) ==="
irf table oirf, impulse(tasa_ref) response(tamn) noci

display _newline "=== TABLA IRF: tasa_ref -> inflacion (VAR estructural) ==="
irf table oirf, impulse(tasa_ref) response(inflacion) noci

display _newline "=== TABLA IRF: tasa_ref -> tc (VAR estructural) ==="
irf table oirf, impulse(tasa_ref) response(tc) noci


/* ----------------------------------------------------------
   7. FEVD CON DUMMIES ESTRUCTURALES
---------------------------------------------------------- */

display _newline "=== FEVD (VAR con dummies) ==="

display "--- FEVD de TAMN ---"
irf table fevd, impulse(tasa_ref tc tamn inflacion) response(tamn) noci

irf graph fevd, response(tamn) ///
    title("FEVD de TAMN (VAR con dummies de choques)") ///
    xtitle("Horizonte (meses)") scheme(s1color)
graph export "$out\g8_fevd_tamn_structural.png", replace width(1200)

display "--- FEVD de Inflacion ---"
irf table fevd, impulse(tasa_ref tc tamn inflacion) response(inflacion) noci

irf graph fevd, response(inflacion) ///
    title("FEVD de Inflacion (VAR con dummies de choques)") ///
    xtitle("Horizonte (meses)") scheme(s1color)
graph export "$out\g9_fevd_inf_structural.png", replace width(1200)


/* ----------------------------------------------------------
   8. ANALISIS POR SUBPERIODOS
   VAR estimado en 3 periodos:
     Pre-GFC     : 2004m1 - 2008m8
     Pre-COVID   : 2009m7 - 2020m2
     Post-COVID  : 2020m3 - 2025m12
   -> Permite ver si el mecanismo de transmision cambio
---------------------------------------------------------- */

display _newline "=== VAR POR SUBPERIODOS ==="

* --- Subperiodo 1: Pre-GFC ---
display _newline "--- SUB-PERIODO 1: Pre-GFC (2004m1 - 2008m8) ---"
preserve
keep if mes >= tm(2004m1) & mes <= tm(2008m8)
display "Obs en sub-periodo 1: " _N
cap var tasa_ref tc tamn inflacion, lags(1/2)
if _rc == 0 {
    varstable
    irf create irf_pregfc, step(18) set("$out\irf_pregfc") replace
    irf table oirf, impulse(tasa_ref) response(tamn) noci
    display "(IRF pre-GFC: tasa_ref -> tamn)"
}
else {
    display "AVISO: muestra insuficiente para VAR en sub-periodo 1"
}
restore

* --- Subperiodo 2: Pre-COVID ---
display _newline "--- SUB-PERIODO 2: Pre-COVID (2009m7 - 2020m2) ---"
preserve
keep if mes >= tm(2009m7) & mes <= tm(2020m2)
display "Obs en sub-periodo 2: " _N
var tasa_ref tc tamn inflacion, lags(1/2)
varstable
irf create irf_precovid, step(18) set("$out\irf_precovid") replace
irf table oirf, impulse(tasa_ref) response(tamn) noci
display "(IRF pre-COVID: tasa_ref -> tamn)"
restore

* --- Subperiodo 3: Post-COVID ---
display _newline "--- SUB-PERIODO 3: Post-COVID (2020m3 - 2025m12) ---"
preserve
keep if mes >= tm(2020m3)
display "Obs en sub-periodo 3: " _N
cap var tasa_ref tc tamn inflacion, lags(1/2)
if _rc == 0 {
    varstable
    irf create irf_postcovid, step(18) set("$out\irf_postcovid") replace
    irf table oirf, impulse(tasa_ref) response(tamn) noci
    display "(IRF post-COVID: tasa_ref -> tamn)"
}
else {
    display "AVISO: muestra insuficiente o error en sub-periodo 3"
}
restore


/* ----------------------------------------------------------
   9. GRAFICO COMPARATIVO IRF: BASELINE vs ESTRUCTURAL
   Usando irf set para superponer ambas IRFs
---------------------------------------------------------- */

display _newline "=== COMPARACION IRF: BASELINE vs CON DUMMIES ==="

* Restaurar IRF del VAR con dummies y comparar con baseline
irf set "$out\irf_structural"
irf table oirf, impulse(tasa_ref) response(tamn) noci

display _newline "Tabla comparativa (baseline vs estructural):"
display "Baseline  (sin dummies): ver irf_base"
display "Estructural (con GFC+COVID+highinfl dummies): ver irf_structural"


/* ----------------------------------------------------------
   10. RESUMEN DE RESULTADOS
---------------------------------------------------------- */

display _newline _newline
display "======================================================"
display "  VAR CON CHOQUES ESTRUCTURALES - RESULTADOS CLAVE"
display "======================================================"
display "  Shocks modelados:"
display "    D_gfc      : Crisis Financiera Global  sep-2008/jun-2009"
display "    D_pandemic : COVID-19                  mar-2020/dic-2021"
display "    D_highinfl : Alta inflacion post-COVID  ene-2022/dic-2023"
display ""
display "  QUE LEER:"
display "  [1] Chow tests  -> coef. de interaccion  (cambio en pass-through)"
display "  [2] Estabilidad -> eigenvalues < 1"
display "  [3] LM test     -> p > 0.05 = sin autocorrelacion"
display "  [4] Granger     -> tasa_ref->tamn y tasa_ref->inflacion"
display "  [5] IRF (mes 7) -> pico de respuesta de TAMN (baseline vs estructural)"
display "  [6] FEVD (mes12)-> % varianza TAMN explicado por tasa_ref"
display "  [7] Subperiodos -> cambio en IRF antes/despues de cada crisis"
display ""
display "  Archivos en: $out"
display "======================================================"
