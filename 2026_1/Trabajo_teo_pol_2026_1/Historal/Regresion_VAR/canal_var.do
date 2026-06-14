/* ============================================================
   CANAL CREDITICIO - MODELO VAR
   Tasa de referencia BCRP -> Tipo de cambio -> TAMN -> Inflacion
   Teoria y Politica Monetaria | ULIMA 2026-1
   Stata 17 | Metodologia: VAR con identificacion Cholesky

   Variables (orden Cholesky - de mas a menos exogena):
     1. tasa_ref  -> instrumento de politica monetaria (exogena)
     2. tc        -> tipo de cambio (reacciona rapido a politica)
     3. tamn      -> tasa activa MN (canal crediticio)
     4. inflacion -> inflacion mensual % (endogena, I(0))

   Datos: enero 2004 - diciembre 2025 (mensual)
   ============================================================ */


/* ----------------------------------------------------------
   0. CONFIGURACION
---------------------------------------------------------- */

clear all
set more off
set linesize 120

cap ssc install estout, replace

global ruta  "C:\Users\51950\Documents\ULIMA\2026_1\Teo. Politica\Trabajo grupal\Regresion_VAR"
global datos "$ruta\datos_var.xlsx"
global out   "$ruta\output"
cap mkdir "$out"


/* ----------------------------------------------------------
   1. IMPORTAR Y PREPARAR DATOS
---------------------------------------------------------- */

import excel "$datos", sheet("datos") firstrow clear

gen    mes = mofd(fecha)
format mes %tm
tsset  mes, monthly
drop   fecha

destring tasa_ref tamn ipc tc, replace force

label var tasa_ref "Tasa referencia BCRP (%)"
label var tamn     "TAMN - Tasa activa MN (%)"
label var ipc      "IPC (base 2009=100)"
label var tc       "Tipo de cambio venta (S/$)"

* Tasa de inflacion mensual (I(0) -> mas estable que log IPC en niveles)
gen inflacion = (ipc / L.ipc - 1) * 100
label var inflacion "Inflacion mensual (%)"

* Log del tipo de cambio
gen ltc = log(tc)
label var ltc "Log Tipo de cambio"

* Primeras diferencias (para pruebas de raiz unitaria)
gen d_tasa = D.tasa_ref
gen d_tamn = D.tamn
gen d_ltc  = D.ltc

display "Datos cargados. Obs: " _N
list mes tasa_ref tamn inflacion tc in 1/3


/* ----------------------------------------------------------
   2. ESTADISTICAS DESCRIPTIVAS
---------------------------------------------------------- */

display _newline "=== ESTADISTICAS DESCRIPTIVAS ==="
summarize tasa_ref tamn inflacion tc, detail

estpost summarize tasa_ref tamn inflacion tc
esttab using "$out\descriptivas_var.rtf", replace ///
    cells("mean(fmt(2) label(Media)) sd(fmt(2) label(Desv.Est.)) min(fmt(2) label(Min)) max(fmt(2) label(Max)) count(fmt(0) label(Obs))") ///
    nomtitle nonumber ///
    title("Estadisticas descriptivas - VAR Canal crediticio Peru 2004-2025")


/* ----------------------------------------------------------
   3. GRAFICOS PRELIMINARES
---------------------------------------------------------- */

* G1: Las 4 series del VAR
twoway (line tasa_ref mes,  lcolor(cranberry)    lwidth(medthick)) ///
       (line tamn     mes,  lcolor(navy)          lwidth(medthick)) ///
       (line tc       mes,  lcolor(forest_green)  lwidth(medthick) lpattern(dash)) ///
       (line inflacion mes, lcolor(orange)         lwidth(medthick) lpattern(shortdash)), ///
    title("Variables del VAR - Peru 2004-2025") ///
    ytitle("Tasas (%)  /  S/$") ///
    legend(label(1 "Tasa ref. BCRP") label(2 "TAMN") ///
           label(3 "Tipo de cambio") label(4 "Inflacion mensual") rows(2)) ///
    scheme(s1color)
graph export "$out\g1_series_var.png", replace width(1400)

* G2: Correlacion cruzada tasa_ref / tamn
xcorr tasa_ref tamn, lags(12) ///
    title("Correlacion cruzada: Tasa ref. vs TAMN") ///
    scheme(s1color)
graph export "$out\g2_xcorr.png", replace width(900)

display "Graficos preliminares guardados."


/* ----------------------------------------------------------
   4. PRUEBAS DE RAIZ UNITARIA
   H0: tiene raiz unitaria (I(1))
   tasa_ref, tamn, tc -> se espera I(1)
   inflacion          -> se espera I(0)
---------------------------------------------------------- */

display _newline "=== PRUEBAS DE RAIZ UNITARIA ==="

* Variables en niveles
foreach v in tasa_ref tamn ltc {
    display _newline "--- ADF: `v' (niveles, con constante) ---"
    dfuller `v', lags(12)
    display _newline "--- ADF: `v' (niveles, con tendencia) ---"
    dfuller `v', lags(12) trend
    display _newline "--- PP: `v' (niveles) ---"
    pperron `v', lags(6)
}

* Inflacion (deberia ser I(0))
display _newline "--- ADF: inflacion (deberia ser estacionaria) ---"
dfuller inflacion, lags(12)

display _newline "=== ADF EN PRIMERAS DIFERENCIAS (confirmar I(1)) ==="
foreach v in d_tasa d_tamn d_ltc {
    display _newline "--- ADF: `v' ---"
    dfuller `v', lags(12)
}


/* ----------------------------------------------------------
   5. PRUEBA DE COINTEGRACION (Johansen)
   H0: no hay cointegracion (rango = 0)
   Si hay cointegracion -> podria preferirse VECM
   Sims (1980): VAR en niveles es valido incluso con I(1)
---------------------------------------------------------- */

display _newline "=== PRUEBA DE COINTEGRACION (JOHANSEN) ==="

varsoc tasa_ref tc tamn inflacion, maxlag(12)

display _newline "--- Johansen (2 rezagos, constante restringida) ---"
vecrank tasa_ref tc tamn inflacion, lags(2) trend(constant) max levela


/* ----------------------------------------------------------
   6. SELECCION DE REZAGOS DEL VAR
   Criterios: AIC, HQIC, SBIC
   Se usa el rezago con * en SBIC como referencia
---------------------------------------------------------- */

display _newline "=== SELECCION DE REZAGOS VAR ==="
varsoc tasa_ref tc tamn inflacion, maxlag(12)


/* ----------------------------------------------------------
   7. ESTIMACION DEL VAR
   Cholesky: tasa_ref -> tc -> tamn -> inflacion
   Interpretacion: shock de politica monetaria -> efectos
   sobre tipo de cambio, tasa activa y precios
---------------------------------------------------------- */

display _newline "=== ESTIMACION DEL VAR(2) ==="

var tasa_ref tc tamn inflacion, lags(1/2)

estimates store var_modelo

esttab var_modelo using "$out\var_resultados.rtf", replace ///
    star(* 0.10 ** 0.05 *** 0.01) b(3) se(3) ///
    title("VAR(2) - Canal crediticio Peru 2004-2025") ///
    mtitle("tasa_ref" "tc" "tamn" "inflacion")


/* ----------------------------------------------------------
   8. ESTABILIDAD DEL VAR
   Todos los modulos de eigenvalues deben ser < 1
---------------------------------------------------------- */

display _newline "=== VERIFICACION DE ESTABILIDAD ==="

varstable

varstable, graph
graph export "$out\g3_estabilidad.png", replace width(800)


/* ----------------------------------------------------------
   9. DIAGNOSTICOS DE RESIDUOS
   H0: sin problema en todos los tests
   p > 0.05 = ok
---------------------------------------------------------- */

display _newline "=== DIAGNOSTICOS DE RESIDUOS ==="

display "--- Autocorrelacion (LM test de Portmanteau) ---"
varlmar, mlag(12)

display "--- Normalidad multivariada (Jarque-Bera) ---"
varnorm

display "--- Heterocedasticidad (test ARCH multivariado) ---"
quietly var tasa_ref tc tamn inflacion, lags(1/2)
foreach eq in tasa_ref tc tamn inflacion {
    quietly predict res_`eq', residuals equation(`eq')
    display "  Skewness/Kurtosis residuos ecuacion `eq':"
    sktest res_`eq'
    drop res_`eq'
}


/* ----------------------------------------------------------
   10. CAUSALIDAD DE GRANGER
   H0: X no Granger-causa a Y
   p < 0.05 -> X ayuda a predecir Y
---------------------------------------------------------- */

display _newline "=== CAUSALIDAD DE GRANGER ==="

vargranger

display ""
display "RELACIONES CLAVE A VERIFICAR:"
display "  tasa_ref -> tamn     : canal crediticio directo"
display "  tasa_ref -> inflacion: mecanismo de transmision a precios"
display "  tasa_ref -> tc       : canal cambiario"
display "  tc -> tamn           : efecto tipo de cambio sobre credito"


/* ----------------------------------------------------------
   11. FUNCIONES IMPULSO-RESPUESTA (OIRF Cholesky)
   Un shock de 1 desv.est. en tasa_ref
   Horizonte: 24 meses
---------------------------------------------------------- */

display _newline "=== FUNCIONES IMPULSO-RESPUESTA ==="

irf create irf_var, step(24) set("$out\irf_var") replace

* IRF principal: tasa_ref -> tamn (canal crediticio)
irf graph oirf, impulse(tasa_ref) response(tamn) ///
    yline(0, lpattern(dash) lcolor(gs10)) ///
    title("IRF: Respuesta de TAMN ante shock en Tasa de Referencia") ///
    ytitle("Puntos porcentuales") xtitle("Meses") ///
    scheme(s1color)
graph export "$out\g4_irf_tasa_tamn.png", replace width(1200)

* IRF: tasa_ref -> inflacion (canal de precios)
irf graph oirf, impulse(tasa_ref) response(inflacion) ///
    yline(0, lpattern(dash) lcolor(gs10)) ///
    title("IRF: Respuesta de Inflacion ante shock en Tasa de Referencia") ///
    ytitle("Puntos porcentuales") xtitle("Meses") ///
    scheme(s1color)
graph export "$out\g5_irf_tasa_inf.png", replace width(1200)

* IRF: tasa_ref -> tc (canal cambiario)
irf graph oirf, impulse(tasa_ref) response(tc) ///
    yline(0, lpattern(dash) lcolor(gs10)) ///
    title("IRF: Respuesta del Tipo de Cambio ante shock en Tasa Ref.") ///
    ytitle("S/$") xtitle("Meses") ///
    scheme(s1color)
graph export "$out\g6_irf_tasa_tc.png", replace width(1200)

* IRF combinada (los 3 canales)
irf graph oirf, impulse(tasa_ref) response(tc tamn inflacion) ///
    yline(0, lpattern(dash) lcolor(gs10)) ///
    title("Canales de transmision: shock en Tasa Ref. BCRP") ///
    xtitle("Meses") scheme(s1color)
graph export "$out\g7_irf_combinada.png", replace width(1400)

* Tablas numericas
display _newline "Tabla OIRF: tasa_ref -> tamn"
irf table oirf, impulse(tasa_ref) response(tamn) noci

display _newline "Tabla OIRF: tasa_ref -> inflacion"
irf table oirf, impulse(tasa_ref) response(inflacion) noci

display _newline "Tabla OIRF: tasa_ref -> tc"
irf table oirf, impulse(tasa_ref) response(tc) noci


/* ----------------------------------------------------------
   12. DESCOMPOSICION DE VARIANZA DEL ERROR (FEVD)
   Que % de la varianza de TAMN e inflacion se debe a
   shocks propios vs shocks de la tasa de referencia?
---------------------------------------------------------- */

display _newline "=== DESCOMPOSICION DE VARIANZA (FEVD) ==="

display "--- FEVD de TAMN ---"
irf table fevd, impulse(tasa_ref tc tamn inflacion) response(tamn) noci

irf graph fevd, response(tamn) ///
    title("FEVD de TAMN: contribucion de cada variable") ///
    xtitle("Horizonte (meses)") scheme(s1color)
graph export "$out\g8_fevd_tamn.png", replace width(1200)

display "--- FEVD de Inflacion ---"
irf table fevd, impulse(tasa_ref tc tamn inflacion) response(inflacion) noci

irf graph fevd, response(inflacion) ///
    title("FEVD de Inflacion: contribucion de cada variable") ///
    xtitle("Horizonte (meses)") scheme(s1color)
graph export "$out\g9_fevd_inf.png", replace width(1200)


/* ----------------------------------------------------------
   13. IRF ACUMULADA (CIRF)
   Pass-through total acumulado de un shock de politica
---------------------------------------------------------- */

display _newline "=== IRF ACUMULADA (CIRF) ==="

irf graph coirf, impulse(tasa_ref) response(tamn) ///
    yline(0, lpattern(dash) lcolor(gs10)) ///
    title("CIRF Acumulada: efecto total de Tasa Ref. sobre TAMN") ///
    ytitle("Puntos porcentuales acumulados") xtitle("Meses") ///
    scheme(s1color)
graph export "$out\g10_cirf_tamn.png", replace width(1200)

display _newline "Tabla CIRF: tasa_ref -> tamn"
irf table coirf, impulse(tasa_ref) response(tamn) noci


/* ----------------------------------------------------------
   14. RESUMEN FINAL
---------------------------------------------------------- */

display _newline _newline
display "============================================"
display "  RESULTADOS CLAVE - VAR CANAL CREDITICIO"
display "============================================"
display "  Var. endogenas : tasa_ref, tc, tamn, inflacion"
display "  Identificacion : Cholesky (tasa_ref exogena)"
display "  Periodo        : muestra comun (IPC disponible)"
display "  Metodologia    : VAR(2)"
display ""
display "  QUE REVISAR:"
display "  [1] Estabilidad    -> todos eigenvalues < 1 en modulo"
display "  [2] LM test        -> p > 0.05 = sin autocorrelacion"
display "  [3] Granger        -> p < 0.05 en tasa_ref -> tamn"
display "  [4] IRF (mes 6-12) -> respuesta pico de TAMN"
display "  [5] FEVD (mes 12)  -> % varianza TAMN explicado por tasa_ref"
display "  [6] CIRF           -> pass-through acumulado total"
display ""
display "  Archivos guardados en: $out"
display "============================================"
