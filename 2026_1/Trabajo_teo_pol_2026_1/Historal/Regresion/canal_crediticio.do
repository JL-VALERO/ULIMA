/* ============================================================
   CANAL CREDITICIO DE TRANSMISION DE POLITICA MONETARIA - PERU
   Tasa de referencia BCRP -> Tasa activa bancaria (TAMN)
   Teoria y Politica Monetaria | ULIMA 2026-1
   Stata 17 | Metodologia: ARDL / ECM
   ============================================================

   DATOS REQUERIDOS (hoja "datos" del Excel):
     fecha     -> fecha como tipo fecha en Excel (col A)
     tasa_ref  -> Tasa referencia BCRP %  [serie PD04722MM]
     tamn      -> TAMN - Tasa activa MN % [serie PN07807NM]
     ipc       -> IPC indice base 2009=100 [serie PN01270PM]
     tc        -> Tipo de cambio venta S/$ [serie PD04638PM]

   Periodo: enero 2004 - diciembre 2024 (mensual)
   ============================================================ */


/* ----------------------------------------------------------
   0. CONFIGURACION
---------------------------------------------------------- */

clear all
set more off
set linesize 120

* Instalar paquetes (se salta si ya estan instalados)
cap ssc install ardl,    replace
cap ssc install estout,  replace

* Ruta del proyecto
global ruta   "C:\Users\51950\Documents\ULIMA\2026_1\Teo. Politica\Trabajo grupal\Regresion"
global datos  "$ruta\datos_canal_crediticio.xlsx"
global out    "$ruta\output"
cap mkdir "$out"


/* ----------------------------------------------------------
   1. IMPORTAR Y PREPARAR DATOS
---------------------------------------------------------- */

import excel "$datos", sheet("datos") firstrow clear

* La columna fecha viene como fecha numerica de Excel -> convertir a mensual
gen     mes = mofd(fecha)
format  mes %tm
tsset   mes, monthly
drop    fecha

* Convertir a numerico por si Excel importo alguna columna como texto
destring tasa_ref tamn ipc tc, replace force

* Etiquetas
label var tasa_ref "Tasa referencia BCRP (%)"
label var tamn     "TAMN - Tasa activa MN (%)"
label var ipc      "IPC (base 2009=100)"
label var tc       "Tipo de cambio venta (S/$)"

* Variables derivadas
gen inflacion = (ipc / L.ipc - 1) * 100
gen brecha    = tamn - tasa_ref
gen d_tamn    = D.tamn
gen d_tasa    = D.tasa_ref

label var inflacion "Inflacion mensual (%)"
label var brecha    "Spread TAMN - Tasa ref (pp)"
label var d_tamn    "Delta TAMN"
label var d_tasa    "Delta Tasa referencia"

display "Datos cargados. Observaciones: " _N
list mes tamn tasa_ref in 1/3


/* ----------------------------------------------------------
   2. ESTADISTICAS DESCRIPTIVAS
---------------------------------------------------------- */

display _newline "=== ESTADISTICAS DESCRIPTIVAS ==="
summarize tasa_ref tamn inflacion tc brecha, detail

* Tabla exportada a Word (requiere estout)
estpost summarize tasa_ref tamn brecha inflacion tc
esttab using "$out\descriptivas.rtf", replace ///
    cells("mean(fmt(2) label(Media)) sd(fmt(2) label(Desv.Est.)) min(fmt(2) label(Min)) max(fmt(2) label(Max)) count(fmt(0) label(Obs))") ///
    nomtitle nonumber ///
    title("Estadisticas descriptivas - Canal crediticio Peru 2004-2024")


/* ----------------------------------------------------------
   3. GRAFICOS
---------------------------------------------------------- */

* Grafico 1: Series en el tiempo
twoway (line tamn mes,     lcolor(navy)   lwidth(medthick)) ///
       (line tasa_ref mes, lcolor(orange) lwidth(medthick) lpattern(dash)), ///
    title("TAMN vs Tasa de referencia BCRP")       ///
    ytitle("Tasa de interes (%)") xtitle("")        ///
    legend(label(1 "TAMN") label(2 "Tasa referencia BCRP") rows(1)) ///
    scheme(s1color)
graph export "$out\g1_series.png", replace width(1200)

* Grafico 2: Spread
twoway line brecha mes, lcolor(cranberry) lwidth(medthick) ///
    yline(0, lpattern(dash) lcolor(gs10))     ///
    title("Spread: TAMN - Tasa de referencia BCRP") ///
    ytitle("Puntos porcentuales") xtitle("")   ///
    scheme(s1color)
graph export "$out\g2_spread.png", replace width(1200)

* Grafico 3: Dispersion
twoway (scatter tamn tasa_ref, mcolor(navy) msymbol(o) msize(small)) ///
       (lfit    tamn tasa_ref, lcolor(red)  lwidth(medthick)),        ///
    title("Relacion TAMN vs Tasa referencia")    ///
    ytitle("TAMN (%)") xtitle("Tasa referencia BCRP (%)") ///
    legend(label(1 "Observaciones") label(2 "Ajuste lineal")) ///
    scheme(s1color)
graph export "$out\g3_scatter.png", replace width(900)

display "Graficos guardados en: $out"


/* ----------------------------------------------------------
   4. PRUEBAS DE RAIZ UNITARIA
   H0: la serie tiene raiz unitaria (no estacionaria)
   No rechazar H0 en niveles -> I(1)
   Rechazar H0 en diferencias -> I(0) en diferencias -> confirmado I(1)
---------------------------------------------------------- */

display _newline "=== PRUEBAS DE RAIZ UNITARIA ==="

* ADF - niveles
display _newline "--- ADF en niveles: TAMN ---"
dfuller tamn,     lags(12) regress
dfuller tamn,     lags(12) trend

display _newline "--- ADF en niveles: Tasa referencia ---"
dfuller tasa_ref, lags(12) regress
dfuller tasa_ref, lags(12) trend

* Phillips-Perron - niveles
display _newline "--- Phillips-Perron en niveles ---"
pperron tamn,     lags(6)
pperron tasa_ref, lags(6)

* ADF - primeras diferencias
display _newline "--- ADF en primeras diferencias ---"
dfuller d_tamn, lags(12)
dfuller d_tasa, lags(12)


/* ----------------------------------------------------------
   5. SELECCION DE REZAGOS
   Usar el rezago marcado con * en la columna SBIC
---------------------------------------------------------- */

display _newline "=== SELECCION DE REZAGOS (varsoc) ==="
varsoc tamn tasa_ref, maxlag(12)


/* ----------------------------------------------------------
   6. MODELO ARDL + BOUNDS TEST
   Cointegración de Pesaran, Shin & Smith (2001)
   H0: no existe relacion de largo plazo (no cointegración)
   Rechazar H0 si F-stat > valor critico I(1) al 5%
---------------------------------------------------------- */

display _newline "=== ARDL - BOUNDS TEST ==="

* Estimacion ARDL con seleccion automatica de rezagos por AIC
ardl tamn tasa_ref, maxlags(6 6) aic ec

* Guardar y exportar coeficientes
estimates store ardl_modelo
esttab ardl_modelo using "$out\ardl_resultados.rtf", replace ///
    star(* 0.10 ** 0.05 *** 0.01) b(3) se(3) ///
    title("Modelo ARDL - Canal crediticio Peru")

* Bounds test de cointegración
estat btest


/* ----------------------------------------------------------
   7. MODELO ECM (CORRECCION DE ERROR)
   Separa el efecto de corto y largo plazo
   Lambda (coef de ECT_L1) debe ser < 0 y significativo
---------------------------------------------------------- */

display _newline "=== MODELO ECM ==="

* Paso 1: Relación de largo plazo (OLS en niveles)
reg tamn tasa_ref
predict ect, residuals
label var ect "Error correction term (ECT)"

* Paso 2: Verificar que el ECT sea estacionario (I(0))
display _newline "--- ADF sobre el ECT (debe rechazar H0) ---"
dfuller ect, lags(4) noconstant

* Paso 3: ECM - efecto de corto plazo + velocidad de ajuste
reg d_tamn d_tasa L.d_tasa L.d_tamn L.ect, robust
estimates store ecm_modelo

esttab ecm_modelo using "$out\ecm_resultados.rtf", replace ///
    star(* 0.10 ** 0.05 *** 0.01) b(3) se(3) ///
    title("ECM - Pass-through de corto plazo")  ///
    mtitle("Delta TAMN")

* Paso 4: Pass-through de largo plazo
reg tamn tasa_ref
display _newline "Pass-through de largo plazo = " _b[tasa_ref]
display "(1.0 = traspaso completo | <1 = incompleto | >1 = mas que proporcional)"


/* ----------------------------------------------------------
   8. DIAGNOSTICOS DEL ECM
   H0 en todos: no hay problema
   p > 0.05 = ok
---------------------------------------------------------- */

display _newline "=== DIAGNOSTICOS ==="

quietly reg d_tamn d_tasa L.d_tasa L.d_tamn L.ect

* Autocorrelacion (Breusch-Godfrey)
display "--- Autocorrelacion (Breusch-Godfrey) ---"
estat bgodfrey, lags(1 2 3 6 12)

* Heterocedasticidad (White)
display "--- Heterocedasticidad (White) ---"
estat imtest, white

* Normalidad de residuos
predict res_ecm, residuals
display "--- Normalidad de residuos (Shapiro-Wilk) ---"
swilk res_ecm


/* ----------------------------------------------------------
   9. FUNCION IMPULSO-RESPUESTA (VAR)
   Ordenamiento Cholesky: tasa_ref primero (exogena)
   -> tasa_ref afecta a tamn en t, pero tamn no afecta
      a tasa_ref en el mismo periodo
---------------------------------------------------------- */

display _newline "=== FUNCION IMPULSO-RESPUESTA ==="

* VAR con ordenamiento correcto: tasa_ref -> tamn
var tasa_ref tamn, lags(1/2)

* Crear IRF (carpeta output)
irf create irf1, step(18) set("$out\irf_canal") replace

* Grafico de la IRF
irf graph oirf, impulse(tasa_ref) response(tamn) ///
    yline(0, lpattern(dash) lcolor(gs10))          ///
    title("IRF: Respuesta de TAMN ante shock en Tasa Referencia") ///
    ytitle("Puntos porcentuales") xtitle("Meses")  ///
    scheme(s1color)
graph export "$out\g4_irf.png", replace width(1200)

* Tabla de valores de la IRF
irf table oirf, impulse(tasa_ref) response(tamn) noci


/* ----------------------------------------------------------
   10. RESUMEN FINAL
---------------------------------------------------------- */

display _newline _newline
display "============================================"
display "  RESULTADOS CLAVE - CANAL CREDITICIO PERU"
display "============================================"
display ""
display "  Var. dependiente : TAMN (tasa activa MN)"
display "  Var. independiente: Tasa de referencia BCRP"
display "  Periodo           : ene-2004 a dic-2024"
display "  Metodologia       : ARDL / ECM"
display ""
display "  QUE REVISAR:"
display "  [1] Raiz unitaria -> ADF/PP: NO rechazar H0 en niveles"
display "  [2] Bounds test   -> F-stat > valor critico I(1) al 5%"
display "  [3] ECT (lambda)  -> coef de L.ect: negativo y sig."
display "  [4] Pass-through CP -> coef de d_tasa en el ECM"
display "  [5] Pass-through LP -> coef de tasa_ref en la reg. de niveles"
display "  [6] Diagnostico   -> Breusch-Godfrey y White: p > 0.05"
display ""
display "  Archivos en: $out"
display "============================================"
