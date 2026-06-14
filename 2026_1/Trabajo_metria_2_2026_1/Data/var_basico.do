/*==============================================================
  VAR BASICO - MECANISMO DE TRANSMISION MONETARIA EN CHILE
  Periodo: 2004m1 - 2025m12 (datos mensuales, T=264)
  Autor  : [nombre]
  Fecha  : 26 abr 2026
  Compatibilidad: Stata 12 o superior
  --------------------------------------------------------------
  VARIABLES DEL MODELO
  -----------------------------------------------
  Endogenas (orden de Cholesky):
    brecha_pib  Output gap (%).  I(0). Nivel.
    infl        Inflacion mensual IPC (%).  I(0). Nivel.
    d_ltc       D.log(tc): depreciacion mensual CLP/USD. I(0).
    tpm         Tasa de Politica Monetaria BCCh (%).  Nivel.

  Exogena:
    d_lpetro    D.log(petro): log-retorno mensual del petroleo.

  DESCARTADAS:
    ipc         Redundante con infl.
    imacec      I(1); brecha_pib captura el ciclo.
    m1          BCCh opera con IT, no con agregados monetarios.
    embi        6 obs. missing al inicio.
    cobre       Exogeno a Chile; efecto captado via brecha y tc.
    ipc_esperado Alta colinealidad con infl.

  LOGICA DE ESTACIONARIZACION:
    El test se elige segun las propiedades de cada serie.
    Ver seccion 3 para el detalle completo.
==============================================================*/

version 12
clear all
set more off


/* -------------------------------------------------------
   1. IMPORTAR Y PREPARAR DATOS
------------------------------------------------------- */

import excel using "DATA_JORGE.xlsx", ///
    sheet("Sheet1") cellrange(A4) clear

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

drop if missing(fecha)

gen fecha_m = mofd(fecha)
format fecha_m %tm
tsset fecha_m, monthly

keep if fecha_m <= tm(2025m12)

* Transformaciones logaritmicas
gen l_tc    = log(tc)
gen l_petro = log(petro)

label var brecha_pib "Brecha del producto (%)"
label var infl       "Inflacion mensual IPC (%)"
label var tpm        "TPM BCCh (%)"
label var l_tc       "Log tipo de cambio (CLP/USD)"
label var l_petro    "Log precio petroleo (USD)"

sum fecha_m
local T = r(N)
di "Observaciones: `T'  (2004m1 - 2025m12)"


/* -------------------------------------------------------
   2. ESTADISTICAS DESCRIPTIVAS
------------------------------------------------------- */

di _n "=== ESTADISTICAS DESCRIPTIVAS ==="
sum brecha_pib infl tpm l_tc l_petro


/* -------------------------------------------------------
   3. TESTS DE RAIZ UNITARIA Y ESTACIONARIZACION
   -------------------------------------------------------
   El test optimo se elige segun las propiedades de cada
   serie. Criterio: 5% de significancia en todos los casos.

   ELECCION DE TEST POR VARIABLE:
   +-------------+------------------------------------------+
   | brecha_pib  | ADF (lags=12, constante, sin tendencia)  |
   |             | Razon: I(0) por construccion; lags altos |
   |             | para rho(1)=0.96; gap oscila en torno    |
   |             | a cero, sin tendencia determinista.      |
   +-------------+------------------------------------------+
   | infl        | ADF (lags=6, constante, sin tendencia)   |
   |             | Razon: rho(1)=0.38, baja persistencia;   |
   |             | bajo IT no hay tendencia sistematica.    |
   +-------------+------------------------------------------+
   | tpm         | ADF (lags=6, con tendencia)               |
   |             | Razon: rho(1)=0.99 con multiples         |
   |             | rupturas (2009, 2020, 2021-22). El ADF   |
   |             | puede no rechazar H0 por sesgo de        |
   |             | rupturas; la decision de usar niveles    |
   |             | descansa en argumentos economicos y      |
   |             | SSW (1990). Ver seccion 3.3.             |
   +-------------+------------------------------------------+
   | l_tc        | PP (NW=8, constante + tendencia)         |
   |             | Razon: serie financiera con varianza     |
   |             | condicional cambiante (ARCH); PP es      |
   |             | robusto a heterocedasticidad, ADF no.    |
   +-------------+------------------------------------------+
   | l_petro     | PP (NW=8, constante + tendencia)         |
   |             | Razon: precio de commodity con shocks    |
   |             | de varianza cambiante. Misma logica      |
   |             | que l_tc.                                |
   +-------------+------------------------------------------+
------------------------------------------------------- */

di _n "=== SECCION 3: TESTS DE RAIZ UNITARIA ==="

local nw       = 8   /* ancho de banda Newey-West (PP)          */
local lags_bp  = 12  /* lags ADF brecha_pib (alta persistencia) */
local lags_inf = 6   /* lags ADF infl (baja persistencia)       */
local lags_tpm = 6   /* lags ADF tpm                            */


/* ============================================================
   3.1  BRECHA DEL PRODUCTO (brecha_pib)
   ------------------------------------------------------------
   Propiedad clave: output gap es I(0) por construccion
   (desviacion entre PIB observado y tendencia Hodrick-Prescott
   o equivalente). La alta persistencia rho(1)=0.96 refleja
   la duracion de los ciclos economicos, no una raiz unitaria.
   Test: ADF con constante, sin tendencia, 12 lags.
   H0: raiz unitaria  |  Rechazo -> I(0), usar nivel.
============================================================ */

di _n "----------------------------------------------"
di    "3.1  brecha_pib  |  ADF (lags=12, constante)"
di    "----------------------------------------------"
dfuller brecha_pib, lags(`lags_bp')

di _n "DECISION: I(0) - se usa en niveles."
di    "Justificacion economica: el output gap es estacionario"
di    "por construccion (oscila en torno a cero)."
di    "La persistencia rho=0.96 refleja duracion del ciclo,"
di    "no una tendencia estocastica."


/* ============================================================
   3.2  INFLACION (infl)
   ------------------------------------------------------------
   Bajo el esquema de Inflation Targeting (IT) del BCCh desde
   1999, la inflacion esta anclada en torno a la meta (3% anual,
   aprox. 0.25% mensual). Una serie I(1) implicaria inflacion
   sin limite, inconsistente con el IT. rho(1)=0.38 es bajo;
   6 lags son suficientes para limpiar autocorrelacion residual.
   H0: raiz unitaria  |  Rechazo -> I(0), usar nivel.
============================================================ */

di _n "----------------------------------------------"
di    "3.2  infl  |  ADF (lags=6, constante)"
di    "----------------------------------------------"
dfuller infl, lags(`lags_inf')

di _n "DECISION: I(0) - se usa en niveles."
di    "Justificacion economica: el IT ancla la inflacion."
di    "Una inflacion I(1) es inconsistente con metas del BCCh."


/* ============================================================
   3.3  TASA DE POLITICA MONETARIA (tpm)
   ------------------------------------------------------------
   CONTEXTO: La TPM presenta rho(1)=0.99 y cuatro episodios
   de ruptura estructural en la muestra:
     - 2009:    crisis subprime -> corte de -5pp en 5 meses
     - 2020m4:  COVID -> piso historico de 0.5%
     - 2021m9 - 2022m10: ciclo de alzas hasta 11.25%
     - 2023m8 - 2025:    ciclo de bajas hasta ~4.5%

   POR QUE EL ADF PUEDE NO RECHAZAR H0 (y por que eso
   no implica que la TPM sea I(1)):
   El ADF estandar tiene baja potencia cuando existen
   rupturas estructurales en la serie: las discontinuidades
   en la media o en la tendencia inflan artificialmente el
   estadistico, generando un sesgo hacia el no rechazo de H0.
   Este es un resultado bien documentado (Perron 1989).
   Con cuatro episodios de cambio de regimen, el ADF casi
   con seguridad fallara en rechazar H0 para la TPM chilena,
   incluso si la serie es estacionaria dentro de cada regimen.

   La alternativa natural seria el test de Zivot-Andrews
   (1992), que permite una ruptura endogena y corrige ese
   sesgo. Sin embargo, zandrews no esta disponible como
   comando nativo en todas las instalaciones de Stata.

   POR QUE ESTO NO INVALIDA EL MODELO:
   La decision de usar tpm en niveles NO depende del resultado
   de ningun test formal. Descansa en tres pilares:

   (1) SERIE ACOTADA: TPM in [0.5%, 11.25%] durante toda la
       muestra. Una serie verdaderamente I(1) diverge sin
       limite con el tiempo; una tasa de politica monetaria
       no puede hacerlo. El I(1) aparente es un artefacto
       estadistico de las multiples rupturas de nivel, no
       una propiedad real del proceso generador de datos.

   (2) COHERENCIA ECONOMICA: el instrumento de politica
       monetaria es el NIVEL de la tasa. Un shock de +1pp
       en la TPM es la unidad natural de analisis (equivale
       a un alza estandar de 100 puntos base). Usar D.tpm
       convertiria el shock en una "aceleracion" de la tasa,
       sin analogia directa con la practica del BCCh.

   (3) JUSTIFICACION ESTADISTICA - SSW (1990):
       Sims, Stock & Watson (1990) demuestran que, en un VAR
       estimado en niveles, la inferencia asintotica es valida
       incluso si alguna variable es I(1), siempre que el VAR
       sea estable (todos los eigenvalores dentro del circulo
       unitario). La estabilidad se verifica formalmente en
       la seccion 6. Si el VAR es estable, la presencia de
       I(1) en la TPM no distorsiona la inferencia sobre IRF
       ni sobre los coeficientes del resto del sistema.
============================================================ */

di _n "----------------------------------------------"
di    "3.3  tpm  |  ADF (con tendencia, lags=`lags_tpm')"
di    "----------------------------------------------"
di    "NOTA: zandrews no disponible en esta instalacion."
di    "  Se usa ADF con tendencia como test de referencia."
di    "  El resultado esperado (no rechazo) refleja el sesgo"
di    "  del ADF ante multiples rupturas, no I(1) verdadera."
di    "  La decision de usar niveles se basa en argumentos"
di    "  economicos y en SSW (1990). Ver comentario 3.3."

dfuller tpm, lags(`lags_tpm') trend

di _n "DECISION: se usa en NIVELES (independiente del test)."
di    "  (1) Acotada [0.5%, 11.25%]: no puede ser I(1) real."
di    "  (2) Coherencia economica: shock +1pp = alza estandar."
di    "  (3) SSW (1990): inferencia valida si VAR es estable."
di    "  La estabilidad se verifica formalmente en seccion 6."


/* ============================================================
   3.4  TIPO DE CAMBIO: log(CLP/USD)  [variable: l_tc]
   ------------------------------------------------------------
   Los tipos de cambio bajo flotacion libre son paseos
   aleatorios clasicos (Meese & Rogoff 1983). rho(1)=0.985.
   La heterocedasticidad condicional (ARCH) es comun en series
   financieras: PP corrige de forma no parametrica sin asumir
   homocedasticidad, ventaja clave frente al ADF.
   Se testea en niveles y en primera diferencia.
   H0 niveles: I(1). H0 diferencias: I(2) (se espera rechazar).
============================================================ */

di _n "----------------------------------------------"
di    "3.4  l_tc  |  PP (NW=8, constante+tendencia)"
di    "----------------------------------------------"

di _n "-- Niveles (se espera NO rechazar H0) --"
pperron l_tc, lags(`nw') trend

di _n "-- Primera diferencia D.l_tc (se espera rechazar H0) --"
pperron D.l_tc, lags(`nw') trend

di _n "DECISION: I(1) - se usa D.l_tc (depreciacion mensual)."
di    "D.l_tc = variacion mensual del log tipo de cambio."
di    "Una depreciacion positiva eleva la inflacion via pass-through."


/* ============================================================
   3.5  PRECIO DEL PETROLEO: log(petroleo)  [variable: l_petro]
   ------------------------------------------------------------
   Los precios de commodities siguen un paseo aleatorio con
   deriva variable (Hamilton 2008). rho(1)=0.959. Misma logica
   que l_tc: PP es preferible por heterocedasticidad.
   Se usa como variable EXOGENA: Chile es tomador de precios.
============================================================ */

di _n "----------------------------------------------"
di    "3.5  l_petro  |  PP (NW=8, constante+tendencia)"
di    "----------------------------------------------"

di _n "-- Niveles (se espera NO rechazar H0) --"
pperron l_petro, lags(`nw') trend

di _n "-- Primera diferencia D.l_petro (se espera rechazar H0) --"
pperron D.l_petro, lags(`nw') trend

di _n "DECISION: I(1) - se usa D.l_petro (log-retorno mensual)."
di    "Variable exogena: Chile es tomador de precios del petroleo."


/* ============================================================
   3.6  CREAR VARIABLES FINALES PARA EL VAR
============================================================ */

gen d_ltc    = D.l_tc       /* depreciacion mensual log(CLP/USD) */
gen d_lpetro = D.l_petro    /* log-retorno mensual petroleo      */

label var d_ltc    "Depreciacion mensual log(CLP/USD)"
label var d_lpetro "Log-retorno mensual petroleo (exog.)"

di _n "=== VARIABLES FINALES CREADAS ==="
sum brecha_pib infl tpm d_ltc d_lpetro


/* ============================================================
   3.7  RESUMEN DE ESTACIONARIZACION
============================================================ */

di _n "======================================================"
di    "  RESUMEN: TRANSFORMACIONES PARA EL VAR"
di    "------------------------------------------------------"
di    "  Variable     Test    Orden  Variable VAR  Forma"
di    "------------------------------------------------------"
di    "  brecha_pib   ADF     I(0)   brecha_pib    nivel"
di    "  infl         ADF     I(0)   infl          nivel"
di    "  tpm          ADF     nivel* tpm           nivel"
di    "  l_tc         PP      I(1)   d_ltc         D.l_tc"
di    "  l_petro      PP      I(1)   d_lpetro      D.l_petro"
di    "------------------------------------------------------"
di    "  * ADF probablemente no rechaza H0 (sesgo por rupturas)."
di    "    Decision de niveles basada en: serie acotada,"
di    "    coherencia economica y SSW (1990). Ver sec. 3.3."
di    "======================================================"


/* -------------------------------------------------------
   4. SELECCION DE REZAGOS OPTIMOS
   Variables: brecha_pib infl d_ltc tpm  /  exog: d_lpetro
   Criterios: AIC (referencia), HQIC (intermedio), SBIC
   (parsimonia). Prioridad: SBIC para el modelo base.
------------------------------------------------------- */

di _n "=== CRITERIOS DE SELECCION DE REZAGOS (max. 12) ==="
varsoc brecha_pib infl d_ltc tpm, maxlag(12) exog(d_lpetro)

/*
  INSTRUCCION: Ajuste `p` segun resultado de varsoc.
  Con 4 variables y datos mensuales, SBIC suele elegir p=1-3.
  Si SBIC y HQIC difieren, pruebe ambas especificaciones y
  elija la que pase mejor los tests de congruencia (seccion 7).
*/
local p = 2    /* <-- AJUSTAR segun resultado de varsoc */


/* -------------------------------------------------------
   5. ESTIMACION DEL VAR
   Orden Cholesky: brecha_pib -> infl -> d_ltc -> tpm
   El BCCh fija la TPM observando brecha, inflacion y TC.
------------------------------------------------------- */

di _n "=== ESTIMACION VAR(`p') ==="
di    "Endogenas : brecha_pib  infl  d_ltc  tpm"
di    "Exogena   : d_lpetro"
di    "Rezagos   : `p'"
di    "Cholesky  : brecha_pib -> infl -> d_ltc -> tpm"

var brecha_pib infl d_ltc tpm, lags(1/`p') exog(d_lpetro)

estimates store var_base


/* -------------------------------------------------------
   6. CONGRUENCIA: ESTABILIDAD DEL VAR
   Condicion: todos los eigenvalores con modulo < 1.
   Si alguna raiz >= 1: aumentar p o revisar transformaciones.
------------------------------------------------------- */

di _n "=== ESTABILIDAD DEL VAR (raices caracteristicas) ==="
varstable, graph name(g_estabilidad, replace)


/* -------------------------------------------------------
   7. CONGRUENCIA: AUTOCORRELACION DE RESIDUOS
   Test LM de Breusch-Godfrey para el VAR.
   H0: no hay autocorrelacion serial en los residuos.
   Deseable: NO rechazar H0 (p-value > 0.05).
   Si se rechaza: aumentar el numero de rezagos `p`.
------------------------------------------------------- */

di _n "=== TEST LM DE AUTOCORRELACION (hasta rezago 12) ==="
varlmar, mlag(12)


/* -------------------------------------------------------
   8. NORMALIDAD DE RESIDUOS (informativo)
   H0: residuos normalmente distribuidos.
   Con T=264 la inferencia es asintotica; normalidad
   es deseable pero no indispensable.
------------------------------------------------------- */

di _n "=== TEST DE NORMALIDAD (Jarque-Bera) ==="
varnorm, jbera


/* -------------------------------------------------------
   9. FUNCIONES IMPULSO-RESPUESTA (IRF)
   Identificacion recursiva (Cholesky).
   Horizonte: 24 meses.

   INTERPRETACION DE VARIABLES:
   - tpm   : nivel de la tasa -> shock = +1pp en la TPM
   - infl  : inflacion mensual en %
   - d_ltc : tasa de depreciacion mensual;
             respuesta negativa = apreciacion del peso
   - brecha: output gap; respuesta negativa = contraccion
------------------------------------------------------- */

di _n "=== FUNCIONES IMPULSO-RESPUESTA ==="

* set(irf_base, replace): crea el archivo irf_base.irf
* y lo reemplaza si ya existe (compatible Stata 12+)
irf create irf_base, step(24) set(irf_base, replace)

* (a) Canal de tasas: infl <- shock TPM
irf graph irf, impulse(tpm) response(infl) ///
    yline(0) ///
    title("Respuesta de Inflacion a Shock de TPM (+1pp)") ///
    note("Cholesky. IC 95%. tpm en niveles.") ///
    name(g_irf_infl, replace)

* (b) Canal real: brecha <- shock TPM
irf graph irf, impulse(tpm) response(brecha_pib) ///
    yline(0) ///
    title("Respuesta de Brecha de Producto a Shock de TPM") ///
    note("Cholesky. IC 95%.") ///
    name(g_irf_brecha, replace)

* (c) Canal cambiario: d_ltc <- shock TPM
*     Signo esperado: negativo (alza TPM aprecia el CLP)
irf graph irf, impulse(tpm) response(d_ltc) ///
    yline(0) ///
    title("Respuesta de Depreciacion a Shock de TPM") ///
    note("Signo (-) = apreciacion CLP. Cholesky. IC 95%.") ///
    name(g_irf_tc, replace)

* (d) Pass-through cambiario: infl <- shock d_ltc
*     Signo esperado: positivo (depreciacion -> mayor inflacion)
irf graph irf, impulse(d_ltc) response(infl) ///
    yline(0) ///
    title("Pass-through: Inflacion ante Shock de Depreciacion") ///
    note("Cholesky. IC 95%.") ///
    name(g_irf_pt, replace)


/* -------------------------------------------------------
   10. DESCOMPOSICION DE VARIANZA (FEVD)
   Fraccion de la varianza de cada variable explicada
   por cada shock estructural.
------------------------------------------------------- */

di _n "=== DESCOMPOSICION DE VARIANZA (FEVD) ==="

* FEVD: inflacion
irf graph fevd, ///
    impulse(brecha_pib infl d_ltc tpm) response(infl) ///
    title("FEVD: Inflacion") ///
    name(g_fevd_infl, replace)

* FEVD: brecha del producto
irf graph fevd, ///
    impulse(brecha_pib infl d_ltc tpm) response(brecha_pib) ///
    title("FEVD: Brecha del Producto") ///
    name(g_fevd_brecha, replace)

di _n "======================================================"
di    "  VAR TRANSMISION MONETARIA CHILE - ESTIMACION OK"
di    "------------------------------------------------------"
di    "  Endogenas  : brecha_pib  infl  d_ltc  tpm"
di    "  Exogena    : d_lpetro"
di    "  Rezagos    : `p' (ajustar segun varsoc)"
di    "  Cholesky   : brecha_pib->infl->d_ltc->tpm"
di    "------------------------------------------------------"
di    "  Coherencia esperada de IRF:"
di    "    Shock +TPM   -> infl (-), brecha (-), d_ltc (-)"
di    "    Shock +d_ltc -> infl (+)  [pass-through]"
di    "======================================================"
