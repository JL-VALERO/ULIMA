/*=============================================
  VAR Simple
  Especificación: VAR(2) | 2004m1–2021m12
  Cholesky: tasa_ref tc tamn pct_ipc
  Generado por: var_completo.do
=============================================*/

*** Configurar serie de tiempo ***
gen time = tm(2004m1) + _n - 1
format time %tm
tsset time, monthly

*** Prueba de estacionariedad (niveles) ***
pperron tasa_ref
pperron tamn
pperron ipc
pperron tc

*** Crear variables ***
* tasa_ref: se usa en niveles
* tamn: se usa en niveles
gen pct_ipc = (ipc/L.ipc-1)*100
* tc: se usa en niveles

*** Prueba de estacionariedad (vars. transformadas) ***
pperron tasa_ref
pperron tamn
pperron pct_ipc
pperron tc

*** Selección de rezagos ***
varsoc tasa_ref tc tamn pct_ipc, maxlag(6)

*** Estimación VAR ***
var tasa_ref tc tamn pct_ipc, lags(1/2)

*** Diagnósticos ***
varwle
varlmar, mlag(6)
varstable, graph dlabel
vargranger

*** IRF ***
irf create irf_run, step(24) set("irf_run") replace
irf graph oirf, impulse(tasa_ref) response(tasa_ref) irf(irf_run)
irf graph oirf, impulse(tasa_ref) response(tc) irf(irf_run)
irf graph oirf, impulse(tasa_ref) response(tamn) irf(irf_run)
irf graph oirf, impulse(tasa_ref) response(pct_ipc) irf(irf_run)
irf graph oirf, impulse(tc) response(tasa_ref) irf(irf_run)
irf graph oirf, impulse(tc) response(tc) irf(irf_run)
irf graph oirf, impulse(tc) response(tamn) irf(irf_run)
irf graph oirf, impulse(tc) response(pct_ipc) irf(irf_run)
irf graph oirf, impulse(tamn) response(tasa_ref) irf(irf_run)
irf graph oirf, impulse(tamn) response(tc) irf(irf_run)
irf graph oirf, impulse(tamn) response(tamn) irf(irf_run)
irf graph oirf, impulse(tamn) response(pct_ipc) irf(irf_run)
irf graph oirf, impulse(pct_ipc) response(tasa_ref) irf(irf_run)
irf graph oirf, impulse(pct_ipc) response(tc) irf(irf_run)
irf graph oirf, impulse(pct_ipc) response(tamn) irf(irf_run)
irf graph oirf, impulse(pct_ipc) response(pct_ipc) irf(irf_run)

*** Descomposición de varianza ***
irf table fevd, noci
