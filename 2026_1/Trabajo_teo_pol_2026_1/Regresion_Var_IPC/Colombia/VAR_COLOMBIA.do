drop Fecha
drop Periodo

g time =tm(2004q1)+_n-1
tsset time, m

**********************************************
**GENERAR EL INDICE DE PRESION CAMBIARIA**
**********************************************

* Crear variaciones
gen var_tc = 100*(tc/L.tc - 1)
gen var_rin = 100*(rin/L.rin - 1)
gen var_ti = 100*(ti/l1.ti-1)

* Normalizar cada componente
egen z_tc = std(var_tc)
egen z_rin = std(var_rin)
egen z_ti = std(var_ti)

* Construir índice de presión cambiaria
gen ippc = z_tc - z_rin + z_ti

* Revisar resultados
summarize var_tc var_rin var_ti z_tc z_rin z_ti ippc

tsline ippc

*Crear inflacion mensual

g col_infl = 100*(ipc/l1.ipc-1)

*crear variacion mensual de credito_privado

g col_vcp = 100*(cp/l1.cp-1)

*crear variacion mensual de produccion_nacional
g col_vpbi = 100*(pib/l1.pib-1)

*crear variable variacon mensual brent "exogena"

g v_brent = 100*(BrentUSDbbl/l1.BrentUSDbbl-1)

**********************************************
**PRUEBA DE ESTACIONARIEDAD**
**********************************************


pperron brent, reg
pperron ippc, reg
pperron col_vcp, reg
pperron col_vpbi, reg
pperron col_infl, reg 


**todas estacionarias**

scc install zandrews
zandrews ippc , maxlags(4)
zandrews v_cp, maxlags(4)
zandrews v_pbi, maxlag(4)
zandrews inf, maxlag(4)

**********************************************
**CORREMOS MODELO VAR**
**********************************************

varsoc ippc col_vcp col_vpbi col_infl, maxlag(6)

varbasic  ippc col_vcp col_vpbi col_infl, lag(1/3) step(16) nograph

varstable, graph dlabel
varlmar, mlag(12)
varwle
varnorm

vargranger

irf graph irf, impulse(ippc) response(col_infl)


irf graph irf, impulse(col_vpbi) response(col_infl)


irf graph irf, impulse(col_vcp) response(col_infl)


irf graph irf, impulse(v_pbi) response(infl)





