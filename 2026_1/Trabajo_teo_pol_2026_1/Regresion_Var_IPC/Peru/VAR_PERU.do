


drop periodo

g time =tm(2004q1)+_n-1
tsset time, m

**********************************************
**GENERAR EL INDICE DE PRESION CAMBIARIA**
**********************************************

* Crear variaciones
gen var_tc = 100*(tc/L.tc - 1)
gen var_rin = 100*(rn/L.rn - 1)
gen var_ti = 100*(tasa_interbancaria/l1.tasa_interbancaria-1)

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

g infl = 100*(ipc/l1.ipc-1)

*crear variacion mensual de credito_privado

g v_cp = 100*(credito_privado/l1.credito_privado-1)

*crear variacion mensual de produccion_nacional
g v_pbi = 100*(produccion_nacional/l1.produccion_nacional-1)


**********************************************
**PRUEBA DE ESTACIONARIEDAD**
**********************************************


pperron vix, reg
pperron ippc, reg
pperron v_cp, reg
pperron v_pbi, reg
pperron infl, reg 

**todas estacionarias**

scc install zandrews
zandrews ippc , maxlags(4)
zandrews v_cp, maxlags(4)
zandrews v_pbi, maxlag(4)
zandrews inf, maxlag(4)

**********************************************
**CORREMOS MODELO VAR**
**********************************************

varsoc vix ippc v_cp v_pbi infl, maxlag(6)

varbasic vix ippc v_cp v_pbi infl, lag(1/1) step(16) nograph

varstable, graph dlabel
varlmar, mlag(12)
varwle
varnorm

vargranger

irf graph irf, impulse(ippc) response(infl)


irf graph irf, impulse(v_cp) response(infl)


irf graph irf, impulse(v_cp) response(v_pbi)


irf graph irf, impulse(v_pbi) response(infl)





