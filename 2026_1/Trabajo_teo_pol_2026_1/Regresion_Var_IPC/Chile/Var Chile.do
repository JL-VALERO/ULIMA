
clear all


drop fecha
drop periodo

g time =tm(2004q1)+_n-1
tsset time, m

**********************************************
**GENERAR EL INDICE DE PRESION CAMBIARIA**
**********************************************

* Crear variaciones
gen chi_vtc = 100*(tc_clp_usd/L.tc_clp_usd - 1)
gen chi_vrin = 100*(reservas_netas_musd/L.reservas_netas_musd - 1)
gen chi_vti = 100*(tasa_interbancaria_tpm/l1.tasa_interbancaria_tpm-1)

* Normalizar cada componente
egen z_tc = std(chi_vtc)
egen z_rin = std(chi_vrin)
egen z_ti = std(chi_vti)

* Construir índice de presión cambiaria
gen chi_ippc = z_tc - z_rin + z_ti

*Sin normalizar

g chi_ippc2 = chi_vtc +chi_vti - chi_vrin


* Revisar resultados
summarize var_tc var_rin var_ti z_tc z_rin z_ti ippc

tsline chi_ippc

*Crear inflacion mensual

g chi_infl = 100*(ipc_chile/l1.ipc_chile-1)

g chi_linfl = 100*(log(ipc_chile)-log(l1.ipc_chile))

*crear variacion mensual de credito_privado

g chi_vcp = 100*(credito_privado_mclp/l1.credito_privado_mclp-1)
g chi_lcp = 100*(log(credito_privado_mclp)-log(l1.credito_privado_mclp))

*crear variacion mensual de produccion_nacional
g chi_vpbi = 100*(imacec_pbi/l1.imacec_pbi-1)

g chi_lpbi = 100*(log(imacec_pbi)-log(l1.imacec_pbi))

*crear variacion del cobre

g v_cobre = 100*(cobre/l1.cobre-1)

**********************************************
**PRUEBA DE ESTACIONARIEDAD**
**********************************************

pperron chi_lpbi, reg




pperron v_cobre, reg
pperron chi_ippc, reg 
pperron chi_vcp, reg 
pperron chi_vpbi, reg 
pperron chi_infl, reg

**todas estacionarias**

scc install zandrews
zandrews ippc , maxlags(4)
zandrews v_cp, maxlags(4)
zandrews v_pbi, maxlag(4)
zandrews inf, maxlag(4)

**********************************************
**CORREMOS MODELO VAR**
**********************************************

varsoc v_cobre chi_ippc chi_vcp chi_vpbi chi_infl

varbasic v_cobre chi_ippc chi_vcp chi_vpbi chi_infl, lag(1/1) step(16) nograph

varsoc v_cobre chi_ippc chi_vcp chi_vpbi chi_infl if !inrange(time, tm(2008m1), tm(2009m12)) & !inrange(time, tm(2020m1), tm(2021m12)), maxlag(6)

varbasic v_cobre chi_ippc chi_vcp chi_vpbi chi_infl if !inrange(time, tm(2008m1), tm(2009m12)) & !inrange(time, tm(2020m1), tm(2021m12)), lag(1/1) step(16) nograph

varstable, graph dlabel
varlmar, mlag(12)
varwle
varnorm

vargranger

irf graph irf, impulse(chi_ippc) response(chi_infl)


irf graph irf, impulse(chi_infl) response(chi_vcp)


irf graph irf, impulse(v_cobre) response(chi_vpbi)


irf graph irf, impulse(chi_vpbi) response(chi_infl)





