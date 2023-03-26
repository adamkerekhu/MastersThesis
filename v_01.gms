

*---------------------------------------------------------------------------
* Input from Excel File
*---------------------------------------------------------------------------
$onEps
$ontext
$call =xls2gms   r="v1"!A2:A8671          i=input.xlsx    o=set_tt.inc
$call =xls2gms   r="v1"!A3:A8671          i=input.xlsx    o=set_tt2.inc
$call =xls2gms   r="v1"!A2:B8671          i=input.xlsx    o=par_co2.inc
$call =xls2gms   r="v1"!K2:L8671          i=input.xlsx    o=par_eprice_low.inc
$call =xls2gms   r="v1"!M2:N8671          i=input.xlsx    o=par_eprice_high.inc
$call =xls2gms   r="v1"!O2:P8671          i=input.xlsx    o=par_eprice_volatile.inc
$call =xls2gms   r="v1"!C2:D8671          i=input.xlsx    o=par_servertemp.inc
$call =xls2gms   r="v1"!E2:F8671          i=input.xlsx    o=par_ambienttemp.inc
$call =xls2gms   r="v1"!G2:H8671          i=input.xlsx    o=par_serverload.inc
$call =xls2gms   r="v1"!I2:J8671          i=input.xlsx    o=par_humidity.inc
$call =xls2gms   r="v1"!Q2:R8671          i=input.xlsx    o=par_gprice.inc
$offtext
set t 
/
$include set_tt.inc
/;
set t2(t)
/
$include set_tt2.inc
/;
parameter E_el(t)
/
$include par_co2.inc
/;
parameter C_el(t)
/
$include par_eprice_high.inc
/;
parameter C_gas(t)
/
$include par_gprice.inc
/;
parameter S_temp(t)
/
$include par_servertemp.inc
/;
parameter A_temp(t)
/
$include par_ambienttemp.inc
/;
parameter L_server(t)
/
$include par_serverload.inc
/;
parameter A_hum(t)
/
$include par_humidity.inc
/;

*---------------------------------------------------------------------------
* General
*---------------------------------------------------------------------------
Variables
Em(t)
Em_total
C(t)
C_total
Z
Qall
;

*---------------------------------------------------------------------------
* Technologies
*---------------------------------------------------------------------------
Scalars
x_boiler              -                                                 /0/
x_turbine             -                                                 /0/
x_AC                  -                                                 /1/
x_EC                  -                                                 /1/
x_AFC                 -                                                 /0/
x_WFC                 -                                                 /0/
x_S                   -                                                 /1/
;

*---------------------------------------------------------------------------
* Electricity grid
*---------------------------------------------------------------------------
Positive variables    
S_el(t)               kWh       
Q_grid_to_EC(t)       kWh
Q_grid_to_IT(t)       kWh
Q_grid_to_AFC(t)      kWh
Q_grid_to_WFC(t)      kWh
Em_el(t)              g
Em_el_total           tonnes
tC_el(t)
C_el_total
;

*---------------------------------------------------------------------------
* Gas grid
*---------------------------------------------------------------------------
Scalars
H_gas                 kWh\m3                                            /10.8/
;

Positive variables
S_gas(t)              m3       
V_gas_to_turbine(t)   m3
V_gas_to_boiler(t)    m3
tC_gas(t)
C_gas_total
;

*---------------------------------------------------------------------------
* Boiler
*---------------------------------------------------------------------------
Scalars
P_boiler              kW                                                /20000/
n_boiler              \%                                                /75/
E_boiler              g\kWh                                             /215/
C_boiler              EUR                                               //
l_boiler              years                                             //
;

Positive variables
Q_boiler_to_AC(t)     kWh
EM_boiler(t)          g
Em_boiler_total       tonnes
;

*---------------------------------------------------------------------------
* Gas turbine
*---------------------------------------------------------------------------
Scalars
P_turbine             kW                                                /20000/
n_turbine             \%                                                /90/
E_turbine             g\kWh                                             /300/
C_turbine             EUR
l_turbine             years
el_rate_turbine       \%                                                /30/
;

Positive variables
Q_turbine_to_AC(t)    kWh
Q_turbine_to_EC(t)    kWh
Q_turbine_loss(t)     kWh
EM_turbine(t)          g
Em_turbine_total      tonnes
;

*---------------------------------------------------------------------------
* Absorption chiller
*---------------------------------------------------------------------------
Scalars
P_AC                  kW                                                /20000/
COP_AC                -                                                 /1.4/
;

Positive Variables
Q_AC_to_S(t)          kW
;

*---------------------------------------------------------------------------
* Electric chiller
*---------------------------------------------------------------------------
Scalars
P_EC                  kW                                                /20000/
COP_EC                -                                                 /3.5/
;

Positive variables
Q_EC_to_S(t)          kW
;

*---------------------------------------------------------------------------
* Airside direct free cooling system
*---------------------------------------------------------------------------
Scalars
Rho                   kg\m3                                             /1.225/
C_p_air               KJ\kg                                             /1.012/
m_AFC                 m3\h                                              /4000000/
n_AFC                 %                                                 /90/
n_ht                                                                    /60/
P_AFC                 kW                                                /3240/
;

Positive variables
m_AFC_to_S(t)         m3
Q_AFC_to_S(t)         kW
;

*---------------------------------------------------------------------------
* Waterside cooling tower free cooling system
*---------------------------------------------------------------------------
Scalars
P_WFC                                                                   /3240/
n_WFC                 \%                                                /90/
m_WFC                 m3\h                                              /4000000/
C_p_water             KJ\kg                                             /4.186/                                                          
;

Positive variables
Q_WFC_to_S(t)         kW
m_WFC_to_S(t)
;

Parameters
Wb_temp(t)           
CT_range(t)
CT_approach(t)
n_CT(t)     
;

Wb_temp(t) =  A_temp(t) * arctan(0.151977 * sqrt(A_hum(t) + 8.313659)) + arctan(A_temp(t) + A_hum(t) ) - arctan(A_hum(t) - 1.676331) + 0.00391838 *sqrt(A_hum(t)*A_hum(t)*A_hum(t)) * arctan(0.023101 * A_hum(t)) - 4.686035
;
CT_range(t)  = S_temp(t) - A_temp(t)
;
CT_approach(t) = A_temp(t) - Wb_temp(t)
;
CT_range(t) $ ( S_temp(t) - A_temp(t) < 0 ) = 0
;
n_CT(t) =  CT_range(t) / ( CT_approach(t) + CT_range(t) )
;

*---------------------------------------------------------------------------
* Cooling storage
*---------------------------------------------------------------------------
Scalars
Q_S_max               kWh                                               /2000000/
SOC_S_max             \%                                                /98/
SOC_S_min             \%                                                /10/
SOC_S_initial         \%                                                /50/
CH_S_max              kW                                                /20000/
DCH_S_max             kW                                                /20000/
l_S                   \%\h                                               
;

Positive variables
Q_S(t)                kWh
Q_S_to_IT(t)          kWh
SOC_S(t)              \%
;

*---------------------------------------------------------------------------
* IT equipment
*---------------------------------------------------------------------------
Scalars
D_all                 kWh                                               /10000/
D_fixed               kWh
;

Positive variables
Server_load(t)        \%
D_IT(t)               kWh
D_cooling(t)          kWh
;

*---------------------------------------------------------------------------
* Equations
*---------------------------------------------------------------------------
Equations

*---------------------------------------------------------------------------
* Objective funcion
*---------------------------------------------------------------------------
obj                    objective function

*---------------------------------------------------------------------------
* Power transformation
*---------------------------------------------------------------------------
max_boiler(t)
max_turbine(t)
max_AC(t)
max_EC(t)
max_AFC(t)
max_WFC(t)

*---------------------------------------------------------------------------
* Node equilibriums
*---------------------------------------------------------------------------
eq_boiler(t)           boiler
eq_turbine_el(t)       gas turbine
eq_turbine_h(t)        gas turbine
eq_AC(t)               absorption chiller
eq_EC(t)               electric chiller
eq_S(t)                cooling storage
eq_IT(t)               demand
eq_AFC(t)              cooling storage
eq_WFC(t)              demand

*---------------------------------------------------------------------------
* Free cooling calculations
*---------------------------------------------------------------------------
qafc(t)
qwfc(t)

*---------------------------------------------------------------------------
* Cooling storage constraints
*---------------------------------------------------------------------------
s_max(t)               maximum state of charge
s_min(t)               min state of charge
s_initial
s_closure
s_ch(t)
s_dch(t)
*s_fix(t)   

*---------------------------------------------------------------------------
* CO2 related calculations
*---------------------------------------------------------------------------
co2_grid(t)            co2 from electricity usage
co2_boiler(t)          co2 from boiler
co2_turbine(t)         co2 from gas turbine
co2_el_all             total co2 from grid
co2_boiler_all         total co2 from boiler
co2_turbine_all        total co2 from turbine
co2_t(t)               co2 at given t slot
co2_all                total co2 from all

*---------------------------------------------------------------------------
* Total cost summaries
*---------------------------------------------------------------------------
cost_gas(t)
cost_el(t)
c_t(t)
cost_all
c_el_all
c_gas_all

Q_all
;

*---------------------------------------------------------------------------
* Objective funcion
*---------------------------------------------------------------------------

obj               ..  Z =e= sum(t, Em_el(t) + Em_boiler(t) + Em_turbine(t) ) / 1000000
*obj               ..  Z =e= sum(t, tC_gas(t) + tC_el(t))
;


*---------------------------------------------------------------------------
* Power maximums
*---------------------------------------------------------------------------
max_boiler(t)      ..  P_boiler * x_boiler * n_boiler /100 =g= Q_boiler_to_AC(t)
;

max_turbine(t)     ..  P_turbine * x_turbine * n_turbine /100 =g= Q_turbine_to_EC(t) + Q_turbine_to_AC(t) 
;

max_AC(t)          ..  P_AC * x_AC * COP_AC =g= Q_AC_to_S(t)
;
   
max_EC(t)          ..  P_EC * x_EC * COP_EC =g= Q_EC_to_S(t)
;

max_AFC(t)         ..  m_AFC * x_AFC =g= m_AFC_to_S(t) 
;

max_WFC(t)         ..  m_WFC * x_WFC =g= m_WFC_to_S(t) 
;
    
*---------------------------------------------------------------------------
* Node equilibriums
*---------------------------------------------------------------------------
eq_boiler(t)       ..  H_gas * V_gas_to_boiler(t) =e= Q_boiler_to_AC(t)
;

eq_turbine_el(t)   ..  H_gas * V_gas_to_turbine(t) * el_rate_turbine / 100 =e= Q_turbine_to_EC(t)
;

eq_turbine_h(t)    ..  H_gas * V_gas_to_turbine(t) * (1 - el_rate_turbine) / 100 =e= Q_turbine_to_AC(t) + Q_turbine_loss(t)
;
          
eq_AC(t)           ..  (Q_turbine_to_AC(t) + Q_boiler_to_AC(t)) * COP_AC =e= Q_AC_to_S(t)
;
 
eq_EC(t)           ..  (Q_grid_to_EC(t) + Q_turbine_to_EC(t)) * COP_EC =e= Q_EC_to_S(t)
;

eq_AFC(T)          ..  Q_grid_to_AFC(t) =e= m_AFC_to_S(t) / m_AFC * P_AFC * n_AFC / 100
;

eq_WFC(T)          ..  Q_grid_to_WFC(t) =e= m_WFC_to_S(t) / m_WFC * P_WFC * n_WFC / 100
;
           
eq_S(t)$t2(t)      ..  Q_S(t) =e= Q_S(t-1) + Q_AC_to_S(t) + Q_EC_to_S(t) + Q_WFC_to_S(t) + Q_AFC_to_S(t) - Q_S_to_IT(t)
;
              
eq_IT(t)           ..  D_all * L_server(t) =e= Q_S_to_IT(t)
;

*---------------------------------------------------------------------------
* Free cooling calculations
*---------------------------------------------------------------------------
qafc(t)            ..  Q_AFC_to_S(t) =e= m_AFC_to_S(t) * C_p_air * CT_range(t) * Rho / 3600 * n_ht /100
;

qwfc(t)            ..  Q_WFC_to_S(t) =e= m_WFC_to_S(t) * C_p_air * CT_range(t) * Rho / 3600 * n_CT(t)
;

*---------------------------------------------------------------------------
* Cooling storage constraints
*---------------------------------------------------------------------------    
s_max(t)           ..  Q_S(t) =l= SOC_S_max / 100 * x_S * Q_S_max
;
             
s_min(t)           ..  Q_S(t) =g= SOC_S_min / 100 * Q_S_max
;

s_initial          ..  Q_S('1') =e= SOC_S_initial /100 * Q_S_max
;

s_closure          ..  Q_S('8670') =e= SOC_S_initial /100 * Q_S_max
;

s_ch(t)            ..  (Q_AC_to_S(t) + Q_EC_to_S(t) + Q_WFC_to_S(t) + Q_AFC_to_S(t)) * x_S =l= CH_S_max
;

s_dch(t)           ..  Q_S_to_IT(t) * x_S =l= DCH_S_max
;

*s_fix(t)           ..  Q_S_to_IT(t) =e= Q_AC_to_S(t) + Q_EC_to_S(t) + Q_WFC_to_S(t) + Q_AFC_to_S(t)
*;

*---------------------------------------------------------------------------
* CO2 related calculations
*---------------------------------------------------------------------------
co2_grid(t)        ..  E_el(t) * ( Q_grid_to_EC(t) + Q_grid_to_AFC(t) + Q_grid_to_WFC(t) ) =e= Em_el(t)
;

co2_boiler(t)      ..  E_boiler * Q_boiler_to_AC(t) =e= Em_boiler(t)
;
       
co2_turbine(t)     ..  E_turbine * V_gas_to_turbine(t) =e= Em_turbine(t)
;

co2_el_all         ..  sum( t, Em_el(t) )/1000000 =e= Em_el_total
;

co2_boiler_all     ..  sum( t, Em_boiler(t) )/1000000 =e= Em_boiler_total
;

co2_turbine_all    ..  sum( t, Em_turbine(t) )/1000000 =e= Em_turbine_total
;

co2_t(t)           ..  Em_el(t) + Em_boiler(t) + Em_turbine(t) =e= Em(t)
;
       
co2_all            ..  Em_el_total + Em_boiler_total + Em_turbine_total =e= Em_total
;    

*---------------------------------------------------------------------------
* Total cost summaries
*---------------------------------------------------------------------------

cost_gas(t)        ..  H_gas * ( V_gas_to_boiler(t) + V_gas_to_turbine(t) ) * C_gas(t)/1000000 =e= tC_gas(t)
;
cost_el(t)         ..  ( Q_grid_to_EC(t) + Q_grid_to_AFC(t) + Q_grid_to_WFC(t) ) * C_el(t)/1000000 =e= tC_el(t)
;
c_t(t)             ..  (tC_gas(t) + tC_el(t)) =e= C(t)
;
cost_all           ..  sum(t, C(t)) =e= C_total
;
c_el_all           ..  sum(t, tC_el(t))  =e= C_el_total
;
c_gas_all          ..  sum(t, tC_gas(t)) =e= C_gas_total
;
Q_all              ..  sum(t, Q_grid_to_EC(t) + V_gas_to_boiler(t)*10.8)/1000000 =e= Qall
;


*---------------------------------------------------------------------------
* Solving the model
*---------------------------------------------------------------------------
Model v_01 /all/
;
Solve v_01 using MIP minimizing Z
;

*---------------------------------------------------------------------------
* Display results
*---------------------------------------------------------------------------
display Z.l
;
display C_el_total.l
;
display C_gas_total.l
;
display C_total.l
;
display Em_total.l
;
display Qall.l
;
*$ontext

Parameters

Q_grid_to_EC_out(t)
Q_grid_to_AFC_out(t)
Q_grid_to_WFC_out(t)
V_gas_to_boiler_out(t)
V_gas_to_turbine_out(t)
Q_boiler_to_AC_out(t)
Q_turbine_to_AC_out(t)
Q_turbine_to_EC_out(t)
Q_AC_to_S_out(t)
Q_EC_to_S_out(t)
Q_WFC_to_S_out(t)
Q_S_to_IT_out(t)
Q_AFC_to_S_out(t)
Em_boiler_out(t)
Em_turbine_out(t)
Em_el_out(t)
Em_out(t)
SOC_S_out(t)
Q_S_out(t)
Wb_temp_out(t)
n_CT_out(t)
C_gas_out(t)
C_el_out(t)
C_out(t)
;

Q_grid_to_EC_out(t) = Q_grid_to_EC.L(t) +EPS
;
Q_grid_to_AFC_out(t) = Q_grid_to_AFC.L(t) +EPS
;
Q_grid_to_WFC_out(t) = Q_grid_to_WFC.L(t) +EPS
;
V_gas_to_boiler_out(t) = V_gas_to_boiler.L(t) +EPS
;
V_gas_to_turbine_out(t) = V_gas_to_turbine.L(t) +EPS
;
Q_boiler_to_AC_out(t) = Q_boiler_to_AC.L(t) +EPS
;
Q_turbine_to_AC_out(t) = Q_turbine_to_AC.L(t) +EPS
;
Q_turbine_to_EC_out(t) = Q_turbine_to_EC.L(t) +EPS
;
Q_AC_to_S_out(t) = Q_AC_to_S.L(t) +EPS
;
Q_EC_to_S_out(t) = Q_EC_to_S.L(t) +EPS
;
Q_WFC_to_S_out(t) = Q_WFC_to_S.L(t) +EPS
;
Q_S_to_IT_out(t) = Q_S_to_IT.L(t) +EPS
;
Q_AFC_to_S_out(t) = Q_AFC_to_S.L(t) +EPS
;
Em_boiler_out(t) = Em_boiler.L(t) +EPS
;
Em_turbine_out(t) = Em_turbine.L(t) +EPS
;
Em_el_out(t) = Em_el.L(t) +EPS
;
Q_S_out(t) = Q_S.L(t) +EPS
;
n_CT_out(t) = n_CT(t) + EPS
;
C_gas_out(t) = tc_gas.L(t) + EPS
;
C_el_out(t) = tc_el.L(t) + EPS
;
C_out(t) = c.L(t) + EPS
;
Wb_temp_out(t) = Wb_temp(t) + EPS
;

display Z.l
;

*---------------------------------------------------------------------------
* Export time series data
*---------------------------------------------------------------------------

execute_unload "resultsv1.gdx"      Q_grid_to_EC_out
                                    V_gas_to_boiler_out
                                    V_gas_to_turbine_out
                                    Q_boiler_to_AC_out
                                    Q_turbine_to_AC_out
                                    Q_turbine_to_EC_out
                                    Q_AC_to_S_out
                                    Q_EC_to_S_out
                                    Q_WFC_to_S_out
                                    Q_S_to_IT_out
                                    Q_AFC_to_S_out
                                    Em_boiler_out
                                    Em_turbine_out
                                    Em_el_out
                                    Em.L
                                    SOC_S.L
                                    Q_S_out
                                    E_el
                                    n_CT_out
                                    Wb_temp_out
                                    C_gas_out
                                    C_el_out
                                    C_out
                                    Q_grid_to_AFC_out
                                    Q_grid_to_WFC_out
     
$ontext                               
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Q_grid_to_EC_out rng=a3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=V_gas_to_boiler_out rng=c3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=V_gas_to_turbine_out rng=e3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Q_boiler_to_AC_out rng=g3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Q_turbine_to_AC_out rng=i3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Q_turbine_to_EC_out rng=k3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Q_AC_to_S_out rng=m3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Q_EC_to_S_out rng=o3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Q_WFC_to_S_out rng=q3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Q_S_to_IT_out rng=s3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Q_AFC_to_S_out rng=u3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Em_boiler_out rng=w3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Em_turbine_out rng=y3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Em_el_out rng=aa3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx var=Em.L rng=ac3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=E_el rng=ai3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=n_CT_out rng=ak3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Wb_temp_out rng=am3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=C_gas_out rng=ao3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=C_el_out rng=aq3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=C_out rng=as3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx var=SOC_S.L rng=ae3 rdim=1 cdim=0 '
;

execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Q_S_out rng=ag3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Q_grid_to_AFC_out rng=au3 rdim=1 cdim=0 '
;
execute 'gdxxrw.exe  resultsv1.gdx o=results_opt.xlsx par=Q_grid_to_WFC_out rng=aw3 rdim=1 cdim=0 '
;

$offtext

$offEps
