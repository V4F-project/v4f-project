# Input parameters for the power system model

# This file contains all the input parameters for the  power system model
# Time series input data is imported from an Excel file, and other parameters values are inputted manually in this file.
# Manually inputted parameters are marked with "MI" (manual input)

# Import pandas for the purpose of using data from Excel
import pandas
# import math for calculations
import math
# Time series input data  for the model consists of outside air temperature, spot prices for electricity and fuels,
# and thermal demand. System temperatures for HPs and TS unit can also be inputted as time series
df = pandas.read_excel('input_data.xlsx') # Specify the name of input data Excel file

# Take note of the units of the parameters. They are specified above each manually inputted parameter

# ======================================================================================================================
# Non-unit specific parameters
# ======================================================================================================================

"MI" # Number of timesteps in simulation [-]. Note: there must be enough data in the Excel file to cover all timesteps
# in order for the model to work
t_sim = 672
# Time vector, set of T = t_sim timesteps
T = range(1,t_sim+1)
"MI" # Length of simulation timestep (delta t) [h]
dt = 0.25

"MI" # Outside air temperature [C]
T_0 = df['T_0 [C]'][:t_sim] # Specify name of column from the input data Excel file

"MI" # Price for heat [€/MWh]
p_Q = 80

"MI" # Electricity spot prices [€/MWh]
p_E = df['Electricity price [€/MWh]'][:t_sim] # Specify name of column from the input data Excel

"MI" # Transmission costs for buying electricity [€/MWh]
p_E_tr_buy = 15
"MI" # Transmission costs for selling electricity [€/MWh]
p_E_tr_sell = 0.8

"MI" # Thermal demand [MW]
Q_d = df['Thermal demand [MW]'][:t_sim] # Specify name of column from the input data Excel

"MI" # List of available fuels (electricity as an option in the case of electric HOB)
fuels = ['gas', 'electricity']
# Fuel prices
p_f = {fuels[0]: df['Gas price [€/MWh]'][:t_sim], fuels[1]: p_E} # Specify name of column from the input data Excel

"MI" # Transmission costs for fuels [€/MWh]
p_f_tr = {fuels[0]: 11, fuels[1]: p_E_tr_buy}

p_f_tax = {fuels[0]: 23.354, fuels[1]: 0}

"MI" # CO2 emission factor for electricity (used in Equation (49))[kg/MWh] (Fingrid estimation for 2024)
omega = 33

# Calculation for emission factors for fuels (used in Equations 64 and 65)
CO2_f = {'gas': 189.2}

"MI" # Additional data if necessary

# ======================================================================================================================
# CHPa: CHP units used for supplying electricity and heat
# ======================================================================================================================

"MI" # Set of CHP units
N_CHPa = ['CHPa1', 'CHPa2', 'CHPa3']

"MI" # CHP unit fuels. The fuels are picked from the "fuels" dictionary for each unit
# indexing in the "fuels dictionary starts from "0"
f_CHPa = {'CHPa1':fuels[0], 'CHPa2':fuels[0], 'CHPa3':fuels[0]}
"MI" # Transmission costs for fuel of the CHP units [€/MWh]
p_f_tr_CHPa = {i: p_f_tr[f_CHPa[i]] for i in N_CHPa}
"MI" # Energy tax for fuels of the CHPa units [€/MWh]
p_f_tax_CHPa = {i: p_f_tax[f_CHPa[i]] - 5.63 for i in N_CHPa}
# Total price of fuel for CHPa units [€/MWh]
p_f_CHPa = {i: p_f[f_CHPa[i]] + p_f_tax_CHPa[i] + p_f_tr_CHPa[i] for i in N_CHPa}
# Flatten p_fuel_CHP into (i, t) -> value (Pyomo accessible format)
p_f_CHPa_flat = {
    (i, t): p_f_CHPa[i].iloc[t - 1]
    for i in N_CHPa
    for t in T
}

# CO2 emission factors for each CHPa unit [kg CP2 / MWh fuel]
CO2_f_CHPa = {i: CO2_f[f_CHPa[i]] for i in N_CHPa}

# check for fuel prices to see if they are realistic
for i in N_CHPa:
    print(f"CHP {i} fuel price at t=1: {p_f_CHPa[i][1]} €/MWh")
# check for CO2 emission factors to see if they are realistic
for i in N_CHPa:
    print(f"CHP {i} CO2 factor at t=1: {CO2_f_CHPa[i]} kg CO2/MWh fuel")

"MI" # Partial efficiencies for electricity production [-]
eta_E_CHPa = {'CHPa1':0.43, 'CHPa2':0.43, 'CHPa3':0.43}
"MI" # Partial efficiencies for heat production [-]
eta_Q_CHPa = {'CHPa1':0.45, 'CHPa2':0.45, 'CHPa3':0.45}
# Heat-to-power ratios [-]
r_EQ_CHPa = {n: eta_E_CHPa[n] / eta_Q_CHPa[n] for n in N_CHPa}
"MI" # Minimum heat generation for each CHP unit [MW]
Q_min_CHPa = {'CHPa1':0.5 / r_EQ_CHPa['CHPa1'], 'CHPa2':0.5 / r_EQ_CHPa['CHPa2'], 'CHPa3':0.4 / r_EQ_CHPa['CHPa3']}
"MI" # Maximum heat generation for each CHPa unit [MW]
Q_max_CHPa = {'CHPa1':1.6 / r_EQ_CHPa['CHPa1'], 'CHPa2':1.6 / r_EQ_CHPa['CHPa2'], 'CHPa3':1.2 / r_EQ_CHPa['CHPa3']}

"MI" # CHP ramp-up time (0-100% power)[h]
rut_CHPa = {'CHPa1':0.25, 'CHPa2':0.25, 'CHPa3':0.25}
"MI" # CHP ramp-down time (100-0% power)[h]
rdt_CHPa = {'CHPa1':0.25, 'CHPa2':0.25, 'CHPa3':0.25}
"MI" # Fixed start-up prices for each CHP unit []
p_su_CHPa = {'CHPa1': 40, 'CHPa2': 40, 'CHPa3': 30}

# ======================================================================================================================
# CHPb: CHP units used to supply only electricity
# ======================================================================================================================

"MI" # Set of CHPb units
N_CHPb = ['CHPb']

"MI" # CHPb unit fuels. The fuels are picked from the "fuels" dictionary for each unit
# indexing in the "fuels dictionary starts from "0"
f_CHPb = {'CHPb':fuels[0]}
"MI" # Transmission costs for fuel of the CHPb units [€/MWh]
p_f_tr_CHPb = {i: p_f_tr[f_CHPb[i]] for i in N_CHPb}
"MI" # Energy tax for fuels of the CHPb units [€/MWh]
p_f_tax_CHPb = {i: p_f_tax[f_CHPb[i]] - 5.63 for i in N_CHPb}
# Total fuel prices for each CHPb unit [€/MWh]
p_f_CHPb = {i: p_f[f_CHPb[i]] + p_f_tax_CHPb[i] + p_f_tr_CHPb[i] for i in N_CHPb}
# Flatten p_fuel_CHP into (i, t) -> value (Pyomo accessible format)
p_f_CHPb_flat = {
    (i, t): p_f_CHPb[i].iloc[t - 1]
    for i in N_CHPb
    for t in T
}

# CO2 emission factors for each CHPa unit
CO2_f_CHPb = {i: CO2_f[f_CHPb[i]] for i in N_CHPb}

# check for fuel prices to see if they are realistic
for i in N_CHPb:
    print(f"CHP {i} fuel price at t=1: {p_f_CHPb[i][1]} €/MWh")
# check for CO2 emission factors to see if they are realistic
for i in N_CHPb:
    print(f"CHP {i} CO2 factor at t=1: {CO2_f_CHPb[i]} kg CO2/MWh fuel")

"MI" # partial efficiencies for electricity production [-]
eta_E_CHPb = {'CHPb':0.43}
"MI" # partial efficiencies for heat production [-]
eta_Q_CHPb = {'CHPb':0.45}
# ratios for heat and power production [-]
r_EQ_CHPb = {i: eta_E_CHPb[i] / eta_Q_CHPb[i] for i in N_CHPb}
"MI" # Maximum heat generation for each CHPb unit [MW]
Q_max_CHPb = {'CHPb': 1.2 / r_EQ_CHPb['CHPb']}
"MI" # Minimum heat generation for each CHPb unit [MW]
Q_min_CHPb = {'CHPb': 0.4 / r_EQ_CHPb['CHPb']}

"MI" # CHPb ramp-up time (0-100% power)[h]
rut_CHPb = {'CHPb':0.25}
"MI" # CHPb ramp-down time (100-0% power)[h]
rdt_CHPb = {'CHPb':0.25}

"MI" # fixed start-up prices for each CHPb unit [€]
p_su_CHPb = {'CHPb': 30}


# ======================================================================================================================
# HPa: HP units used to recover waste heat from the CHP and ES units and to operate in ATW mode
# ======================================================================================================================

"MI" # Set of HPa units
N_HPa = ['HPa']

"MI" # Maximum heating output for each HPa unit [MW]
Q_max_HPa = {'HPa':0.6}
"MI" # Minimum heating output for each HPa unit [MW]
Q_min_HPa = {'HPa':0.01}

"MI" # Heat sink inlet temperatures for each HPa unit [C]
T_sink_in_HPa = {'HPa': 35}
"MI" # Heat sink outlet temperatures for each HPa unit [C]
# Specify name of column from the input data Excel
T_sink_out_HPa = {'HPa': df['T_sink_out [C]'][:t_sim]}
# Flatten T_sink_out into (i, t) -> value (Pyomo accessible format)
T_sink_out_HPa_flat = {(i, t): T_sink_out_HPa[i].iloc[t - 1] for i in N_HPa for t in T}
# Calculation for the thermodynamic mean temperatures for the heat sinks [K] (Equation (9))
T_sink_lm_HPa = {
    i: {
        t: ((T_sink_out_HPa[i].iloc[t - 1] + 273.15) - (T_sink_in_HPa[i] + 273.15)) /
           math.log((T_sink_out_HPa[i].iloc[t - 1] + 273.15) / (T_sink_in_HPa[i] + 273.15))
        for t in T
    }
    for i in N_HPa
}

"MI" # Heat source inlet temperatures for each HP unit [C]
T_source_in_HPa = {'HPa': 20}
"MI" # Heat source outlet temperatures for each HP unit [C]
T_source_out_HPa = {'HPa': 15}
# Calculation for the thermodynamic mean temperatures for the heat sources [K] (Equation (10))
T_source_lm_HPa = {i: ((T_source_in_HPa[i] + 273.15) - (T_source_out_HPa[i] + 273.15)) \
                      / math.log((T_source_in_HPa[i] + 273.15) / (T_source_out_HPa[i] + 273.15)) for i in N_HPa}

# Linear fit parameters for calculating the second law efficiency
# (eta_2nd = a_2nd * mean temperature lift + b_2nd) (Equation (11))
# based on data collected by Zühlsdorf et al. (2023) and presented by Magni et al. (2024)
a_2nd = 0.0056
b_2nd = 0.1558

# COP calculation for each HPa unit [-] (Equations (7))
COP_HPa = {
    i: {
        t: T_sink_lm_HPa[i][t] / (T_sink_lm_HPa[i][t] - T_source_lm_HPa[i]) *
           (a_2nd * (T_sink_lm_HPa[i][t] - T_source_lm_HPa[i]) + b_2nd)
        for t in T
    }
    for i in N_HPa
}
# Flatten COP into (i, t) -> value (Pyomo accessible format)
COP_HPa_flat = {(i, t): COP_HPa[i][t] for i in N_HPa for t in T}

"MI" # Ramp-up times from 0-100% power for each HPa unit [h]
rut_HPa = {'HPa': 20 / 60}
"MI" # Ramp-down times from 100-0% power for each HPa unit [h]
rdt_HPa = {'HPa': 20 / 60}

"MI" # Price for HPa unit startup [€]
p_su_HPa = {'HPa':5}

"MI" # Maximum value for the COP reduction factor during ramp-up [-]
n_ru_max_HPa = {'HPa':0.2}

# COP in ATW mode calculation for each HPa unit [-]
COP_ATW_HPa = {
    i: {
        t: T_sink_lm_HPa[i][t] / (T_sink_lm_HPa[i][t] - (T_0.iloc[t - 1] + 273.15)) *
           (a_2nd * (T_sink_lm_HPa[i][t] - (T_0.iloc[t - 1] + 273.15)) + b_2nd)
        for t in T
    }
    for i in N_HPa
}
# Flatten ATW mode COP into (i, t) -> value (Pyomo accessible format)
COP_ATW_HPa_flat = {(i, t): COP_ATW_HPa[i][t] for i in N_HPa for t in T}

"MI" # ATW mode minimum temperature [C]
T_min_ATW ={'HPa':5}
"MI" # Theoretical maximum heating output for each HPa unit in ATW mode [MW]
# If there is no possibility for ATW operation, Q_max_ATW_HPa = 0
Q_max_ATW_HPa = {'HPa':0.4}

# Achievable maximum heating output for each HPa unit in ATW mode based on outside air temperature [MW] (Equation (16))
mu = 0.12

Q_max_ATW_HPa_real = {
    i: {
        t: Q_max_ATW_HPa[i] * (1 - math.exp(-mu * T_0.iloc[t - 1]))
        for t in T
    }
    for i in N_HPa
}
# Flatten Q_ATW_max into (i, t) -> value (Pyomo accessible format)
Q_max_ATW_HPa_real_flat = {(i, t): Q_max_ATW_HPa_real[i][t] for i in N_HPa for t in T}

# Check for COP values to see if they are realistic
for i in N_HPa:
    print(f"Heat pump {i} COP at t=1: {COP_HPa[i][1]}")
    print(f'Q_ATW_max {i} at t=1: {Q_max_ATW_HPa_real[i][1]}')

"MI" # Waste heat flow estimation coefficient [1/C]
a_w = 0.005
"MI" # Waste heat flow estimation coefficient [-]
b_w = 0.5
"MI" # Efficiency for ES [-]
eta_ES = 0.9
"MI" # Electricity flow going through the ES  [MW]
E_ES = 2

# ======================================================================================================================
# HOB
# ======================================================================================================================

"MI" # Set of HOB units
N_HOB = ['HOB']

"MI" # HOB unit fuels. Pick a fuel from the "fuels" dictionary for each unit
# indexing in the "fuels dictionary starts from "0"
f_HOB = {'HOB': fuels[0]}
"MI" # Energy tax for fuels for HOB units [€/MWh]
p_f_tax_HOB = {i: p_f_tax[f_HOB[i]] for i in N_HOB}
"MI" # Transmission price for fuel of the HOB units
p_f_tr_HOB = {i: p_f_tr[f_HOB[i]] for i in N_HOB}
# Total fuel prices for each HOB unit [€/MWh]
p_f_HOB = {i: p_f[f_HOB[i]] + p_f_tax_HOB[i] + p_f_tr_HOB[i] for i in N_HOB}
# Flatten p_fuel_HOB into (i, t) -> value (Pyomo accessible format)
p_f_HOB_flat = {
    (i, t): p_f_HOB[i].iloc[t - 1]
    for i in N_HOB
    for t in T
}
# CO2 emission factors for each HOB unit
CO2_f_HOB = {i: CO2_f[f_HOB[i]] for i in N_HOB}

# Check for fuel prices to see if they are realistic
for i in N_HOB:
    print(f"HOB {i} fuel price at t=1: {p_f_HOB[i][1]} €/MWh")
# Check for CO2 emission factors to see if they are realistic
for i in N_HOB:
    print(f"HOB {i} CO2 factor at t=1: {CO2_f_HOB[i]} kg CO2/MWh fuel")

"MI" # HOB unit heat generation efficiencies [-]
eta_Q_HOB = {'HOB':0.95}
"MI" # Maximum heat output for each HOB unit [MW]
Q_max_HOB = {'HOB':5}
"MI" # Minimum heat output for each HOB unit [MW]
Q_min_HOB = {'HOB':0.05}

"MI" # HOB ramp-up time (0-100% power)[h]
rut_HOB = {'HOB': 10 / 60}
"MI"# HOB ramp-down time (100-0% power)[h]
rdt_HOB = {'HOB': 10 / 60}
"MI" # Fixed start-up price [€]
p_su_HOB = {'HOB':20}

# ======================================================================================================================
# TS
# Variables and parameters
# ======================================================================================================================

# Water specific heat capacity [J/(kgC)]
c_water = 4186
# Water density [kg/m^3]
rho_water = 1000

"MI" # Set of TS units
N_TS = ['TS']

"MI" # Water column diameter [m]
d = {'TS':6}
"MI" # TS unit hot and cold zone temperatures [C]
T_H_TS = {'TS':95}
T_C_TS = {'TS':35}
"MI" # lowered cold zone temperature in TS unit [C]
# If there is no integrated HPb unit: T_C_new_TS[i] = T_C_TS[i]
T_C_new_TS = {'TS':25}
"MI" # Maximum theoretical storage capacity in each TS unit (before HPb integration) [MWh]
Q_max_TS = {'TS':20 * (T_H_TS['TS'] - T_C_new_TS['TS']) / (T_H_TS['TS'] - T_C_TS['TS'])}
"MI" # Thermal transmittance (U-value) = conductivity of material / wall thickness [W/(m^2K)]
U_TS = {'TS':0.25}
"MI" # Minimum allowed storage level [-]
lvl_min_TS = {'TS':0}
"MI" # Maximum allowed storage level [-]
lvl_max_TS = {'TS':1}

# Check for maximum heat storage capacity
for i in N_TS:
    print(f"TS{i} maximum storage capacity: {Q_max_TS[i] * lvl_max_TS[i]} MWh")

"MI" # Efficiencies of heat transfer into TS units [-]
eta_in_TS = {'TS':0.95}
"MI" # Efficiencies of heat transfer out of TS units [-]
eta_out_TS = {'TS':0.95}
"MI" # Maximum achievable heat flow into TS units [MW]
Q_in_max_TS = {'TS':2.5}
"MI" # Maximum achievable heat flow out of TS units [MW]
Q_out_max_TS = {'TS':2.5}
"MI" # Minimum achievable heat flow into TS units [MW]
Q_in_min_TS = {'TS':0.01}
"MI" # Minimum achievable heat flow out of TS units [MW]
Q_out_min_TS = {'TS':0.01}

"MI" # Start-up prices for charging and discharging [€]
p_in_su_TS = {'TS':5}
p_out_su_TS = {'TS':5}

# Heat loss coefficients for idle storage (oemof.thermal)
# Losses through the lateral surface of the high-temperature section of the water body (dt in [s]) [-] (Equation (41))
beta_TS = {i: U_TS[i] * 4 / (d[i] * rho_water * c_water) * dt * 3600 for i in N_TS}
# Losses through the total lateral surface assuming the storage is empty (dt in [s]) [-] (Equation (42))
gamma_TS_dict = {
    t: U_TS[i] * 4 / (d[i] * rho_water * c_water * (T_H_TS[i] - T_C_new_TS[i])) \
       * (T_C_new_TS[i] - T_0.iloc[t - 1]) * dt * 3600
    for t in T for i in N_TS
}
gamma_TS = {i:gamma_TS_dict for i in N_TS}
# Flatten gamma into (i, t) -> value (Pyomo accessible format)
gamma_TS_flat = {(i, t): gamma_TS[i][t] for i in gamma_TS for t in gamma_TS[i]}
# Losses through the bottom and top surfaces (dt in [s])[MWh] (Equation (43))
delta_TS_dict = {
    t: U_TS[i] * math.pi * (d[i] * d[i]) / 4 * ((T_H_TS[i] - T_0.iloc[t - 1]) + (T_C_new_TS[i] - T_0.iloc[t - 1])) \
       * dt * 3600 / (3.6 * 10 ** 9)
    for t in T for i in N_TS
}
delta_TS = {i:delta_TS_dict for i in N_TS}
# Flatten delta into (i, t) -> value (Pyomo accessible format)
delta_TS_flat = {(i, t): delta_TS[i][t] for i in delta_TS for t in delta_TS[i]}

# ======================================================================================================================
# HPb: HPs integrated with TS units
# ======================================================================================================================

"MI" # Set of HPb units
N_HPb = ['HPb']
"MI" # Maximum heating output for each HPb unit [MW]
Q_max_HPb = {'HPb':0.8}
"MI" # Minimum heating output for each HPb unit [MW]
Q_min_HPb = {'HPb':0.01}

"MI" # Heat sink inlet temperatures for each HPb unit [C]
T_sink_in_HPb = {'HPb': T_C_TS['TS']}
"MI" # Heat sink outlet temperatures for each HP unit [C]
T_sink_out_HPb = {'HPb': df['T_sink_out [C]'][:t_sim]} # Specify name of column from the input data Excel
# Calculation for the thermodynamic mean temperatures for the heat sinks [K] (Equation (9))
T_sink_lm_HPb = {
    i: {
        t: ((T_sink_out_HPb[i].iloc[t - 1] + 273.15) - (T_sink_in_HPb[i] + 273.15)) /
           math.log((T_sink_out_HPb[i].iloc[t - 1] + 273.15) / (T_sink_in_HPb[i] + 273.15))
        for t in T
    }
    for i in N_HPb
}

"MI" # Heat source inlet temperatures for each HPb unit [K]
T_source_in_HPb = {'HPb': T_C_TS['TS']}
"MI" # Heat source outlet temperatures for each HPb unit [K]
T_source_out_HPb = {'HPb': T_C_new_TS['TS']}
# Calculation for the thermodynamic mean temperatures for the heat sources [K] (Equation (10))
T_source_lm_HPb = {i: ((T_source_in_HPb[i] + 273.15) - (T_source_out_HPb[i] + 273.15)) \
                      / math.log((T_source_in_HPb[i] + 273.15) / (T_source_out_HPb[i] + 273.15)) for i in N_HPb}

# COP calculation for each HPb unit [-] (Equation (7))
COP_HPb = {
    i: {
        t: T_sink_lm_HPb[i][t] / (T_sink_lm_HPb[i][t] - T_source_lm_HPb[i]) *
           (a_2nd * (T_sink_lm_HPb[i][t] - T_source_lm_HPb[i]) + b_2nd)
        for t in T
    }
    for i in N_HPb
}
# Flatten COP into (i, t) -> value (Pyomo accessible format)
COP_HPb_flat = {(i, t): COP_HPb[i][t] for i in N_HPb for t in T}

"MI" # Ramp-up times from 0-100% power for each HPb unit [h]
rut_HPb = {'HPb': 30 / 60}
"MI" # Ramp-down times from 100-0% power for each HPb unit [h]
rdt_HPb = {'HPb': 30 / 60}

"MI" # Price for HPb unit startup [€]
p_su_HPb = {'HPb':10}

"MI" # Maximum value for the COP reduction factor during ramp-up [-]
n_ru_max_HPb = {'HPb':0.2}

# Check for COP values to see if they are realistic
for i in N_HPb:
    print(f"Heat pump {i} COP at t=1: {COP_HPb[i][1]}")



