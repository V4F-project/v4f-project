# Power system dispatch optimization model

# This file contains the Pyomo optimization model for a power system containing heat pumps, CHP units, heat-only boilers
# and thermal storage units.
# Ín order to run the model, the name of the input data (e.g. "input_data") file needs to be specified in this file
# requires input section is marked with "MI".
# Constraint and expression formulations are explained very briefly in this file. Fore more detailed information on
#  the formulation of the constraints and expressions, references for the equations presented in the Thesis are given.

# The model is formulated using Pyomo optimization language
# Import pyomo environment
from pyomo.environ import *
from pyomo.opt import SolverFactory
# Import pandas for the purpose of using importing the results to Excel
import pandas
"MI"# Import model parameters from data file
import Input_data as id # specify name of the data input file

# Model type and name
model = ConcreteModel(name = "Generic_power_system")

# ======================================================================================================================
# Non-unit specific parameters
# ======================================================================================================================

# Simulation time horizon
model.N_T = Set(initialize=id.T)
# Simulation time interval [h]
model.dt = Param(initialize=id.dt)
# Assuming model.T is a list-like ordered set of integers
T_list = list(model.N_T)
# The last timestep
T_max = T_list[-1]
# The previous timestep (t-1) is modelled as its own parameter for the looped simulation time horizon
model.t_prev = Param(model.N_T, initialize=lambda model, t: T_list[-1] if t == T_list[0] else t - 1)

# Electricity price [€/MWh]
model.p_E = Param(model.N_T, initialize={i: id.p_E[i - 1] for i in id.T})
# Electricity transmission prices for buying and selling [€/MWh]
model.p_E_tr_buy = Param(initialize=id.p_E_tr_buy)
model.p_E_tr_sell = Param(initialize=id.p_E_tr_sell)
# Price for heat [€/MWh]
model.p_Q = Param(initialize=id.p_Q)

# Outside air temperature [C]
model.T_0 = Param(model.N_T, initialize={i: id.T_0[i - 1] for i in id.T})


# ======================================================================================================================
# CHPa: CHP units used for supplying electricity and heat
# Variables and parameters
# ======================================================================================================================

# Set of CHPa units
model.N_CHPa = Set(initialize=id.N_CHPa)

# Binary decision variables
# Status variables [-]
model.y_CHPa = Var(model.N_CHPa, model.N_T, domain=Binary)
# Start-up variables [-]
model.y_su_CHPa = Var(model.N_CHPa, model.N_T, domain=Binary)

# Continuous decision variables
# Electricity generation [MW]
model.E_CHPa = Var(model.N_CHPa, model.N_T, domain=NonNegativeReals)
# Heat generation [MW]
model.Q_CHPa = Var(model.N_CHPa, model.N_T, domain=NonNegativeReals)

# Parameters
# Fuel prices including spot, transmission and tax [€/MWh]
model.p_f_CHPa = Param(model.N_CHPa, model.N_T, initialize=id.p_f_CHPa_flat)
# CO2 emission coefficients of fuels [kg CO2/MWh fuel]
model.CO2_f_CHPa = Param(model.N_CHPa, initialize=id.CO2_f_CHPa)
# Partial efficiencies for electricity production [-]
model.eta_E_CHPa = Param(model.N_CHPa, initialize=id.eta_E_CHPa)
# partial efficiencies for heat production [-]
model.eta_Q_CHPa = Param(model.N_CHPa, initialize=id.eta_Q_CHPa)
# Power to heat ratios [-]
model.r_EQ_CHPa = Param(model.N_CHPa, initialize=id.r_EQ_CHPa)
# Maximum heat generation capacities [MW]
model.Q_max_CHPa = Param(model.N_CHPa, initialize=id.Q_max_CHPa)
# Minimum heat generation capacities [MW]
model.Q_min_CHPa = Param(model.N_CHPa, initialize=id.Q_min_CHPa)
# Ramp-up times [h]
model.rut_CHPa = Param(model.N_CHPa, initialize=id.rut_CHPa)
# Ramp_down times [h]
model.rdt_CHPa = Param(model.N_CHPa, initialize=id.rdt_CHPa)
# Start-up prices [€]
model.p_su_CHPa = Param(model.N_CHPa, initialize=id.p_su_CHPa)

# ======================================================================================================================
# CHPa: CHP units used for supplying electricity and heat
# Constraints and expressions
# ======================================================================================================================

# CHPa unit start-up is tracked: model.y_su_CHPa[i,t] gains value of 1 when model.y_CHPa[i, t] switches from 0 to 1.
# (Equation (1))
def CHPa_startup_rule(model, i, t):
    return model.y_su_CHPa[i, t] >= model.y_CHPa[i, t] - model.y_CHPa[i, model.t_prev[t]]
model.CHPa_startup_constraint = Constraint(model.N_CHPa, model.N_T, rule=CHPa_startup_rule)

# Start-up costs for CHPa units [€] (Equation (2))
def C_su_CHPa_rule(model, i,t):
    return model.y_su_CHPa[i, t] * model.p_su_CHPa[i]
model.C_su_CHPa = Expression(model.N_CHPa, model.N_T, rule=C_su_CHPa_rule)

# Maximum heat production capacity of the CHPa units (Equation (3))
def CHPa_Q_max_rule(model, i, t):
    return model.Q_CHPa[i,t] <= model.Q_max_CHPa[i] * model.y_CHPa[i, t]
model.CHPa_Q_max_constraint = Constraint(model.N_CHPa, model.N_T, rule=CHPa_Q_max_rule)

# Minimum heat production capacity of the CHPa units (Equation (3))
def CHPa_Q_min_rule(model, i, t):
    return model.Q_CHPa[i,t] >= model.Q_min_CHPa[i] * model.y_CHPa[i, t]
model.CHPa_Q_min_constraint = Constraint(model.N_CHPa, model.N_T, rule=CHPa_Q_min_rule)

# Heat to power ratio
def CHPa_EQ_ratio_rule(model, i,t):
    return model.E_CHPa[i,t] == model.r_EQ_CHPa[i] * model.Q_CHPa[i,t]
model.CHPa_ratio_constraint = Constraint(model.N_CHPa, model.N_T, rule=CHPa_EQ_ratio_rule)

# CHPa unit ramp-up constraint (Equation (4))
def CHPa_rampup_rule(model, i, t):
    return model.Q_CHPa[i,t] - model.Q_CHPa[i, model.t_prev[t]] <= model.Q_max_CHPa[i] / model.rut_CHPa[i] * model.dt
model.CHPa_rampup_constraint = Constraint(model.N_CHPa, model.N_T, rule=CHPa_rampup_rule)

# CHPa unit ramp-down constraint (Equation (5))
def CHPa_rampdown_rule(model, i, t):
    return model.Q_CHPa[i, model.t_prev[t]] - model.Q_CHPa[i, t] <= model.Q_max_CHPa[i] / model.rdt_CHPa[i] * model.dt
model.CHPa_rampdown_constraint = Constraint(model.N_CHPa, model.N_T, rule=CHPa_rampdown_rule)

# Fuel consumption costs of CHPa units [€] (Equation (45))
def C_con_CHPa_rule(model, i,t):
    return model.p_f_CHPa[i, t] * model.Q_CHPa[i, t] / model.eta_Q_CHPa[i] * model.dt
model.C_con_CHPa = Expression(model.N_CHPa, model.N_T, rule=C_con_CHPa_rule)

# Revenue from selling electricity generated by CHPa units [€] (Equation (46))
def Revenue_E_a_rule(model,i,t):
    return (model.p_E[t] - model.p_E_tr_sell) * model.E_CHPa[i, t] * model.dt
model.R_E_a = Expression(model.N_CHPa, model.N_T, rule=Revenue_E_a_rule)

# Total costs for each CHPa unit [€] (Equation (50))
def C_tot_CHPa_rule(model, i,t):
    return model.C_con_CHPa[i,t] + model.C_su_CHPa[i,t]
model.C_tot_CHPa = Expression(model.N_CHPa, model.N_T, rule=C_tot_CHPa_rule)

# ======================================================================================================================
# CHPb: CHP units used to supply only electricity
# Variables and parameters
# ======================================================================================================================

# Set of CHPb units
model.N_CHPb = Set(initialize=id.N_CHPb)

# Binary decision variables
# Status variables [-]
model.y_CHPb = Var(model.N_CHPb, model.N_T, domain=Binary)
# Start-up variables [-]
model.y_su_CHPb = Var(model.N_CHPb, model.N_T, domain=Binary)

# Continuous decision variables
# Electricity generation [MW]
model.E_CHPb = Var(model.N_CHPb, model.N_T, domain=NonNegativeReals)
# Heat generation [MW]
model.Q_CHPb = Var(model.N_CHPb, model.N_T, domain=NonNegativeReals)

# Parameters
# Fuel prices including spot, transmission, and tax [€/MWh]
model.p_f_CHPb = Param(model.N_CHPb, model.N_T, initialize=id.p_f_CHPb_flat)
# CO2 emission coefficients for fuels [kg CO2/MWh fuel]
model.CO2_f_CHPb = Param(model.N_CHPb, initialize=id.CO2_f_CHPb)
# Partial efficiencies for electricity production [-]
model.eta_E_CHPb = Param(model.N_CHPb, initialize=id.eta_E_CHPb)
# Partial efficiencies for heat production [-]
model.eta_Q_CHPb = Param(model.N_CHPb, initialize=id.eta_Q_CHPb)
# Power to heat ratios [-]
model.r_EQ_CHPb = Param(model.N_CHPb, initialize=id.r_EQ_CHPb)
# Maximum heat generation capacity [MW]
model.Q_max_CHPb = Param(model.N_CHPb, initialize=id.Q_max_CHPb)
# Minimum heat generation capacity [MW]
model.Q_min_CHPb = Param(model.N_CHPb, initialize=id.Q_min_CHPb)
# Ramp-up times [h]
model.rut_CHPb = Param(model.N_CHPb, initialize=id.rut_CHPb)
# Ramp-down times [h]
model.rdt_CHPb = Param(model.N_CHPb, initialize=id.rdt_CHPb)
# Start-up prices [€/start-up]
model.p_su_CHPb = Param(model.N_CHPb, initialize=id.p_su_CHPb)

# ======================================================================================================================
# CHPb: CHP units used to supply only electricity
# Constraints and expressions
# ======================================================================================================================

# CHPb unit start-up is tracked: model.y_su_CHPb[i,t] gains value of 1 when model.y_CHPb[i, t] switches from 0 to 1.
# (Equation (1))
def CHPb_startup_rule(model, i, t):
    return model.y_su_CHPb[i, t] >= model.y_CHPb[i, t] - model.y_CHPb[i, model.t_prev[t]]
model.CHPb_startup_constraint = Constraint(model.N_CHPb, model.N_T, rule=CHPb_startup_rule)

# Costs for starting CHPb units [€] (Equation (2))
def C_su_CHPb_rule(model, i,t):
    return model.y_su_CHPb[i, t] * model.p_su_CHPb[i]
model.C_su_CHPb = Expression(model.N_CHPb, model.N_T, rule=C_su_CHPb_rule)

# Maximum heat generation capacity of the CHPb units (Equation (3))
def CHPb_Q_max_rule(model, i, t):
    return model.Q_CHPb[i,t] <= model.Q_max_CHPb[i] * model.y_CHPb[i, t]
model.CHPb_Q_max_constraint = Constraint(model.N_CHPb, model.N_T, rule=CHPb_Q_max_rule)

# Maximum heat generation capacity of the CHPb units (Equation (3))
def CHPb_Q_min_rule(model, i, t):
    return model.Q_CHPb[i,t] >= model.Q_min_CHPb[i] * model.y_CHPb[i, t]
model.CHPb_Q_min_constraint = Constraint(model.N_CHPb, model.N_T, rule=CHPb_Q_min_rule)

# Ratio of heat and power in the CHPb units
def CHPb_EQ_ratio_rule(model, i,t):
    return model.E_CHPb[i,t] == model.r_EQ_CHPb[i] * model.Q_CHPb[i,t]
model.CHPb_ratio_constraint = Constraint(model.N_CHPb, model.N_T, rule=CHPb_EQ_ratio_rule)

# CHPb ramp-up constraint (Equation (4))
def CHPb_rampup_rule(model, i, t):
    return model.Q_CHPb[i,t] - model.Q_CHPb[i, model.t_prev[t]] <= model.Q_max_CHPb[i] / model.rut_CHPb[i] * model.dt
model.CHPb_rampup_constraint = Constraint(model.N_CHPb, model.N_T, rule=CHPb_rampup_rule)

# CHPb ramp-down constraint (Equation (5))
def CHPb_rampdown_rule(model, i, t):
    return model.Q_CHPb[i, model.t_prev[t]] - model.Q_CHPb[i, t] <= model.Q_max_CHPb[i] / model.rdt_CHPb[i] * model.dt
model.CHPb_rampdown_constraint = Constraint(model.N_CHPb, model.N_T, rule=CHPb_rampdown_rule)

# Fuel consumption costs for the CHPb units [€] (Equation (45))
def C_con_CHPb_rule(model, i, t):
    return model.p_f_CHPb[i, t] * model.Q_CHPb[i, t] / model.eta_Q_CHPb[i] * model.dt
model.C_con_CHPb = Expression(model.N_CHPb, model.N_T, rule=C_con_CHPb_rule)


# Revenue from selling electricity generated by the CHPb units [€]  (Equation (46))
def Revenue_E_b_rule(model,i, t):
    return (model.p_E[t] - model.p_E_tr_sell) * model.E_CHPb[i, t] * model.dt
model.R_E_b = Expression(model.N_CHPb, model.N_T, rule=Revenue_E_b_rule)

# Price of heat generated by the CHPb units that is not utilized [€] (Equation (47))
def C_Ql_CHPb_rule(model,i, t):
    return model.p_Q * model.Q_CHPb[i,t] * model.dt
model.C_Ql_CHPb = Expression(model.N_CHPb, model.N_T, rule=C_Ql_CHPb_rule)

# Total costs for each CHPb unit [€] (Equation (50))
def C_tot_CHPb_rule(model, i,t):
    return model.C_con_CHPb[i,t] + model.C_su_CHPb[i,t] + model.C_Ql_CHPb[i,t]
model.C_tot_CHPb = Expression(model.N_CHPb, model.N_T, rule=C_tot_CHPb_rule)

# ======================================================================================================================
# HPa: HP units used to recover waste heat from the CHP and ES units and to operate in ATW mode
# Variables and parameters
# ======================================================================================================================

# Set of HPa units
model.N_HPa = Set(initialize=id.N_HPa)

# Binary decision variables
# Status variables [-]
model.y_HPa = Var(model.N_HPa, model.N_T, domain=Binary)
# Start-up variables [-]
model.y_su_HPa = Var(model.N_HPa, model.N_T, domain=Binary)

# Continuous decision variables
# Heat generation [MW]
model.Q_HPa = Var(model.N_HPa, model.N_T, domain=NonNegativeReals)
# Recovered waste heat to be upgraded [MW]
model.Q_rec_HPa = Var(model.N_HPa, model.N_T, domain=NonNegativeReals)
# heat generated in ATW mode [MW]
model.Q_ATW_HPa = Var(model.N_HPa, model.N_T, domain=NonNegativeReals)
# COP reduction factor during ramp-up [-]
model.n_ru_HPa = Var(model.N_HPa, model.N_T, domain=NonNegativeReals)

# Parameters
# Start-up prices [€/start-up]
model.p_su_HPa = Param(model.N_HPa, initialize=id.p_su_HPa)
# Maximum heat generation capacity [MW]
model.Q_max_HPa = Param(model.N_HPa, initialize=id.Q_max_HPa)
# Minimum heat generation capacity [MW]
model.Q_min_HPa = Param(model.N_HPa, initialize=id.Q_min_HPa)
# Maximum heat generation capacity in ATW mode [MW]
model.Q_max_ATW_HPa_real = Param(model.N_HPa, model.N_T, initialize=id.Q_max_ATW_HPa_real_flat)
# COP [-]
model.COP_HPa = Param(model.N_HPa, model.N_T, initialize=id.COP_HPa_flat)
# COP in ATW mode [-]
model.COP_ATW_HPa = Param(model.N_HPa, model.N_T, initialize=id.COP_ATW_HPa_flat)
# Minimum outside air temperature for ATW mode [C]
model.T_min_ATW_HPa = Param(model.N_HPa, initialize=id.T_min_ATW)
# Maximum value for COP reduction factor during ramp-up [-]
model.n_ru_max_HPa = Param(model.N_HPa, initialize=id.n_ru_max_HPa)
# Waste heat flow coefficient [1/C]
model.a_w = Param(initialize=id.a_w)
# Waste heat flow coefficient [-]
model.b_w = Param(initialize=id.b_w)
# Efficiency of the ES [-]
model.eta_ES = Param(initialize=id.eta_ES)
# Electricity flow through the ES [MW]
model.E_ES = Param(initialize=id.E_ES )
# Ramp-up time [h]
model.rut_HPa = Param(model.N_HPa, initialize=id.rut_HPa)
# Ramp-down time [h]
model.rdt_HPa = Param(model.N_HPa, initialize=id.rdt_HPa)
# HPa sink outlet temperature
model.T_sink_out_HPa = Param(model.N_HPa, model.N_T, initialize=id.T_sink_out_HPa_flat)

# ======================================================================================================================
# HPa: HP units used to recover waste heat from the CHP and ES units and to operate in ATW mode
# Constraints and expressions
# ======================================================================================================================

"MI" # Maximum available waste heat for HPa [MW] (Equation (14)). Waste heat flow for each HPa unit needs to be
# specified individually
def HPa_rec_max_rule(model, t):
    return ((1 - model.eta_Q_CHPa['CHPa1'] - model.eta_E_CHPa['CHPa1']) * model.Q_CHPa['CHPa1',t] \
            / model.eta_Q_CHPa['CHPa1'] + (1 - model.eta_Q_CHPa['CHPa2'] - model.eta_E_CHPa['CHPa2']) \
            * model.Q_CHPa['CHPa2',t] / model.eta_Q_CHPa['CHPa2'] + (1 - model.eta_ES) * model.E_ES) \
               * (model.a_w * model.T_0[t] + model.b_w)
model.HPa_rec_max = Expression(model.N_T, rule=HPa_rec_max_rule)

"MI" # Recovered amount of waste heat cannot exceed the available amount for HPa (Equation (14)) This constraint needs
# to be specified for each HPa unit specifically
def HPa_rec_rule(model, t):
    return model.Q_rec_HPa['HPa',t] <= model.HPa_rec_max[t]
model.HPa_rec_constraint = Constraint(model.N_T, rule=HPa_rec_rule)

"MI" # Needs to be defined separately for each HPa unit separately.
# Penalty for available waste heat that is not recovered with HPa [€] (Equation (15))
def C_Ql_HPa_rule(model, t):
    return model.p_Q * (model.HPa_rec_max[t] - model.Q_rec_HPa['HPa',t]) * model.dt
model.C_Ql_HPa = Expression(model.N_T, rule = C_Ql_HPa_rule)

# Heat generated via WHR is denoted with an ancillary variable "model.Q_HPa_up" (up = upgrade)
def HPa_up_rule(model,i,t):
    return model.COP_HPa[i,t] * model.Q_rec_HPa[i,t] / (model.COP_HPa[i,t] - 1)
model.Q_HPa_up = Expression(model.N_HPa, model.N_T, rule=HPa_up_rule)

# Maximum heat generation capacity of HPa units in ATW mode [MW] (Equation (16))
# ATW operation is possible only whenever outside air temperature exceeds the minimum threshold
def Q_ATW_rule(model,i,t):
    if model.T_0[t] >= model.T_min_ATW_HPa[i]:
        return model.Q_ATW_HPa[i,t] <= model.y_HPa[i,t] * model.Q_max_ATW_HPa_real[i,t]
    else:
        return model.Q_ATW_HPa[i,t] == 0
model.Q_ATW_constraint = Constraint(model.N_HPa, model.N_T, rule=Q_ATW_rule)

# Heat output of HPa units is equal the sum of WHR and ATW operation (Equation (17))
def Q_HPa_rule(model,i,t):
    return model.Q_HPa[i,t] == model.Q_HPa_up[i,t] + model.Q_ATW_HPa[i,t]
model.Q_HPa_constraint = Constraint(model.N_HPa, model.N_T, rule=Q_HPa_rule)

# Minimum heat generation capacity of HPa unit [MW] (Equation (3))
def Q_HPa_min_rule(model,i, t):
    return model.Q_HPa[i,t] >= model.y_HPa[i,t] * model.Q_min_HPa[i]
model.Q_HPa_min_constraint = Constraint(model.N_HPa, model.N_T, rule=Q_HPa_min_rule)

# Maximum heat generation capacity of HPa units [MW] (Equation (3))
def Q_HPa_max_rule(model,i,t):
    return model.Q_HPa[i,t] <= model.y_HPa[i,t] * model.Q_max_HPa[i]
model.Q_HPa_max_constraint = Constraint(model.N_HPa, model.N_T, rule=Q_HPa_max_rule)

# HPa unit start-up is tracked: model.y_su_HPa[i,t] gains value of 1 when model.y_HPa[i, t] switches from 0 to 1.
# (Equation (1))
def HPa_startup_rule(model, i, t):
    return model.y_su_HPa[i, t] >= model.y_HPa[i, t] - model.y_HPa[i, model.t_prev[t]]
model.HPa_startup_constraint = Constraint(model.N_HPa, model.N_T, rule=HPa_startup_rule)

# Start-up costs for HPa units [€] (Equation (2))
def C_su_HPa_rule(model, i, t):
    return model.y_su_HPa[i, t] * model.p_su_HPa[i]
model.C_su_HPa = Expression(model.N_HPa, model.N_T, rule=C_su_HPa_rule)

# Maximum ramp-up rate (Equation (4))
def ru_HPa_max_rule(model, i,t):
    return model.Q_HPa[i,t] - model.Q_HPa[i,model.t_prev[t]] <= model.Q_max_HPa[i] / model.rut_HPa[i] * model.dt
model.ru_HPa_max_constraint = Constraint(model.N_HPa, model.N_T, rule=ru_HPa_max_rule)

# Maximum ramp-down rate (Equation (5))
def rd_HPa_max_rule(model, i,t):
    return model.Q_HPa[i,model.t_prev[t]] - model.Q_HPa[i,t] <= model.Q_max_HPa[i] / model.rdt_HPa[i] * model.dt
model.rd_HPa_max_constraint = Constraint(model.N_HPa, model.N_T, rule=rd_HPa_max_rule)

# Electricity consumption costs for HPa units [€] (Equation (6))
def C_con_HPa_rule(model, i, t):
    return (model.p_E[t] + model.p_E_tr_buy) * (model.Q_ATW_HPa[i,t] / model.COP_ATW_HPa[i,t] \
                                                + model.Q_HPa_up[i,t] / model.COP_HPa[i,t]) * model.dt
model.C_con_HPa = Expression(model.N_HPa, model.N_T, rule=C_con_HPa_rule)

# COP reduction factor for HPa units (Equation (12))
def n_ru_a_rule(model,i,t):
    return (model.n_ru_HPa[i,t] >= model.n_ru_max_HPa[i] * model.rut_HPa[i] * (model.Q_HPa[i,t] \
                  - model.Q_HPa[i,model.t_prev[t]]) / (model.Q_max_HPa[i] * model.dt))
model.n_ru_a_constraint=Constraint(model.N_HPa, model.N_T, rule=n_ru_a_rule)

# Upper bound for the COP reduction factor for HPa units [-] (Equation (12))
def n_ru_max_a_rule(model,i,t):
    return model.n_ru_HPa[i,t] <= model.n_ru_max_HPa[i]
model.n_ru_a_max_constraint=Constraint(model.N_HPa, model.N_T, rule=n_ru_max_a_rule)

# Ramp-up costs for HPa units [€] (Equation (13))
def C_ru_HPa_rule(model, i, t):
    return (model.p_E[t] + model.p_E_tr_buy) * (1.16 * model.n_ru_max_HPa[i] + 0.97) / model.COP_HPa[i,t] \
                * model.Q_max_HPa[i] / 2 * model.n_ru_HPa[i,t] * model.dt
model.C_ru_HPa = Expression(model.N_HPa, model.N_T, rule=C_ru_HPa_rule)

# Total costs for each HPa unit [€] (Equation (50))
# Thermal loss costs for HPa units are added later to the objective function
def C_tot_HPa_rule(model, i,t):
    return model.C_con_HPa[i,t] + model.C_su_HPa[i,t] + model.C_ru_HPa[i,t]
model.C_tot_HPa = Expression(model.N_HPa, model.N_T, rule=C_tot_HPa_rule)

# ======================================================================================================================
# HOB
# Variables and parameters
# ======================================================================================================================

# Set HOB units
model.N_HOB = Set(initialize=id.N_HOB)

# Binary decision variables
# Status variables [-]
model.y_HOB = Var(model.N_HOB, model.N_T, domain=Binary)
# Start-up variables [-]
model.y_su_HOB = Var(model.N_HOB, model.N_T, domain=Binary)

# Continuous decision variables
# Heat generation [MW]
model.Q_HOB = Var(model.N_HOB, model.N_T, domain=NonNegativeReals)

# Parameters
# Fuel prices [€/MWh]
model.p_f_HOB = Param(model.N_HOB, model.N_T, initialize=id.p_f_HOB_flat)
# CO2 emission coefficients [kg CO2/MWh fuel]
model.CO2_f_HOB = Param(model.N_HOB, initialize=id.CO2_f_HOB)
# Maximum heat generation capacities [MW]
model.Q_max_HOB = Param(model.N_HOB, initialize=id.Q_max_HOB)
# Minimum heat generation capacities [MW]
model.Q_min_HOB = Param(model.N_HOB, initialize=id.Q_min_HOB)
# Heat generation efficiencies [-]
model.eta_Q_HOB = Param(model.N_HOB, initialize=id.eta_Q_HOB)
# Ramp-up times [h]
model.rut_HOB = Param(model.N_HOB, initialize=id.rut_HOB)
# Ramp-down times [h]
model.rdt_HOB = Param(model.N_HOB, initialize=id.rdt_HOB)
# Start-up prices [€/start]
model.p_su_HOB = Param(model.N_HOB, initialize=id.p_su_HOB)

# ======================================================================================================================
# HOB
# Constraints and expressions
# ======================================================================================================================

# HOB unit start-up is tracked: model.y_su_HOB[i,t] gains value of 1 when model.y_HOB[i, t] switches from 0 to 1.
# (Equation (1))
def HOB_su_rule(model,i,t):
    return model.y_su_HOB[i,t] >= model.y_HOB[i,t] - model.y_HOB[i,model.t_prev[t]]
model.HOB_su_constraint = Constraint(model.N_HOB, model.N_T, rule=HOB_su_rule)

# Start-up costs [€] (Equation (2))
def C_su_HOB_rule(model,i, t):
    return model.y_su_HOB[i,t] * model.p_su_HOB[i]
model.C_su_HOB = Expression(model.N_HOB, model.N_T, rule=C_su_HOB_rule)

# Maximum heat production capacity of the HOB units (Equation (3))
def Q_HOB_max_rule(model, i, t):
    return model.Q_HOB[i,t] <= model.y_HOB[i,t] * model.Q_max_HOB[i]
model.Q_HOB_max_constraint = Constraint(model.N_HOB, model.N_T, rule=Q_HOB_max_rule)

# Minimum heat production capacity of the HOB units (Equation (3))
def Q_HOB_min_rule(model, i, t):
    return model.Q_HOB[i,t] >= model.y_HOB[i,t] * model.Q_min_HOB[i]
model.Q_HOB_min_constraint = Constraint(model.N_HOB, model.N_T, rule=Q_HOB_min_rule)

# Ramp-up rate (Equation (4))
def ru_HOB_max_rule(model, i,t):
    return model.Q_HOB[i,t] - model.Q_HOB[i,model.t_prev[t]] <= model.Q_max_HOB[i] / model.rut_HOB[i] * model.dt
model.ru_HOB_max_constraint = Constraint(model.N_HOB, model.N_T, rule=ru_HOB_max_rule)

# Ramp-down rate (Equation (5))
def rd_HOB_max_rule(model, i, t):
    return model.Q_HOB[i, model.t_prev[t]] - model.Q_HOB[i,t] <= model.Q_max_HOB[i] / model.rdt_HOB[i] * model.dt
model.rd_HOB_max_constraint = Constraint(model.N_HOB, model.N_T, rule=rd_HOB_max_rule)

# HOB fuel consumption costs [€] (Equation (45))
def C_con_HOB_rule(model, i, t):
    return model.p_f_HOB[i,t] * model.Q_HOB[i,t] / model.eta_Q_HOB[i] * model.dt
model.C_con_HOB = Expression(model.N_HOB, model.N_T, rule=C_con_HOB_rule)

# Total costs for each HOB unit [€] (Equation (50))
def C_tot_HOB_rule(model, i,t):
    return model.C_con_HOB[i,t] + model.C_su_HOB[i,t]
model.C_tot_HOB = Expression(model.N_HOB, model.N_T, rule=C_tot_HOB_rule)

# ======================================================================================================================
# TS
# Variables and parameters
# ======================================================================================================================

# Set of TS units
model.N_TS = Set(initialize=id.N_TS)

# Binary decision variables
# Status variables for charging [-]
model.y_in_TS = Var(model.N_TS, model.N_T, domain=Binary)
# Status variables for discharging [-]
model.y_out_TS = Var(model.N_TS, model.N_T, domain=Binary)
# Start-up variables for charging [-]
model.y_in_su_TS = Var(model.N_TS, model.N_T, domain=Binary)
# Start-up variables for discharging [-]
model.y_out_su_TS = Var(model.N_TS, model.N_T, domain=Binary)
# Status variables for heat transferred from HPa => TS unit for each possible pairing [-]
model.y_in_HP_TS = Var(model.N_TS, model.N_HPa, model.N_T, domain=Binary)
# Status variables for heat transferred from HPa => TS units [-]
model.y_in_HP_tot_TS = Var(model.N_TS, model.N_T, domain=Binary)

# Continuous decision variables
# Heat charged into a TS unit [MW]
model.Q_in_TS = Var(model.N_TS, model.N_T, domain=NonNegativeReals)
# Heat discharged from a TS unit [MW]
model.Q_out_TS = Var(model.N_TS, model.N_T, domain=NonNegativeReals)
# Cumulative stored heat in a TS unit [MWh]
model.Q_TS = Var(model.N_TS, model.N_T, domain=NonNegativeReals)
# Heat transferred from HPa unit to a TS unit for each possible pairing [MW]
model.Q_in_HP_TS = Var(model.N_TS, model.N_HPa, model.N_T, domain=NonNegativeReals)
# Heat transferred from HPa units to TS units [MW]
model.Q_in_HP_tot_TS = Var(model.N_TS, model.N_T, domain=NonNegativeReals)

# Parameters
# Maximum storage capacities [MWh]
model.Q_max_TS = Param(model.N_TS, initialize=id.Q_max_TS)
# Maximum allowed storage levels [-]
model.lvl_max_TS = Param(model.N_TS, initialize=id.lvl_max_TS)
# Minimum allowed storage levels [-]
model.lvl_min_TS = Param(model.N_TS, initialize=id.lvl_min_TS)
# Efficiencies of heat transfer into a TS unit (charging) [-]
model.eta_in_TS = Param(model.N_TS, initialize=id.eta_in_TS)
# Efficiencies of heat transfer out of a TS unit (discharging) [-]
model.eta_out_TS = Param(model.N_TS, initialize=id.eta_out_TS)
# Maximum charge rate [MW]
model.Q_in_max_TS = Param(model.N_TS, initialize=id.Q_in_max_TS)
# Maximum discharge rate [MW]
model.Q_out_max_TS = Param(model.N_TS, initialize=id.Q_out_max_TS)
# Minimum charge rate [MW]
model.Q_in_min_TS = Param(model.N_TS, initialize=id.Q_in_min_TS)
# Minimum discharge rate [MW]
model.Q_out_min_TS = Param(model.N_TS, initialize=id.Q_out_min_TS)
# Price for starting the charging action [€]
model.p_in_su_TS = Param(model.N_TS, initialize=id.p_in_su_TS)
# Price for starting the discharging action [€]
model.p_out_su_TS = Param(model.N_TS, initialize=id.p_out_su_TS)
# Losses through the lateral surface of the high-temperature section of the water body [-]
model.beta_TS = Param(model.N_TS, initialize=id.beta_TS)
# Losses through the total lateral surface assuming the storage is empty [-]
model.gamma_TS = Param(model.N_TS, model.N_T, initialize=id.gamma_TS_flat)
# Losses through the bottom and top surfaces [MWh]
model.delta_TS = Param(model.N_TS, model.N_T, initialize=id.delta_TS_flat)
# Hot zone temperature [C]
model.T_H_TS = Param(model.N_TS, initialize=id.T_H_TS)
# Cold zone temperature [C]
model.T_C_TS = Param(model.N_TS, initialize=id.T_C_TS)
# New cold zone temperature [C]
model.T_C_new_TS = Param(model.N_TS, initialize=id.T_C_new_TS)

# ======================================================================================================================
# HPb: HPs integrated with TS units
# Variables and parameters
# ======================================================================================================================

# Set of HPb units
model.N_HPb = Set(initialize=id.N_HPb)

# Binary decision variables
# Status variables for HPb units [-]
model.y_HPb = Var(model.N_HPb, model.N_T, domain=Binary)
# Start-up variables [-]
model.y_su_HPb = Var(model.N_HPb, model.N_T, domain=Binary)


# Continuous decision variables
# Recovered waste heat [MW]
model.Q_rec_HPb = Var(model.N_HPb, model.N_T, domain=NonNegativeReals)
# Generated heat [MW]
model.Q_HPb = Var(model.N_HPb, model.N_T, domain=NonNegativeReals)
# COP reduction factor during ramp-up [-]
model.n_ru_HPb = Var(model.N_HPb, model.N_T, domain=NonNegativeReals)

# Parameters
# Start-up prices [€]
model.p_su_HPb = Param(model.N_HPb, initialize=id.p_su_HPb)
# Maximum heat generation capacities [MW]
model.Q_max_HPb = Param(model.N_HPb, initialize=id.Q_max_HPb)
# minimum heat generation capacities [MW]
model.Q_min_HPb = Param(model.N_HPb, initialize=id.Q_min_HPb)
# COP [-]
model.COP_HPb = Param(model.N_HPb, model.N_T, initialize=id.COP_HPb_flat)
# Maximum value for COP reduction factor during ramp-up [-]
model.n_ru_max_HPb = Param(model.N_HPb, initialize=id.n_ru_max_HPb)
# Ramp-up times [h]
model.rut_HPb = Param(model.N_HPb, initialize=id.rut_HPb)
# Ramp-down times [h]
model.rdt_HPb = Param(model.N_HPb, initialize=id.rdt_HPb)

# ======================================================================================================================
# HPb: HPs integrated with TS units
# Constraints and expressions
# ======================================================================================================================

"MI" # Heat recovery for each HPb-TS pair unit needs to be specified individually
# HPb heat recovery from TS return water (Equation (35)).
def Q_rec_HPb_rule(model, t):
    if model.Q_max_HPb['HPb'] > 0:
        return model.Q_rec_HPb['HPb',t] == model.Q_out_TS['TS',t] * (model.T_C_TS['TS'] - model.T_C_new_TS['TS']) \
            / (model.T_H_TS['TS'] - model.T_C_TS['TS'])
    else:
        return Constraint.Skip
model.Q_rec_HPb_constraint = Constraint(model.N_T, rule=Q_rec_HPb_rule)

# HPb heat generation [MW] (Equation (36))
def Q_HPb_rule(model, i, t):
    if model.Q_max_HPb[i] > 0:
        return model.Q_HPb[i,t] == model.COP_HPb[i,t] * model.Q_rec_HPb[i,t] / (model.COP_HPb[i,t] - 1)
    else:
        return model.Q_HPb[i,t] == 0
model.Q_HPb_constraint = Constraint(model.N_HPb, model.N_T, rule=Q_HPb_rule)

# HP unit start-up is tracked: model.y_su_HPb[i,t] gains value of 1 when model.y_HPb[i, t] switches from 0 to 1.
# (Equation (1))
def HPb_startup_rule(model, i, t):
    if model.Q_max_HPb[i] > 0:
        return model.y_su_HPb[i, t] >= model.y_HPb[i, t] - model.y_HPb[i, model.t_prev[t]]
    else:
        return model.y_HPb[i, t] + model.y_su_HPb[i, t] == 0
model.HPb_startup_constraint = Constraint(model.N_HPb, model.N_T, rule=HPb_startup_rule)

# Start-up costs [€] (Equation (2))
def C_su_HPb_rule(model, i, t):
    return model.y_su_HPb[i, t] * model.p_su_HPb[i]
model.C_su_HPb = Expression(model.N_HPb, model.N_T, rule=C_su_HPb_rule)

# Maximum heat generation capacities [MW] (Equation (3))
def Q_HPb_max_rule(model,i, t):
    if model.Q_max_HPb[i] > 0:
        return model.Q_HPb[i,t] <= model.y_HPb[i,t] * model.Q_max_HPb[i]
    else:
        return Constraint.Skip
model.Q_HPb_max_constraint = Constraint(model.N_HPb, model.N_T, rule=Q_HPb_max_rule)

# Minimum heat generation capacities [MW] (Equation (3))
def Q_HPb_min_rule(model,i, t):
    if model.Q_max_HPb[i] > 0:
        return model.Q_HPb[i,t] >= model.y_HPb[i,t] * model.Q_min_HPb[i]
    else:
        return Constraint.Skip
model.Q_HPb_min_constraint = Constraint(model.N_HPb, model.N_T, rule=Q_HPb_min_rule)

# Ramp up rates (Equation (4))
def ru_HPb_max_rule(model, i,t):
    return model.Q_HPb[i,t] - model.Q_HPb[i,model.t_prev[t]] <= model.Q_max_HPb[i] / model.rut_HPb[i] * model.dt
model.ru_HPb_max_constraint = Constraint(model.N_HPb, model.N_T, rule=ru_HPb_max_rule)

# Ramp-down rates (Equation (5))
def rd_HPb_max_rule(model, i,t):
    return model.Q_HPb[i,model.t_prev[t]] - model.Q_HPb[i,t] <= model.Q_max_HPb[i] / model.rdt_HPb[i] * model.dt
model.rd_HPb_max_constraint = Constraint(model.N_HPb, model.N_T, rule=rd_HPb_max_rule)

# Electricity consumption costs [€] (Equation (6))
def C_con_HPb_rule(model, i, t):
    return (model.p_E[t] + model.p_E_tr_buy) * model.Q_HPb[i,t] / model.COP_HPb[i,t] * model.dt
model.C_con_HPb = Expression(model.N_HPb, model.N_T, rule=C_con_HPb_rule)

# COP reduction factor (Equation (12))
def n_ru_b_rule(model,i,t):
    if model.Q_max_HPb[i] > 0:
        return model.n_ru_HPb[i,t] >= model.n_ru_max_HPb[i] * model.rut_HPb[i] * (model.Q_HPb[i,t] \
                          - model.Q_HPb[i,model.t_prev[t]]) / (model.Q_max_HPb[i] * model.dt)
    else:
        return model.n_ru_HPb[i, t] == 0
model.n_ru_b_constraint = Constraint(model.N_HPb, model.N_T, rule=n_ru_b_rule)

# Upper bound for the COP reduction factor [-] (Equation (12))
def n_ru_max_b_rule(model,i,t):
    if model.Q_max_HPb[i] > 0:
        return model.n_ru_HPb[i,t] <= model.n_ru_max_HPb[i]
    else:
        return Constraint.Skip
model.n_ru_b_max_constraint = Constraint(model.N_HPb, model.N_T, rule=n_ru_max_b_rule)

# HPb ramp-up costs [€] (Equation (13))
def C_ru_HPb_rule(model, i, t):
    return (model.p_E[t] + model.p_E_tr_buy) * (1.16 * model.n_ru_max_HPb[i] + 0.97) / model.COP_HPb[i,t] \
                * model.Q_max_HPb[i] / 2 * model.n_ru_HPb[i,t] * model.dt
model.C_ru_HPb = Expression(model.N_HPb, model.N_T, rule=C_ru_HPb_rule)

# Total costs for each HPb unit [€] (Equation (50))
def C_tot_HPb_rule(model, i,t):
    return model.C_con_HPb[i,t] + model.C_su_HPb[i,t] + model.C_ru_HPb[i,t]
model.C_tot_HPb = Expression(model.N_HPb, model.N_T, rule=C_tot_HPb_rule)

# ======================================================================================================================
# HPa interaction with TS
# Constraints and expressions
# ======================================================================================================================

# Start-up tracking for charging of TS units (Equation (20))
def TS_in_startup_rule(model, i, t):
    return model.y_in_su_TS[i, t] >= 0.25 * (model.y_in_TS[i, t] + model.y_in_HP_tot_TS[i,t]) \
        - 0.5 * (model.y_in_TS[i, model.t_prev[t]] + model.y_in_HP_tot_TS[i, model.t_prev[t]])
model.TS_in_startup_constraint = Constraint(model.N_TS, model.N_T, rule=TS_in_startup_rule)

# Check HPa sink outlet temperature to check if it can charge a TS unit alone (Equation (21))
def HP_charge_TS_permission_rule(model, i, j, t):
    if model.T_sink_out_HPa[j,t] < model.T_H_TS[i]:
        # HPa colder than TS → only allowed if other sources are charging
        return model.y_in_HP_TS[i,j,t] <= model.y_in_TS[i,t]
    else:
        # HPa hot enough → always allowed
        return Constraint.Skip
model.HP_charge_TS_permission_constraint \
    = Constraint(model.N_TS, model.N_HPa, model.N_T, rule=HP_charge_TS_permission_rule)

# TS unit charging with HPa units is active if any of the HPa units is charging the TS unit (Equation (22))
def y_in_HP_tot_rule(model, i, t):
   return model.y_in_HP_tot_TS[i,t] <= sum(model.y_in_HP_TS[i,j,t] for j in model.N_HPa)
model.Q_HP_in_constraint = Constraint(model.N_TS, model.N_T, rule= y_in_HP_tot_rule)

# Total heat flow from HPa units to each TS unit (Equation (23))
def Q_HP_in_tot_rule(model, i, t):
    return model.Q_in_HP_tot_TS[i,t] == sum(model.Q_in_HP_TS[i,j,t] for j in model.N_HPa)
model.Q_HP_in_tot_constraint = Constraint(model.N_TS, model.N_T, rule=Q_HP_in_tot_rule)

# Heat flow from a HPa unit to TS units is limited by HPa heat generation (Equation (24))
def Q_HP_TS_availability_rule(model, j, t):
    return sum(model.Q_in_HP_TS[i,j,t] for i in model.N_TS) <= model.Q_HPa[j,t]
model.Q_HP_TS_availability_constraint = Constraint( model.N_HPa, model.N_T, rule=Q_HP_TS_availability_rule)

# Heat flow from HPa to TS is limited by HPa maximum capacity (Equation (25))
def Q_in_HP_max_rule(model, i, j, t):
    return model.Q_in_HP_TS[i,j,t] <= model.Q_max_HPa[j] * model.y_in_HP_TS[i,j,t]
model.Q_in_HP_max_constraint = Constraint(model.N_TS, model.N_HPa, model.N_T, rule=Q_in_HP_max_rule)

# Heat flow from HPa to TS is limited by HPa minimum capacity (Equation (25))
def Q_in_HP_min_rule(model, i, j, t):
    return model.Q_in_HP_TS[i,j,t] >= model.Q_min_HPa[j] * model.y_in_HP_TS[i,j,t]
model.Q_in_HP_min_constraint = Constraint(model.N_TS, model.N_HPa, model.N_T, rule=Q_in_HP_min_rule)

# Maximum charge rates with heat from HPa units (Equation (26))
def Q_HP_in_tot_max_rule(model, i, t):
    return model.Q_in_HP_tot_TS[i,t] <= model.y_in_HP_tot_TS[i,t] * model.Q_in_max_TS[i]
model.Q_HP_in_tot_max_constraint = Constraint(model.N_TS, model.N_T, rule=Q_HP_in_tot_max_rule)

# Minimum charge rates with heat from HPa units (Equation (26))
def Q_HP_in_tot_min_rule(model, i, t):
    return model.Q_in_HP_tot_TS[i,t] >= model.y_in_HP_tot_TS[i,t] * model.Q_in_min_TS[i]
model.Q_HP_in_tot_min_constraint = Constraint(model.N_TS, model.N_T, rule=Q_HP_in_tot_min_rule)

# Maximum charge rates from all sources (Equation (29))
def Q_in_TS_tot_max_rule(model, i, t):
    return model.Q_in_TS[i,t] + model.Q_in_HP_tot_TS[i,t] <= model.Q_in_max_TS[i]
model.Q_TS_in_tot_max_constraint = Constraint(model.N_TS, model.N_T, rule=Q_in_TS_tot_max_rule)

# Simultaneous charging with HPa units and discharging are not possible (Equation (39))
def Q_TS_in_HP_and_out_rule(model,i,t):
    epsilon = 0.001
    return model.y_out_TS[i,t] + epsilon * sum(model.y_in_HP_TS[i,j,t] for j in model.N_HPa) <= 1
model.Q_TS_in_HP_and_out_constraint = Constraint(model.N_TS, model.N_T, rule=Q_TS_in_HP_and_out_rule)

# ======================================================================================================================
# TS
# Constraints and expressions
# ======================================================================================================================

# Cumulative stored heat in TS units ( Equation (18))
"MI" # Needs to be specified for each unit separately due possibility of HPb units
# The initial state for the level of charge in the TS can be specified if desired.
# Otherwise the model picks the optimal initial state
def Q_TS_rule(model,t):
    #if t == 1:
     #   return model.Q_TS['TS',t] == model.Q_max_TS['TS'] * model.lvl_min_TS['TS']
    #else:
        return model.Q_TS['TS',t] == model.Q_TS['TS',model.t_prev[t]] * (1 - model.beta_TS['TS']) \
            - model.gamma_TS['TS',t] * model.Q_max_TS['TS'] - model.delta_TS['TS',t] \
            + (model.Q_in_TS['TS',t] + model.Q_in_HP_tot_TS['TS',t]) * model.dt * model.eta_in_TS['TS'] \
            - (model.Q_out_TS['TS',t] + model.Q_rec_HPb['HPb',t]) * model.dt / model.eta_out_TS['TS']
model.Q_TS_constraint = Constraint(model.N_T, rule=Q_TS_rule)

# Minimum storage capacity for TS units (Equation (19))
def Q_min_capacity_TS_rule(model, i, t):
    return model.Q_TS[i,t] >= model.Q_max_TS[i] * model.lvl_min_TS[i]
model.Q_TS_min_capacity_constraint = Constraint(model.N_TS, model.N_T, rule = Q_min_capacity_TS_rule)

# Maximum storage capacity fot TS units (Equation (19))
def Q_max_capacity_TS_rule(model, i, t):
    return model.Q_TS[i,t] <= model.Q_max_TS[i] * model.lvl_max_TS[i]
model.Q_TS_max_capacity_constraint = Constraint(model.N_TS, model.N_T, rule = Q_max_capacity_TS_rule)

# Heat availability for charging through CHP and HOB (Equation (27))
def Q_TS_in_availability_rule(model,t):
    return sum(model.Q_in_TS[i,t] for i in model.N_TS) <= sum(model.Q_CHPa[i,t] for i in model.N_CHPa) \
        + sum(model.Q_HOB[i,t] for i in model.N_HOB)
model.Q_TS_in_availability_constraint = Constraint(model.N_T, rule=Q_TS_in_availability_rule)

# Maximum charge rates with heat from CHP and HOB units (Equation (28))
def Q_in_max_TS_rule(model, i, t):
    return model.Q_in_TS[i,t] <= model.y_in_TS[i,t] * model.Q_in_max_TS[i]
model.Q_in_max_TS_constraint = Constraint(model.N_TS, model.N_T, rule=Q_in_max_TS_rule)

# Minimum charge rates (Equation (28))
def Q_TS_in_min_rule(model,i,t):
    return model.Q_in_TS[i,t] >= model.y_in_TS[i,t] * model.Q_in_min_TS[i]
model.TS_in_min_constraint = Constraint(model.N_TS, model.N_T, rule=Q_TS_in_min_rule)

"MI" # Needs to specified for each TS unit separately due to possible inclusion on a HPb unit
# Heat availability for discharging (Equation (30))
def Q_TS_out_availability_rule(model,t):
    return (model.Q_out_TS['TS',t] + model.Q_rec_HPb['HPb',t]) / model.eta_out_TS['TS'] * model.dt <= model.Q_TS['TS',t]
model.Q_TS_out_availability_constraint = Constraint(model.N_T, rule=Q_TS_out_availability_rule)

# Maximum discharge rates (Equation (31))
def Q_out_max_TS_rule(model, i, t):
    return model.Q_out_TS[i,t] / model.eta_out_TS[i] <= model.y_out_TS[i,t] * model.Q_out_max_TS[i]
model.Q_out_max_Ts_constraint = Constraint(model.N_TS, model.N_T, rule=Q_out_max_TS_rule)

# Minimum discharge rates (Equation (31))
def Q_TS_out_min_rule(model,i,t):
    return model.Q_out_TS[i,t] / model.eta_out_TS[i] >= model.y_out_TS[i,t] * model.Q_out_min_TS[i]
model.TS_out_min_constraint = Constraint(model.N_TS, model.N_T, rule=Q_TS_out_min_rule)

# Start-up tracking for discharging of TS units  (Equation (32))
def TS_out_startup_rule(model, i, t):
    return model.y_out_su_TS[i, t] >= model.y_out_TS[i, t] - model.y_out_TS[i, model.t_prev[t]]
model.TS_out_startup_constraint = Constraint(model.N_TS, model.N_T, rule=TS_out_startup_rule)

# Simultaneous charging and discharging are not possible (Equation (38))
def Q_TS_in_and_out_rule(model,i,t):
    return model.y_in_TS[i,t] + model.y_out_TS[i,t] <= 1
model.Q_TS_in_and_out_constraint = Constraint(model.N_TS, model.N_T, rule=Q_TS_in_and_out_rule)

# Start-up costs for TS units [€] (Equation (40))
def C_TS_in_out_su_rule(model, i, t):
    return model.y_in_su_TS[i,t] * model.p_in_su_TS[i] + model.y_out_su_TS[i,t] * model.p_out_su_TS[i]
model.C_TS_in_out_su = Expression(model.N_TS, model.N_T, rule=C_TS_in_out_su_rule)

# Ancillary variable for heat leaking out of TS units while they are idle [MWh]
def Q_TS_leak_rule(model,i,t):
    return (model.Q_TS[i,model.t_prev[t]] * model.beta_TS[i] + model.gamma_TS[i,t] * model.Q_max_TS[i] \
           + model.delta_TS[i,t])
model.Q_TS_leak = Expression(model.N_TS, model.N_T, rule=Q_TS_leak_rule)

"MI" # Need to be defined separately for each TS unit. Thermal loss costs for TS units [€] (Equation (44))
def C_Ql_TS_rule(model, t):
    return model.p_Q * (model.Q_TS_leak['TS',t] + ((1 - model.eta_in_TS['TS']) \
                     * (model.Q_in_TS['TS',t] + model.Q_in_HP_tot_TS['TS',t]) + (1 / model.eta_out_TS['TS'] - 1) \
                       * (model.Q_out_TS['TS',t] + model.Q_rec_HPb['HPb',t])) * model.dt)
model.C_Ql_TS = Expression( model.N_T, rule=C_Ql_TS_rule)

# ======================================================================================================================
# Thermal demand
# ======================================================================================================================

# Thermal demand [MW]
model.Q_d = Param(model.N_T, initialize={i: id.Q_d[i - 1] for i in id.T})

# Thermal demand is perfectly fulfilled at each timestep (Equation (48))
def Q_d_rule(model, t):
    return sum(model.Q_HPa[i,t] for i in model.N_HPa) + (sum(model.Q_CHPa[i,t] for i in model.N_CHPa)) \
        + sum(model.Q_HPb[i, t] for i in model.N_HPb)\
        + sum(model.Q_HOB[i,t] for i in model.N_HOB) - (sum((model.Q_in_TS[i,t] + \
                                                             model.Q_in_HP_tot_TS[i,t]) for i in model.N_TS)) \
        + sum(model.Q_out_TS[i,t] for i in model.N_TS) == model.Q_d[t]
model.Q_d_constraint = Constraint(model.N_T, rule=Q_d_rule)

# ======================================================================================================================
# CO2 emissions
# ======================================================================================================================

# CO2 emissions from bought electricity [kg CO2/MWh]
model.omega = Param(initialize=id.omega)

# Total CO2 emissions [kg CO2] (Equation (49))
def CO2_rule(model,t):
    return  (sum(model.CO2_f_CHPa[i] * model.Q_CHPa[i,t] / model.eta_Q_CHPa[i] for i in model.N_CHPa) \
             + sum(model.CO2_f_CHPb[i] * model.Q_CHPb[i,t] / model.eta_Q_CHPb[i] for i in model.N_CHPb) \
             + sum(model.CO2_f_HOB[i] * model.Q_HOB[i,t] / model.eta_Q_HOB[i] for i in model.N_HOB)) * model.dt \
                + (sum(model.Q_HPa_up[i,t] / model.COP_HPa[i,t] + \
                       model.Q_ATW_HPa[i,t] / model.COP_ATW_HPa[i,t] for i in model.N_HPa) \
                   + sum(model.Q_HPb[i,t] / model.COP_HPb[i,t] for i in model.N_HPb)) * model.omega * model.dt
model.CO2 = Expression(model.N_T, rule=CO2_rule)

# ======================================================================================================================
# Objective function: minimize operation costs and maximize electricity trading
# ======================================================================================================================

# Net operation costs [€]
def C_net_rule(model,t):
    return sum(model.C_tot_HPa[i,t] for i in model.N_HPa) \
        + model.C_Ql_HPa[t] \
        + sum(model.C_tot_HPb[i,t] for i in model.N_HPb) \
        + sum(model.C_tot_HOB[i,t] for i in model.N_HOB) \
        + sum(model.C_tot_CHPa[i,t] for i in model.N_CHPa) \
        + sum(model.C_tot_CHPb[i,t] for i in model.N_CHPb) \
        + model.C_Ql_TS[t] \
        + sum(model.C_TS_in_out_su[i,t] for i in model.N_TS) \
        - sum(model.R_E_a[i,t] for i in model.N_CHPa) \
        - sum(model.R_E_b[i,t] for i in model.N_CHPb)
model.C_net = Expression(model.N_T, rule=C_net_rule)

# Objective function (Equation (45))
def model_objective_rule(model):
    return sum(
        sum(model.C_tot_HPa[i, t] for i in model.N_HPa) \
        + model.C_Ql_HPa[t] \
        + sum(model.C_tot_HPb[i, t] for i in model.N_HPb) \
        + sum(model.C_tot_HOB[i, t] for i in model.N_HOB) \
        + sum(model.C_tot_CHPa[i, t] for i in model.N_CHPa) \
        + sum(model.C_tot_CHPb[i, t] for i in model.N_CHPb) \
        + model.C_Ql_TS[t] \
        + sum(model.C_TS_in_out_su[i, t] for i in model.N_TS) \
        - sum(model.R_E_a[i, t] for i in model.N_CHPa) \
        - sum(model.R_E_b[i, t] for i in model.N_CHPb)
        for t in model.N_T)
model.obj = Objective(rule=model_objective_rule, sense=minimize)

# ======================================================================================================================
# Solving the optimization problem
# ======================================================================================================================

# Picking "gurobi" as the solver and solving the model
solver = SolverFactory('gurobi')
results = solver.solve(model, tee=True)

# Check solver status
print(f"Solver Status: {results.solver.status}")
print(f"Solver Termination Condition: {results.solver.termination_condition}")

# If the solver was successfully, print the results
if results.solver.status == SolverStatus.ok and results.solver.termination_condition == TerminationCondition.optimal:
    print("Optimal solution found:")
    # Print out the total revenue
    print(f"Net operation costs: {model.obj()} €, Total CO2 emissions: {sum(model.CO2[t]() for t in model.N_T)} kg")
else:
    print("Solver did not find an optimal solution. Check model formulation or constraints.")


# ======================================================================================================================
# Exporting the simulation results to an Excel file
# ======================================================================================================================

# Create a list of simulation results for each timestep
# Timesteps
results = {'Time': list(model.N_T)}
# Time in hours
results[f'Time_in_h'] = [t * id.dt for t in model.N_T]
for i in model.N_CHPa:
    # Heat generation of each CHPa unit [MW]
    results[f'Q_{i} [MW]'] = [model.Q_CHPa[i, t].value for t in model.N_T]
    # Electricity generation of each CHPa unit [MW]
    results[f'E_{i} [MW]'] = [model.E_CHPa[i, t].value for t in model.N_T]
for i in model.N_CHPb:
    # Electricity generation of each CHPb unit [MW]
    results[f'E_{i} [MW]'] = [model.E_CHPb[i, t].value for t in model.N_T]
# Total fuel consumption costs of all CHP units [€]
results[f'C_CHP_con [€]'] = [sum(model.C_con_CHPa[i,t]() for i in model.N_CHPa) \
                             + sum(model.C_con_CHPb[i,t]() for i in model.N_CHPb) for t in model.N_T]
# Total start-up costs of all CHP units [€]
results[f'C_CHP_su [€]'] = [sum(model.C_su_CHPa[i,t]() for i in model.N_CHPa) \
                            + sum(model.C_su_CHPb[i,t]() for i in model.N_CHPb) for t in model.N_T]
# Total heat production of all CHPa units [MW]
results[f'Q_CHPa_tot [MW]'] = [sum(model.Q_CHPa[i, t].value for i in model.N_CHPa) for t in model.N_T]
# Total electricity production of all CHP units [MW]
results[f'E_CHP_tot [MW]'] = [sum(model.E_CHPa[i, t].value for i in model.N_CHPa) + \
                              sum(model.E_CHPb[i, t].value for i in model.N_CHPb) for t in model.N_T]
for i in model.N_HPa:
    # Heat output of each HPa unit through WHR [MW]
    results[f'Q_up_{i} [MW]'] = [model.Q_HPa_up[i, t]() for t in model.N_T]
    # COP of each HPa unit [-]
    results[f'COP_{i} [-]'] = [model.COP_HPa[i, t] for t in model.N_T]
    # Heat output of each HPa unit in ATW mode [MW]
    results[f'Q_ATW_{i} [MW]'] = [model.Q_ATW_HPa[i, t]() for t in model.N_T]
    # COP in ATW mode of each HPa unit [-]
    results[f'COP_ATW_{i} [-]'] = [model.COP_ATW_HPa[i, t] for t in model.N_T]
for i in model.N_HPb:
    # Heat output of each HPb unit[MW]
    results[f'Q_{i} [MW]'] = [model.Q_HPb[i, t].value for t in model.N_T]
    # COP of each HPb unit [-]
    results[f'COP_{i} [-]'] = [model.COP_HPb[i, t] for t in model.N_T]
# Total electricity consumption costs of all HPs [€]
results[f'C_HP_con [€]'] = [sum(model.C_con_HPa[i,t]() for i in model.N_HPa) + \
                            sum(model.C_con_HPb[i,t]() for i in model.N_HPb) for t in model.N_T]
# Total ramp-up costs of all HPs [€]
results[f'C_HP_ru [€]'] = [sum(model.C_ru_HPa[i,t]() for i in model.N_HPa) + \
                           sum(model.C_ru_HPb[i,t]() for i in model.N_HPb) for t in model.N_T]
# Total start-up costs of all HPs [€]
results[f'C_HP_su [€]'] = [sum(model.C_su_HPa[i,t]() for i in model.N_HPa) + \
                           sum(model.C_su_HPb[i,t]() for i in model.N_HPb) for t in model.N_T]
# Total electricity input of all HPs [MW]
results[f'E_HP_tot [MW]'] = [sum(model.Q_HPa_up[i,t]() / model.COP_HPa[i,t] + \
                                 model.Q_ATW_HPa[i,t]() / model.COP_ATW_HPa[i,t] for i in model.N_HPa) + \
                             sum(model.Q_HPb[i,t].value / model.COP_HPb[i,t] for i in model.N_HPb) for t in model.N_T]
for i in model.N_HOB:
    # Heat output of each HOB unit [MW]
    results[f'Q_{i} [MW]'] = [model.Q_HOB[i, t].value for t in model.N_T]
# Fuel consumption costs of each HOB unit [€]
results[f'C_HOB_con [€]'] = [sum(model.C_con_HOB[i,t]() for i in model.N_HOB) for t in model.N_T]
# Start-up costs of each HOB unit [€]
results[f'C_su_HOB [€]'] = [sum(model.C_su_HOB[i,t]() for i in model.N_HOB) for t in model.N_T]
for i in model.N_TS:
    # Heat charged directly from HPa units to TS units [MW]
    results[f'Q_HP_in_{i} [MW]'] = [model.Q_in_HP_tot_TS[i, t].value for t in model.N_T]
    # Heat charged into the TS units from CHPa and HOB units [MW]
    results[f'Q_in_{i} [MW]'] = [model.Q_in_TS[i, t].value for t in model.N_T]
    # Heat discharged from the TS units [MW]
    results[f'Q_out_{i} [MW]'] = [model.Q_out_TS[i, t].value for t in model.N_T]
    # Cumulative stored heat in the TS units [MWh]
    results[f'Q_{i} [MWh]'] = [model.Q_TS[i, t].value for t in model.N_T]
# Start-up costs related to TS charging and discharging [€]
results[f'C_in_out_su_TS [€]'] = [sum(model.C_TS_in_out_su[i,t]() for i in model.N_TS) for t in model.N_T]
# Revenue gained from selling electricity produced via the CHP units [€]
results[f'R_E [€]'] = [sum(model.R_E_a[i,t]() for i in model.N_CHPa) \
                       + sum(model.R_E_b[i,t]() for i in model.N_CHPb) for t in model.N_T]
# Cost of bought electricity to operate all the HPs [€]
results[f'C_E_bought [€]'] = [(sum(model.Q_HPa_up[i,t]() / model.COP_HPa[i,t] + model.Q_ATW_HPa[i,t]() \
                                   / model.COP_ATW_HPa[i,t] for i in model.N_HPa) \
                                     + sum(model.Q_HPb[i,t].value / model.COP_HPb[i,t] for i in model.N_HPb)) \
                                         * (model.p_E[t] + model.p_E_tr_buy) * model.dt for t in model.N_T]
# Total produced CO2 emissions [kg]
results[f'CO2 [kg]'] = [model.CO2[t]() for t in model.N_T]
# Cost of heat losses in the TS [€]
results[f'C_Ql_TS [€]'] = [model.C_Ql_TS[t]() for t in model.N_T]
# Heat demand [MW]
results[f'Q_d [MW]'] = [model.Q_d[t] for t in model.N_T]
# Electricity price [€/MWh]
results[f'p_E [€/MWh]'] = [model.p_E[t] for t in model.N_T]
# Outside air temperature [C]
results[f'T_0 [C]'] = [model.T_0[t] for t in model.N_T]
# Net operation costs (model objective) [€]
results[f'C_net [€]'] = [model.C_net[t]() for t in model.N_T]

# Create DataFrame
df = pandas.DataFrame(results)

# Start with an empty row
totals = {}

"MI" # Columns with pure sums not related to the length of the timestep i.e. costs and emissions
# Specify names of result columns that have calculated sums
sum_cols = ['R_E [€]', 'CO2 [kg]', 'C_net [€]','C_E_bought [€]']

"MI" # Columns with sums that are related to the length of the timestep.
# Multiplying power dispatch with the timestep results in energy, MW * h = MWh
scaled_cols = ['E_CHP_tot [MW]', 'E_HP_tot [MW]']

# Fill totals
for col in sum_cols:
    totals[col] = df[col].sum()

for col in scaled_cols:
    totals[col] = df[col].sum() * model.dt

# Append the totals row
df.loc['Total'] = totals

# Save to Excel
df.to_excel('Simulation_results.xlsx', index=False, engine='openpyxl')


