# -*- coding: utf-8 -*-
"""
Created on Fri Dec  2 21:47:06 2022

@author: Dev
"""
from oemof.solph import (Sink, Source, Transformer, Bus, Flow, Model,
                         EnergySystem, processing, views)
#from oemof.solph import EnergySystem

from oemof.thermal.compression_heatpumps_and_chillers import calc_cops
from oemof.tools import economics
from oemof import solph

import pandas as pd

my_index = pd.date_range('1/1/2019', periods =8760, freq='H' )


#from oemof.solph import create_year_index

#my_index = create_year_index(2011)
My_system = EnergySystem(timeindex = my_index)
#My_system.set_snapshots(my_index)


#xls =pd.ExcelFile('C:\Users\Dev\Desktop\BBH mastearbeit_files\data_malla.xls')
data = pd.read_excel(r"C:\Users\Dev\Desktop\BBH mastearbeit_files\data_malla.xlsx",  sheet_name = 'Timeseries')

#create and components to Energysystem
b_electricity = Bus(label='electricity', balanced= False)
#b_gas = Bus(label='gas', balanced = False)
b_heat= Bus(label='heat', balanced= False)

My_system.add(b_electricity,b_heat)


#calculating annuities
epc_pv = economics.annuity(capex=1300, n=20, wacc=0.04)
epc_heat_pump = economics.annuity(capex= 1600, n = 20, wacc= 0.04)
epc_elec_boiler= economics.annuity(capex = 70, n= 20, wacc= 0.04)

#Sources
My_system.add(Source(label='source_electricity', outputs = {b_electricity: Flow(variable_costs= 0.46, emission_factor= 0.489)}))
My_system.add(Source(label = 're_roof_pv', outputs = {b_electricity: Flow(fix = data['sun_availability [-]'], nominal_value=56 )},\
                                 investment= solph.Investment(ep_costs=epc_pv)))



#My_system.add(Source(label ='source_gas', outputs={b_gas : Flow(variable_costs= 0.15, emission_factor=0.201)}))

#Sources
#My_system.add(Source(label='', outputs={b_gas : Flow()}))

#Demands
My_system.add(Sink(label='demand_elctricity', inputs = {b_electricity : Flow(fix = data['demand_electricity [kW]'])}))


#excess and shortage variable to avoid infeasible problems
My_system.add(Sink(label='excess_power', inputs={b_electricity:Flow(variable_costs = -0.115)}))

#calculating cops
data['COP_heat_pump']   = calc_cops(mode='heat_pump', temp_high= data['temperature of heating system [°C]'],
                                     temp_low=data['ambient_temperature [°C]'], quality_grade=0.45)

#Modeling Transformers
My_system.add(Transformer(label='t_heatpump', inputs = {b_electricity : Flow()}, outputs= {b_heat:Flow(nominal_value = 50000000000,\
           investment= solph.Investment(ep_costs=epc_heat_pump))}, conversion_factors= {b_electricity: data['COP_heat_pump']}))
    
My_system.add(Transformer(label='tr_elec_boiler', inputs= {b_electricity : Flow()}, outputs = {b_heat:Flow(max = 10000000000)},\
            investment = solph.Investment(ep_costs = epc_elec_boiler)))
    
    
My_system.add(Sink(label='demand_heat', inputs= {b_heat: Flow(fix = data['demand_heat [kW]'])}))
    



#solph.Investment()

#create optimization model based on energy system
optimization_model = Model(energysystem=My_system)

#solve Problem
optimization_model.solve(solver = 'cbc', solve_kwargs = {'tee' : True})

