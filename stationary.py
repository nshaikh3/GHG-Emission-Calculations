import pandas as pd
import numpy as np

inventory_data = pd.ExcelFile(r'C:\Users\NoorShaikh\Documents\GHG App\S1&2 Data Collection Template.xlsx')
emission_factors = pd.ExcelFile(r'C:\Users\NoorShaikh\Documents\GHG App\Emission_Factors.xlsx')

stationary_combustion = inventory_data.parse('Scope 1 - Stationary')
for index, row in stationary_combustion.iterrows():
    if row['Scope 1'] == 'Facility unique ID':
        break
    stationary_combustion.drop([index],inplace=True)
stationary_combustion = stationary_combustion.rename(columns=stationary_combustion.iloc[0]).drop(stationary_combustion.index[0])
stationary_combustion = stationary_combustion.set_index(['Fuel Type', 'Facility unique ID', 'Country', 'Year'])

s1_emission_factors = emission_factors.parse('Stationary Combustion')
s1_emission_factors = s1_emission_factors.set_index('Fuel Type')

conversions = emission_factors.parse('Conversions')
conversions = conversions.set_index(['Convert From', 'Convert To'])

stationary_combustion_conversion = pd.merge(stationary_combustion, s1_emission_factors['Unit'], left_index=True, right_index=True)
stationary_combustion_conversion = stationary_combustion_conversion.set_index(['Unit_x', 'Unit_y'], append=True).rename_axis(index={'Unit_x': 'Convert From', 'Unit_y': 'Convert To'})
stationary_combustion_conversion = pd.merge(stationary_combustion_conversion, conversions['Multiply By'],left_index=True,right_index=True, how='left')
stationary_combustion_conversion['Multiply By'].fillna(0, inplace=True)
stationary_combustion_conversion['Final Fuel Consumption'] = stationary_combustion_conversion['Fuel Consumption'] * stationary_combustion_conversion['Multiply By']

stationary_combustion_final = pd.merge(stationary_combustion_conversion, s1_emission_factors, left_index=True, right_index=True, how='left').reset_index(level=['Convert From', 'Convert To'])

stationary_combustion_co2_emissions = stationary_combustion_final['Final Fuel Consumption'] * stationary_combustion_final['CO2 Factor (kg CO2 per mmBtu)']
stationary_combustion_ch4_emissions = stationary_combustion_final['Final Fuel Consumption'] * stationary_combustion_final['CH4 Factor (g CH4 per mmBtu)']
stationary_combustion_n2o_emissions = stationary_combustion_final['Final Fuel Consumption'] * stationary_combustion_final['N2O Factor (g N2O per mmBtu)']
stationary_combustion_total_emissions = stationary_combustion_co2_emissions + (stationary_combustion_ch4_emissions*28/1000) + (stationary_combustion_n2o_emissions*265/1000)
stationary_combustion_summary_emissions = pd.concat([stationary_combustion_co2_emissions, stationary_combustion_ch4_emissions, stationary_combustion_n2o_emissions, stationary_combustion_total_emissions], axis=1)
stationary_combustion_summary_emissions.columns = ['kgCO2', 'kgCH4', 'kgN2O', 'kgCO2e']
print(stationary_combustion_summary_emissions)
