import pandas as pd
import numpy as np

inventory_data = pd.ExcelFile(r'C:\Users\NoorShaikh\Documents\GHG App\S1&2 Data Collection Template.xlsx')
emission_factors = pd.ExcelFile(r'C:\Users\NoorShaikh\Documents\GHG App\Emission_Factors.xlsx')

purchased_energy = inventory_data.parse('Scope 2 - Purchased Energy')
for index, row in purchased_energy.iterrows():
    if row['Scope 2'] == 'Facility unique ID':
        break
    purchased_energy.drop([index],inplace=True)
purchased_energy = purchased_energy.rename(columns=purchased_energy.iloc[0]).drop(purchased_energy.index[0])
purchased_energy = purchased_energy.set_index(['Purchased energy type', 'Grid Region', 'Facility unique ID', 'Country', 'Year'])

s2_emission_factors = emission_factors.parse('Purchased Energy')
s2_emission_factors = s2_emission_factors.set_index(['Type', 'Grid Region'])

conversions = emission_factors.parse('Conversions')
conversions = conversions.set_index(['Convert From', 'Convert To'])

purchased_energy_conversion = pd.merge(purchased_energy, s2_emission_factors['Unit'], left_index=True, right_index=True)
purchased_energy_conversion = purchased_energy_conversion.set_index(['Unit_x', 'Unit_y'], append=True).rename_axis(index={'Unit_x': 'Convert From', 'Unit_y': 'Convert To'})
purchased_energy_conversion = pd.merge(purchased_energy_conversion, conversions['Multiply By'],left_index=True,right_index=True, how='left')
purchased_energy_conversion['Final Fuel Consumption'] = purchased_energy_conversion['Consumption'] * purchased_energy_conversion['Multiply By']

purchased_energy_final = pd.merge(purchased_energy_conversion, s2_emission_factors,left_index=True,right_index=True, how='left').reset_index(level=['Convert From', 'Convert To'])
purchased_energy_final.fillna({'CH4 Factor\n(kg/kWh)': 0, 'N2O Factor\n(kg/kWh)': 0}, inplace=True)

purchased_energy_co2_emissions = purchased_energy_final['Final Fuel Consumption'] * purchased_energy_final['CO2 Factor\n(kg/kWh)']
purchased_energy_ch4_emissions = purchased_energy_final['Final Fuel Consumption'] * purchased_energy_final['CH4 Factor\n(kg/kWh)']
purchased_energy_n2o_emissions = purchased_energy_final['Final Fuel Consumption'] * purchased_energy_final['N2O Factor\n(kg/kWh)']
purchased_energy_total_emissions = purchased_energy_co2_emissions + (purchased_energy_ch4_emissions*28/1000) + (purchased_energy_n2o_emissions*265/1000)
purchased_energy_summary_emissions = pd.concat([purchased_energy_co2_emissions, purchased_energy_ch4_emissions, purchased_energy_n2o_emissions, purchased_energy_total_emissions], axis=1)
purchased_energy_summary_emissions.columns = ['kgCO2', 'kgCH4', 'kgN2O', 'kgCO2e']
print(purchased_energy_summary_emissions)