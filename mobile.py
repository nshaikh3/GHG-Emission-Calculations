import pandas as pd
import numpy as np

inventory_data = pd.ExcelFile(r'C:\Users\NoorShaikh\Documents\GHG App\S1&2 Data Collection Template.xlsx')
emission_factors = pd.ExcelFile(r'C:\Users\NoorShaikh\Documents\GHG App\Emission_Factors.xlsx')

mobile_combustion = inventory_data.parse('Scope 1 - Mobile')
for index, row in mobile_combustion.iterrows():
    if row['Scope 1'] == 'Facility unique ID':
        break
    mobile_combustion.drop([index],inplace=True)
mobile_combustion = mobile_combustion.rename(columns=mobile_combustion.iloc[0]).drop(mobile_combustion.index[0])
mobile_combustion = mobile_combustion.set_index(['Data Type', 'Fuel Type', 'Vehicle Type', 'Facility unique ID', 'Country', 'Year'])
mobile_combustion.fillna({'Fuel Consumption': 0, 'Distance Travelled': 0}, inplace=True)

s1_mc_emission_factors = emission_factors.parse('Mobile Combustion')
s1_mc_emission_factors = s1_mc_emission_factors.set_index(['Fuel Type', 'Vehicle Type'])

conversions = emission_factors.parse('Conversions')
conversions = conversions.set_index(['Convert From', 'Convert To'])

mobile_combustion_conversion = pd.merge(mobile_combustion, s1_mc_emission_factors[['Unit', 'MPG Units']], left_index=True, right_index=True)
mobile_combustion_conversion = mobile_combustion_conversion.set_index(['Fuel Unit', 'Unit'], append=True).rename_axis(index={'Fuel Unit': 'Convert From', 'Unit': 'Convert To'})
mobile_combustion_conversion = pd.merge(mobile_combustion_conversion, conversions['Multiply By'],left_index=True,right_index=True, how='left')

units = {'mile': 1, 'km': 1.60934, 'nautical mile': 0.868976}

mobile_combustion_conversion = mobile_combustion_conversion.reset_index(level=['Data Type'])
mobile_combustion_conversion['Fuel - Distance Activity'] = mobile_combustion_conversion['Distance Unit'].str.lower().map(units)
mobile_combustion_conversion['Multiply By'] = np.where(mobile_combustion_conversion['Data Type'] == 'Fuel Use', mobile_combustion_conversion['Multiply By'], mobile_combustion_conversion['Fuel - Distance Activity'])

mobile_combustion_conversion['Final Fuel Consumption'] = np.where(mobile_combustion_conversion['Data Type'] == 'Fuel Use', mobile_combustion_conversion['Fuel Consumption'] * mobile_combustion_conversion['Multiply By'], mobile_combustion_conversion['Distance Travelled'] * mobile_combustion_conversion['Multiply By'] / mobile_combustion_conversion['Fuel Efficiency'])

mobile_combustion_final = pd.merge(mobile_combustion_conversion, s1_mc_emission_factors,left_index=True,right_index=True, how='left').reset_index(level=['Convert From', 'Convert To']).set_index(['Data Type'], append = True)

mobile_combustion_co2_emissions = mobile_combustion_final['Final Fuel Consumption'] * mobile_combustion_final['CO2 Factor\n(kg / unit)']
mobile_combustion_ch4_emissions = mobile_combustion_final['Final Fuel Consumption'] * mobile_combustion_final['CH4 Factor\n(kg / unit)']
mobile_combustion_n2o_emissions = mobile_combustion_final['Final Fuel Consumption'] * mobile_combustion_final['N2O Factor\n(kg / unit)']
mobile_combustion_total_emissions = mobile_combustion_co2_emissions + (mobile_combustion_ch4_emissions*28) + (mobile_combustion_n2o_emissions*265)
mobile_combustion_summary_emissions = pd.concat([mobile_combustion_co2_emissions, mobile_combustion_ch4_emissions, mobile_combustion_n2o_emissions, mobile_combustion_total_emissions], axis=1)
mobile_combustion_summary_emissions.columns = ['kgCO2', 'kgCH4', 'kgN2O', 'kgCO2e']
print(mobile_combustion_summary_emissions)