import pandas as pd
import numpy as np

inventory_data = pd.ExcelFile(r'C:\Users\NoorShaikh\Documents\GHG App\S1&2 Data Collection Template.xlsx')
emission_factors = pd.ExcelFile(r'C:\Users\NoorShaikh\Documents\GHG App\Emission_Factors.xlsx')

fugitive = inventory_data.parse('Scope 1 - Fugitives')
for index, row in fugitive.iterrows():
    if row['Scope 1'] == 'Facility unique ID':
        break
    fugitive.drop([index],inplace=True)
fugitive = fugitive.rename(columns=fugitive.iloc[0]).drop(fugitive.index[0])
fugitive = fugitive.set_index(['Refrigerant Type', 'Facility unique ID', 'Country', 'Year'])

s1_f_emission_factors = emission_factors.parse('Fugitives')
s1_f_emission_factors = s1_f_emission_factors.set_index('Refrigerant Type')

units = {'lb': 1/2.2, 'metric ton': 1000, 'long ton': 1016, 'short ton': 907, 'kg': 1, 'kilograms': 1}

fugitive['Quantity Serviced:'] *= fugitive['Unit of Quantity Serviced (dropdown list):'].str.lower().map(units)
fugitive['Unit of Quantity Serviced (dropdown list):'] = 'Kilograms'

fugitive['Quantity Recycled:'] *= fugitive['Unit of Quantity Recycled (dropdown list):'].str.lower().map(units)
fugitive['Unit of Quantity Recycled (dropdown list):'] = 'Kilograms'

fugitive['Quantity Recycled:'].fillna(0, inplace=True)

fugitive['Final Quantity'] = fugitive['Quantity Serviced:'] - fugitive['Quantity Recycled:']

fugitive_summary_emissions = fugitive['Final Quantity'].mul(s1_f_emission_factors['AR5 (kgCO2e)'])
fugitive_summary_emissions.columns = ['kgCO2e']
print(fugitive_summary_emissions)