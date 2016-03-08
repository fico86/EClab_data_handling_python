# -*- coding: utf-8 -*-
"""
Created on Wed Feb 24 09:55:18 2016

@author: Binayak

Sorting galvanostatic data from EC-lab by cycles and normalised by mass, to plot in Origin.

Input format: 
Excel file containing 3 sheets named "gal_files", "mass_grid", "selected_cycles"
eg. TiN + Carbon ink mixture vs carbon control ink mixture, both cycled in Ar and then O2

"gal_files" sheet:
    columns: list of treatment on samples, eg. cycled in Ar, then O2
    rows: list of diffrent samples, eg. TiN + carbon, carbon control
    values:cv file names of with extention of the cooresponding CVs
-------------------------------------------------------------------------------
        | Ar                               | O2                               |
-------------------------------------------------------------------------------
TiN + C | (Ar cv file name with extention) | (Ar cv file name with extention) |
-------------------------------------------------------------------------------
C       | (Ar cv file name with extention) | (Ar cv file name with extention) |
-------------------------------------------------------------------------------

"mass_grid" sheet:
    columns: list diffrent masses to be normalised by 
    rows: list of diffrent samples, must match the rows in "gal_files"
    values: corresponding masses in grams, masses set to 0 will not be sorted
---------------------------------------------------------------------
        | active mass            | Tin           | carbon           |
---------------------------------------------------------------------
TiN + C | (mass of TiN + Carbon) | (mass of TiN) | (mass of carbon) |
---------------------------------------------------------------------
C       | (mass of Carbon)       | 0             | (mass of Carbon) |
---------------------------------------------------------------------

"selected_cycles'
    columns: list of treatment on samples, must match columns of "gal_files"
    rows: general list of numbers starting from 0, does not effect the sorting
    values: cycle numbers to be sorted, for each treatment, set to 0 finishes cycle selection
--------------
  | Ar | O2  |
--------------
0 | 1  | 1   |
--------------
1 | 3  | 5   |
--------------
2 | 0  | 10  |
--------------
3 | 0  | 15  |
--------------
4 | 0  | 20  |
--------------

In output file, will be sorted into sheets based on the normalising masses. 
The first file is 'as measured', ie. not normalised, followed by the masses listed in the "mass_grid". 
When mass is set to 0, that particuler normalisation is ignored, ie the TiN normalised sheet, will only have the TiN + C gal values.
"""

import pandas as pd
import numpy as np

#getting the start line number in a cv file
def getline(thefilepath, desired_line_number): 
    if desired_line_number < 1: return ''
    for current_line_number, line in enumerate(open(thefilepath, 'rU')):
        if current_line_number == desired_line_number-1: return line
    return ''

#converting data in cv file into Pandas dataframe format, grouped by cycle number    
def text_to_grouped_dataframe(filename): 
    line = getline(filename, 2)
    datarow = int(line[18:21])
    return pd.DataFrame(pd.read_csv(filename, sep = None, skiprows = datarow-1)).groupby('half cycle')
    
    #attaching columns labels, cycle numbers and normalising by mass to particuler cycle from particuler cv dataset.    
def cycle_dataframe_labeled(grouped, cycle, label, printlabel, mass): 
    if printlabel is False:
        label = ''
        
    if mass == 1:
        header = pd.DataFrame([['Capacity', 'Potential vs Li/Li+', 
                                'Capacity', 'Potential vs Li/Li+'],
                               ['mAh', 'V', 'mAh', 'V'],
                               ['cycle {} {} charge'.format(str(cycle), label), 'cycle {} {} charge'.format(str(cycle), label),
                                'cycle {} {} discharge'.format(str(cycle), label), 'cycle {} {} discharge'.format(str(cycle), label)]], 
                                columns=['{}a {} charge'.format(str(cycle), label), '{}b {} charge'.format(str(cycle), label),
                                         '{}a {} discharge'.format(str(cycle), label), '{}b {} discharge'.format(str(cycle), label)])
    else:
        header = pd.DataFrame([['Specific Capacity', 'Potential vs Li/Li+', 
                                'Specific Capacity', 'Potential vs Li/Li+'],
                               ['mAh/g', 'V', 'mAh/g', 'V'],
                               ['cycle {} {} charge'.format(str(cycle), label), 'cycle {} {} charge'.format(str(cycle), label),
                                'cycle {} {} discharge'.format(str(cycle), label), 'cycle {} {} discharge'.format(str(cycle), label)]], 
                                columns=['{}a {} charge'.format(str(cycle), label), '{}b {} charge'.format(str(cycle), label),
                                         '{}a {} discharge'.format(str(cycle), label), '{}b {} discharge'.format(str(cycle), label)])
                                
    value = pd.DataFrame([grouped.get_group(2*cycle-2)['Capacity/mA.h'].values/mass, 
                          grouped.get_group(2*cycle-2)['Ewe/V'].values, 
                          grouped.get_group(2*cycle-1)['Capacity/mA.h'].values/mass, 
                          grouped.get_group(2*cycle-1)['Ewe/V'].values], 
                          index = ['{}a {} charge'.format(str(cycle), label), 
                                   '{}b {} charge'.format(str(cycle), label),
                                   '{}a {} discharge'.format(str(cycle), label), 
                                   '{}b {} discharge'.format(str(cycle), label)]).transpose()
    dataframe = header.append(value, ignore_index=True,)
    return dataframe
    
def gal_capacity(grouped, mass, label):
    if mass == 1:
        header = pd.DataFrame([['cycle', 'Capacity', 'Capacity'],
                               ['', 'mAh', 'mAh'],
                               ['{}'.format(label), '{} charge'.format(label), '{} discharge'.format(label)]])
        
    else:
        header = pd.DataFrame([['cycle', 'Spacific Capacity', 'Spacific Capacity'],
                               ['', 'mAh/g', 'mAh/g'],
                               ['{}'.format(label), '{} charge'.format(label), '{} discharge'.format(label)]])
    
    
    for cycle in np.arange(1,len(grouped)/2):
        charge = grouped.get_group(2*cycle-2)['Capacity/mA.h'].values[-1]/mass
        discharge = grouped.get_group(2*cycle-1)['Capacity/mA.h'].values[-1]/mass
        values = pd.DataFrame([[cycle, charge, discharge]])
        header = header.append(values, ignore_index=True)
    
    return header
                               

    
def gal_compare_by_mass(input_file, output_file):
    
    inputs = pd.read_excel(input_file, sheetname = None)
    gal_files = inputs['gal_files']
    mass_grid = inputs['mass_grid']
    selected_cycles = inputs['selected_cycles']
    
    gal_data = gal_files.applymap(text_to_grouped_dataframe)
    
    data_out = {}
    data_out['as mesured'] = {}
    
    for state in selected_cycles.columns.values:
        no_mass_out = pd.DataFrame()
        for cycle in selected_cycles[state]:
            if cycle == 0:
                break
            for con in gal_data.index.values:
                data = cycle_dataframe_labeled(gal_data[state][con], cycle, str(con) + ' ' + state, True, 1)
                no_mass_out = pd.concat([no_mass_out, data], axis = 1)
        
        data_out['as mesured'][state] = no_mass_out
        

    
    for mass in mass_grid.columns.values:
        data_out[mass] = {}
        for state in selected_cycles.columns.values:
            mass_out = pd.DataFrame()
            for cycle in selected_cycles[state]:
                if cycle == 0:
                    break
                for con in gal_data.index.values:
                    if mass_grid[mass][con] == 0:
                        break
                    data = cycle_dataframe_labeled(gal_data[state][con], cycle, str(con) + ' ' + state, True, mass_grid[mass][con])
                    mass_out = pd.concat([mass_out, data], axis = 1)
            data_out[mass][state] = mass_out
            
            
    capacity_out = {}
    capacity_out['as mesured'] = {}
    
    for state in selected_cycles.columns.values:
        cap_out = pd.DataFrame()
        for con in gal_data.index.values:
            data = gal_capacity(gal_data[state][con], 1, con + ' ' + state)
            cap_out = pd.concat([cap_out,data], axis = 1)
            
        capacity_out['as mesured'][state] = cap_out
    
    for mass in mass_grid.columns.values:
        capacity_out[mass] = {}
        for state in selected_cycles.columns.values:
            cap_out = pd.DataFrame()
            for con in gal_data.index.values:
                if mass_grid[mass][con] == 0:
                        break
                data = gal_capacity(gal_data[state][con], mass_grid[mass][con], con +' ' + state)
                cap_out = pd.concat([cap_out,data], axis = 1)
        capacity_out[mass][state] = cap_out

    for mass in data_out.keys():
        with pd.ExcelWriter('{} - {}.xlsx'.format(output_file, mass)) as writer:
            for state in data_out[mass].keys():
                data_out[mass][state].to_excel(writer, sheet_name = state + ' ' + 'profile', header = False, index = False)
                capacity_out[mass][state].to_excel(writer, sheet_name = state + ' ' + 'capacity', header = False, index = False)


"""           
    
"""

    