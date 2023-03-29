# %% [markdown]
# ## Import necessary packages

import pandas as pd
import csv
import os
from math import inf
from shutil import copy
import openpyxl
from urllib import request, error
from fnmatch import fnmatch
import warnings
import numpy as np
from download_sand import download_sdk
from time import strftime

# %% [markdown]
# ## Define functions for downloading or locating locally OSeMOSYS SAND file
def get_data_coll(country,scenario):
    data_filename_pattern = country+' Data Collection.xlsx'
    data_folder = "./inputs/OSeMOSYS Starter Kits/Data Collection Files/"
    data_file = None
    data_coll_file=None
    for file in os.listdir(os.path.join(data_folder)):
        if fnmatch(file, data_filename_pattern):
            data_file = file
            break
    if data_file is not None: data_coll_file=pd.ExcelFile(data_folder+data_file)
    if data_coll_file is not None: return data_coll_file
    data_urls_location = f"./resources/Mapping/download_links/Data_collection_links.csv"
    data_location = data_folder+data_filename_pattern
    download_sdk(country, data_urls_location, data_location)
    for file in os.listdir(os.path.join(data_folder)):
        if fnmatch(file, data_filename_pattern):
            data_file = file
            break
    try: 
        data_coll_file=pd.ExcelFile(data_folder+data_file)
        return data_coll_file
    except:
        print('Could not collect data collection file')
# %% [markdown]
# ## Define functions for preparing FlexTool 2.0 input data

def flextool_units(data_coll_file, ft_template, year):
    df_sets=pd.read_excel(data_coll_file,sheet_name="SETS",skiprows=1,usecols=[1,2],index_col=1)
    df_sets=df_sets.drop(['No longer used','Additional Technology'])

    df_costs=pd.read_excel(data_coll_file,sheet_name="3.1 Technology Costs",skiprows=1)
    df_costs=df_costs[['Technology Code','Cost Parameter',year]]
    df_costs=df_costs.pivot_table(index='Technology Code', columns='Cost Parameter',values=year)

    df_input_act=pd.read_excel(data_coll_file,sheet_name="3.2 Input Activity Ratios",skiprows=1)
    df_input_act=df_input_act.pivot_table(index='Technology Code',values=year)
    df_eff=1/df_input_act

    df_life=pd.read_excel(data_coll_file,sheet_name="3.5 Operational Life",skiprows=1)
    df_life=df_life.pivot_table(index='Technology Code', values='Operational Life')

    df_caps=pd.read_excel(data_coll_file,sheet_name="Data in Brief Table 1",skiprows=1,usecols=[0,1,2,3,4],index_col=0)
    df_caps.index.name=None
    print(df_caps)
    df_costs.index.name=None
    df_caps=df_caps.drop('No longer used')
    df_caps.loc[:,'Technology Code']=df_sets['Code']
    df_caps=df_caps.pivot_table(index='Technology Code',values=year)

    ft_units = pd.read_excel(ft_template, sheet_name="units")
    ft_units.index = ft_units['unit type']
    ft_units['capacity (MW)'] = df_caps
    ft_units['capacity (MW)'].fillna(0, inplace=True)

    ft_unit_type = pd.read_excel(ft_template, sheet_name="unit_type")
    ft_unit_type.index = ft_unit_type['unit type']
    ft_unit_type['efficiency'] = df_eff
    ft_unit_type['O&M cost/MWh'] = df_costs['Variable Cost ($/GJ)']
    ft_unit_type['fixed cost/kW/year'] = df_costs['Fixed Cost ($/kW/yr)']
    ft_unit_type['inv.cost/kW'] = df_costs['Capital Cost ($/kW)']
    ft_unit_type['lifetime'] = df_life
    ft_unit_type[['efficiency','O&M cost/MWh','fixed cost/kW/year','inv.cost/kW','lifetime']].fillna(0, inplace=True)
    return ft_units, ft_unit_type

def flextool_fuels(data_coll_file, ft_template, year):
    fuels={'Crude oil':'OIL','Biomass':'BIO','Coal':'COA','Natural Gas':'NGS'}
    df_emissions = pd.read_excel(data_coll_file,sheet_name="Data in Brief Table 7",usecols=[0,1],index_col=0)
    df_emissions.index = df_emissions.index.map(fuels)
    df_emissions=df_emissions.loc[df_emissions.index.dropna()]

    df_costs=pd.read_excel(data_coll_file,sheet_name="3.1 Technology Costs",skiprows=1)
    df_costs=df_costs[['Technology Code','Cost Parameter',year]]
    df_costs=df_costs.pivot_table(index='Technology Code', columns='Cost Parameter',values=year)
    df_costs=df_costs[df_costs.index.str.contains('MIN|IMP', na = False)]
    df_costs=df_costs[~df_costs.index.str.contains('DEMIN', na = False)]
    df_costs.index = df_costs.index.str[3:6]
    df_costs=df_costs.groupby(df_costs.index).mean()

    ft_fuel = pd.read_excel(ft_template, sheet_name="fuel")
    ft_fuel.index = ft_fuel['fuel']
    ft_fuel['fuel (price/MWh)'] = df_costs['Variable Cost ($/GJ)'].apply(lambda x: x*3.6)
    ft_fuel['CO2 content (t/MWh)'] = df_emissions.apply(
            lambda x: x*3.6/1000)
    ft_fuel.fillna(0,inplace=True)
    return ft_fuel


def flextool_dem(country, data_coll_file, ft_template, subfolder):
    # dem_all = pd.read_csv(subfolder +"All_Demand_UTC_2015.csv", skiprows=1)
    # dem_all.index = dem_all.iloc[:, 0]
    # dem_all.drop(columns=dem_all.columns[0], inplace=True)
    # dem = dem_all.loc[:, country]
    dem=pd.read_excel(data_coll_file,sheet_name='4.2 Elc Demand Profile Raw Data',skiprows=2)
    dem=dem.iloc[:,1]
    ft_dem = pd.read_excel(ft_template, sheet_name="ts_energy")
    ft_dem.columns = ft_dem.loc[0]
    ft_dem.drop(index=ft_dem.index[[0, 1]], inplace=True)
    ft_dem.index = ft_dem.iloc[:, 0]
    ft_dem.drop(columns=ft_dem.columns[[0, 1]], inplace=True)
    ft_dem['nodeA'] = dem.values
    return ft_dem


def flextool_cf(country, data_coll_file, ft_template, subfolder):
    csp_all = pd.read_csv(subfolder +"CSP 2010-2017.csv", skiprows=1)
    csp_all.index = csp_all.iloc[:, 0]
    csp_all.drop(columns=csp_all.columns[0], inplace=True)
    csp=csp_all.loc[:, country]/100

    wof_all = pd.read_csv(subfolder + "Woff 2010-2017.csv", skiprows=1)
    wof_all.index = wof_all.iloc[:, 0]
    wof_all.drop(columns=wof_all.columns[0], inplace=True)
    try:
        wof=wof_all.loc[:, country]/100
    except:
        wof=None
    pv=pd.read_excel(data_coll_file,sheet_name='3.6 Raw PV CapFac (auto)',skiprows=4)
    pv=pv.iloc[:,2]/100
    won=pd.read_excel(data_coll_file,sheet_name='3.6 Raw Onshore Wind CFs (auto)',skiprows=4)
    won=won.iloc[:,2]/100
    ft_cf = pd.read_excel(ft_template, sheet_name="ts_cf")
    ft_cf.index = ft_cf.iloc[:, 0]
    ft_cf.drop(index=ft_cf.index[0], inplace=True)
    ft_cf.index = ft_cf.iloc[:, 0]
    ft_cf.drop(columns=ft_cf.columns[[0, 1]], inplace=True)
    ft_cf['SOL001'] = pv.values
    ft_cf['SOL001S'] = pv.values
    ft_cf['CSP002'] = csp.values
    ft_cf['WND001'] = won.values
    if wof is not None: ft_cf['WND002'] = wof.values
    ft_cf['WND001S'] = won.values

    return ft_cf


def flextool_inflows(country, data_coll_file, ft_template, subfolder):
    hydro=pd.read_excel(data_coll_file,sheet_name='3.6 Raw Hydro CFs (auto)',skiprows=4)
    
    inflow=pd.DataFrame(np.nan,index=range(0,8760),columns=["Hydro"],dtype=float)
    inflow[0:744]=hydro.iloc[0,3]/100
    inflow[744:1416]=hydro.iloc[0,4]/100
    inflow[1416:2160]=hydro.iloc[0,5]/100
    inflow[2160:2880]=hydro.iloc[0,6]/100
    inflow[2880:3624]=hydro.iloc[0,7]/100
    inflow[3624:4344]=hydro.iloc[0,8]/100
    inflow[4344:5088]=hydro.iloc[0,9]/100
    inflow[5088:5832]=hydro.iloc[0,10]/100
    inflow[5832:6552]=hydro.iloc[0,11]/100
    inflow[6552:7296]=hydro.iloc[0,12]/100
    inflow[7296:8016]=hydro.iloc[0,13]/100
    inflow[8016:8760]=hydro.iloc[0,14]/100
    
    ft_inflows= pd.read_excel(ft_template, sheet_name="ts_inflow")
    ft_inflows.index = ft_inflows.iloc[:, 0]
    ft_inflows.drop(index=ft_inflows.index[0], inplace=True)
    ft_inflows.index = ft_inflows.iloc[:, 0]
    ft_inflows.drop(columns=ft_inflows.columns[[0, 1]], inplace=True)
    
    ft_inflows['HYD001']=inflow.values
    ft_inflows['HYD002']=inflow.values
    ft_inflows['HYD003']=inflow.values
    return ft_inflows

def ft_calc(country, data_coll_file, ft_template, subfolder, year):
    print("Start converting data to the appropriate format...")
    ft_units, ft_unit_type = flextool_units(data_coll_file, ft_template, year)
    ft_fuel = flextool_fuels(data_coll_file, ft_template,year)
    ft_dem = flextool_dem(country, data_coll_file, ft_template, subfolder)
    ft_cf = flextool_cf(country, data_coll_file, ft_template, subfolder)
    ft_inflows=flextool_inflows(country, data_coll_file, ft_template, subfolder)
    print("Convertion done!")
    return ft_units, ft_unit_type, ft_fuel, ft_dem, ft_cf, ft_inflows

# %% [markdown]
# ## Write and save FlexTool template

def extract_data(country, scenario, ft_template, ft_data):
    print("Preparing to extract data...")
    ft_template_copy_filename =  './resources/flexTool-v2.0/InputData/models/' + country + "_" + scenario + "_FT.xlsx"
    ft_running_input_filename = './resources/flexTool-v2.0/InputData/' + "model_to_run.xlsx"
    copy(ft_template, ft_template_copy_filename)

    writer = pd.ExcelWriter(ft_template_copy_filename, engine='openpyxl', mode="a", if_sheet_exists="overlay")
    writer.workbook = openpyxl.load_workbook(ft_template_copy_filename)

    ft_units, ft_unit_type, ft_fuel, ft_dem, ft_cf, ft_inflows = ft_data

    print("Extracting data...")
    ft_fuel.to_excel(writer, sheet_name='fuel', header=False, index=False, startrow=1, startcol=0)
    ft_units.to_excel(writer, sheet_name='units', header=False, index=False, startrow=1, startcol=0)
    ft_unit_type.to_excel(writer, sheet_name='unit_type', header=False, index=False, startrow=1, startcol=0)
    ft_dem.to_excel(writer, sheet_name='ts_energy', header=False, index=False, startrow=3, startcol=2)
    ft_cf.to_excel(writer, sheet_name='ts_cf', header=False, index=False, startrow=2, startcol=2)
    ft_inflows.to_excel(writer, sheet_name='ts_inflow', header=False, index=False, startrow=2, startcol=2)
    writer.close()
    print("Extracting done!")
    copy(ft_template_copy_filename,ft_running_input_filename )
    return ft_template_copy_filename

def create_copy(ft_template_copy_filename, new_dir):
    copied_file_path = new_dir+country + "_" + scenario + "_FT.xlsx"
    os.makedirs(new_dir, exist_ok=True)
    copy(ft_template_copy_filename, copied_file_path)
    
# %% [markdown]
# # Choose country and scenario

def run(country, scenario):
    ft_template ="./resources/Mapping/Templates/template.xlsx"
    print("\n---------- FlexTool Starter Data Kits ----------\n")
    year = int(input("Enter a simulation year from 2015 to 2018 for FlexTool: "))
    if not country or not scenario: raise Exception("You need to provide a country, a scenario and a simulation year for FlexTool.")
    data_coll_file=get_data_coll(country, scenario)
    ts_folder = "./resources/Mapping/timeseries_inputs/"
    ft_data = ft_calc(country, data_coll_file, ft_template, ts_folder,year)
    ft_template_copy_filename=extract_data(country, scenario,ft_template, ft_data)
    
    new_dir = f"./runs/FlexTool/{country}_{year}_run_{strftime('%Y-%m-%d_%H-%M-%S')}/"
    create_copy(ft_template_copy_filename, new_dir)
    return new_dir

if __name__ == "__main__":
    country = input("Enter a country: ")
    scenario='Base'
    run(country, scenario)
