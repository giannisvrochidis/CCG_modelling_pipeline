import pandas as pd
import csv
import os
from math import inf
from shutil import copy
import openpyxl
from time import strftime
from urllib import request, error
from fnmatch import fnmatch
import warnings
import numpy as np
import download_sand
from download_sand import download_sdk
import xlwings as xw
# %% [markdown]
# ## Define functions for downloading or locating locally OSeMOSYS SAND file

def get_sdk_data_as_df(country, scenario, sdk_path):
    if sdk_path: print("Reading SDK for country: " + country + ", scenario: " + scenario + "...")

    wb=xw.Book(sdk_path)
    sheet=wb.sheets['Parameters']
    try: 
        df = sheet.range('A1:BN48757')
        df=pd.DataFrame(df.value)
        df.columns = df.iloc[0]
        df = df[1:]
        wb.app.quit()

        # df= pd.read_excel(sdk_path, sheet_name="Parameters")
        # df.to_csv('df.csv')
        return df
    except: print("Failed to read SDK.")


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

def flextool_units(df, ft_template, year, output_csv):
    print(df[df['Parameter']=='VariableCost'])
    df_pwr = df[df.TECHNOLOGY.str.contains('PWR', na=False)]
    df_pwr = df_pwr[df_pwr["Parameter"].isin(['CapitalCost', 'EmissionActivityRatio', 'FixedCost', 'InputActivityRatio', 'OperationalLife',
                                              'OutputActivityRatio', 'ResidualCapacity', 'VariableCost'])]

    df_units = df_pwr.pivot_table(index='TECHNOLOGY', columns='Parameter', values=['Time indipendent variables', str(year)], aggfunc='sum')
    df_units.columns = df_units.columns.droplevel()
    df_units = df_units.groupby(axis=1, level=0).sum()
    df_units['ResidualCapacity'] = df_units['ResidualCapacity'] .apply(
        lambda x: x*1000)
    df_units['Efficiency'] = df_units['OutputActivityRatio'] / \
        (df_units['InputActivityRatio'])
    print(df_pwr)
    ft_units = pd.read_excel(ft_template, sheet_name="units")
    ft_units.index = ft_units['unit type']
    # ft_units['capacity (MW)'] = df_units['ResidualCapacity']
    # ft_units.replace([inf, -inf], 0, inplace=True)
    
    results=pd.read_csv(output_csv)
    results.index=results.Dim2
    capacities=results.loc[results.Variable=='TotalCapacityAnnual']
    capacities=capacities.loc[capacities.Dim3==str(year)]['ResultValue']
    capacities.dropna(inplace=True)

    ft_units.loc[:,'capacity (MW)'] = capacities.loc[:]*1000
    ft_units['capacity (MW)'].fillna(0, inplace=True)

    ft_unit_type = pd.read_excel(ft_template, sheet_name="unit_type")
    ft_unit_type.index = ft_unit_type['unit type']
    ft_unit_type['efficiency'] = df_units['Efficiency']
    ft_unit_type['O&M cost/MWh'] = df_units['VariableCost']
    ft_unit_type['fixed cost/kW/year'] = df_units['FixedCost']
    ft_unit_type['inv.cost/kW'] = df_units['CapitalCost']
    ft_unit_type['lifetime'] = df_units['OperationalLife']
    return ft_units, ft_unit_type

def flextool_fuels(df, ft_template,year):
    df_fuels = df[df.TECHNOLOGY.str.contains('MIN', na=False)]
    df_fuels = df_fuels[~df_fuels.TECHNOLOGY.str.contains('DEMIN', na=False)]
    df_fuels = df_fuels[df_fuels.Parameter.isin(
        ['EmissionActivityRatio', 'VariableCost'])]
    df_fuels = df_fuels.pivot_table(
        index='TECHNOLOGY', columns='Parameter', values=str(year), aggfunc='sum')
    df_fuels.index = df_fuels.index.str[3:6]
    df_fuels['VariableCost'] = df_fuels['VariableCost'].apply(lambda x: x*3.6)
    df_fuels['EmissionActivityRatio'] = df_fuels['EmissionActivityRatio'].apply(
        lambda x: x*3.6/1000)
    ft_fuel = pd.read_excel(ft_template, sheet_name="fuel")
    ft_fuel.index = ft_fuel['fuel']
    ft_fuel['fuel (price/MWh)'] = df_fuels['VariableCost']
    ft_fuel['CO2 content (t/MWh)'] = df_fuels['EmissionActivityRatio']
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

def flextool_nodes(country, df, output_csv, ft_template,year):
    results=pd.read_csv(output_csv)
    filtered_results = results[results['Dim2'].str.contains('ELC')]
    filtered_results = filtered_results[filtered_results['Variable'] == 'ProductionByTechnologyAnnual']

    df_sand=(df[df.TECHNOLOGY.isin(filtered_results[filtered_results.Dim4==str(2070)]['Dim2'].values)])
    eff=(df_sand[df_sand.Parameter.isin(['InputActivityRatio','OutputActivityRatio'])])
    eff=eff[['Parameter','TECHNOLOGY','FUEL',str(year)]]
    eff=eff.drop_duplicates()
    pivot_df = eff.pivot_table(index='TECHNOLOGY', columns='Parameter', values=str(year),  aggfunc='sum')
    pivot_df['Efficiency'] = pivot_df['OutputActivityRatio']/pivot_df['InputActivityRatio'] 
    new_df = pivot_df.reset_index()[['TECHNOLOGY', 'Efficiency']]
    new_df.columns = ['Technology', 'Efficiency']

    filtered_results=(filtered_results[filtered_results.Dim4==str(year)])
    new_df.index=new_df['Technology']
    filtered_results.index=filtered_results['Dim2']
    filtered_results['Efficiency']=new_df['Efficiency']
    filtered_results['Elec consumption']=filtered_results['ResultValue']/filtered_results['Efficiency']

    total_elec_demand=(filtered_results['Elec consumption'].sum())/ 0.0000036  #convert pj to mwh

    ft_nodes = pd.read_excel(ft_template, sheet_name="gridNode")
    ft_nodes.index = ft_nodes['node']
    ft_nodes.loc['nodeA','demand (MWh)']=total_elec_demand
    ft_nodes.loc['nodeA','capacity margin (MW)']=total_elec_demand*(df[df['Parameter']=='ReserveMargin'][str(year)].values[0]-1)/8760 #same reserve margin coefficient as in SAND
    return ft_nodes

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

def ft_calc(country, df, data_coll_file, ft_template, subfolder, year, output_csv):
    print("Start converting data to the appropriate format...")
    ft_units, ft_unit_type = flextool_units(df, ft_template, year, output_csv)
    ft_fuel = flextool_fuels(df, ft_template, year)
    ft_dem = flextool_dem(country, data_coll_file, ft_template, subfolder)
    ft_cf = flextool_cf(country, data_coll_file, ft_template, subfolder)
    ft_inflows=flextool_inflows(country, data_coll_file, ft_template, subfolder)
    ft_nodes=flextool_nodes(country, df, output_csv, ft_template,year)
    print("Convertion done!")
    return ft_units, ft_unit_type, ft_fuel, ft_dem, ft_cf, ft_inflows, ft_nodes

# %% [markdown]
# ## Write and save FlexTool template

def extract_data(country, scenario, ft_template, ft_data):
    print("Preparing to extract data...")
    ft_template_copy_filename =  os.getcwd()+'./resources/flexTool-v2.0/InputData/' + country + "_" + scenario + "_FT.xlsx"
    ft_running_input_filename = './resources/flexTool-v2.0/InputData/' + "model_to_run.xlsx"
    copy(ft_template, ft_template_copy_filename)

    writer = pd.ExcelWriter(ft_template_copy_filename, engine='openpyxl', mode="a", if_sheet_exists="overlay")
    writer.workbook = openpyxl.load_workbook(ft_template_copy_filename)

    ft_units, ft_unit_type, ft_fuel, ft_dem, ft_cf, ft_inflows, ft_nodes= ft_data

    print("Extracting data...")
    ft_fuel.to_excel(writer, sheet_name='fuel', header=False, index=False, startrow=1, startcol=0)
    ft_units.to_excel(writer, sheet_name='units', header=False, index=False, startrow=1, startcol=0)
    ft_unit_type.to_excel(writer, sheet_name='unit_type', header=False, index=False, startrow=1, startcol=0)
    ft_dem.to_excel(writer, sheet_name='ts_energy', header=False, index=False, startrow=3, startcol=2)
    ft_cf.to_excel(writer, sheet_name='ts_cf', header=False, index=False, startrow=2, startcol=2)
    ft_inflows.to_excel(writer, sheet_name='ts_inflow', header=False, index=False, startrow=2, startcol=2)
    ft_nodes.to_excel(writer, sheet_name='gridNode', header=False, index=False, startrow=1, startcol=0)
    writer.close()
    print("Extracting done!")
    copy(ft_template_copy_filename,ft_running_input_filename)
    return ft_template_copy_filename

def create_copy(country, scenario, ft_template_copy_filename, output_dir):
    copied_file_path = f"{output_dir}/{country}_{scenario}_FT.xlsx"
    try:
         if not os.path.exists(output_dir): os.makedirs(output_dir)
         copy(ft_template_copy_filename, copied_file_path)
         
    except:
        print('Failed to create FlexTool running folder')


# %% [markdown]
# # Choose country and scenario

def run(country, scenario, output_csv, data_source_path, output_dir):
    ft_template ="./resources/Mapping/Templates/template.xlsx"
    print("\n---------- OSeMOSYS to FlexTool v2.0 ----------\n")
    year = int(input("Enter a simulation year from 2015 to 2070 for FlexTool: "))
    if not country or not scenario: raise Exception("You need to provide a country, a scenario and a simulation year for FlexTool.")
    df = get_sdk_data_as_df(country, scenario, data_source_path)
    data_coll_file=get_data_coll(country, scenario)
    ts_folder = "./resources/Mapping/timeseries_inputs/"
    ft_data = ft_calc(country, df, data_coll_file, ft_template, ts_folder,year,output_csv)
    ft_template_copy_filename=extract_data(country, scenario,ft_template, ft_data)
    create_copy(country, scenario, ft_template_copy_filename, output_dir)


if __name__ == "__main__":
    country = input("Enter a country: ")
    scenario = input("Select SDK scenario (Base, NZv1, NZv2, LCv1, LCv2, FF): ")
    output_csv=f"./testing/SDK_dummy_results.csv"
    output_dir = f"./runs/FlexTool/{country}_{strftime('%Y-%m-%d_%H-%M-%S')}"
    model_dir_path, data_source_path = download_sand.run(country, scenario, output_dir)
    run(country, scenario, output_csv, data_source_path, output_dir)
