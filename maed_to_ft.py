import os
import pandas as pd
import xlwings as xw
from utils import read_configuration, format_path
from shutil import copy
import openpyxl

def read_maed_results(maed_results, maed_years, first_col, last_col):
    wb1 = xw.Book(maed_results)
    sheet_1 = wb1.sheets['Sheet1']
    range_1 = sheet_1.range(( 1, first_col), (1, last_col))
    indelc=pd.DataFrame(range_1.value)
    indelc.index=maed_years
    indelc=indelc.T
    reselc=pd.DataFrame(range_1.value)
    reselc.index=maed_years
    reselc=reselc.T
    comelc=pd.DataFrame(range_1.value)
    comelc.index=maed_years
    comelc=comelc.T
    indelc[:]=sheet_1.range((264, first_col), (264, last_col)).value
    reselc[:]=sheet_1.range((751, first_col), (751, last_col)).value
    comelc[:]=sheet_1.range((677, first_col), (677, last_col)).value
    total_elec=indelc+comelc+reselc
    wb1.close()
    return total_elec
    
def calc_demand(maed_outputs, maed_years, start_year, end_year):
    parameter=pd.Series(list(maed_outputs.T.loc[:,0]))
    print(parameter)
    parameter.index=maed_years
    parameter_out=pd.Series(range(start_year, end_year+1))
    parameter_out.index=parameter_out.values
    parameter_out.loc[:]=None
    print(parameter_out)
    parameter_out.loc[maed_years]=parameter
    print(parameter_out)
    parameter_out=parameter_out.interpolate(method='linear')
    parameter_out=pd.DataFrame(parameter_out).T
    parameter_out=parameter_out.fillna(parameter[maed_years[0]])
    return parameter_out


def flextool_nodes(country, ft_template, total_elec_demand, reserve_margin):
    ft_nodes = pd.read_excel(ft_template, sheet_name="gridNode")
    ft_nodes.index = ft_nodes['node']
    ft_nodes.loc['nodeA','demand (MWh)']=total_elec_demand/ 0.0000036 
    ft_nodes.loc['nodeA','capacity margin (MW)']=total_elec_demand*(reserve_margin-1)/8760
    return ft_nodes

def extract_data(country, scenario, ft_template, ft_nodes):
    print("Preparing to extract data...")
    ft_template_copy_filename =  os.getcwd()+'./resources/flexTool-v2.0/InputData/' + country + "_" + scenario + "_FT.xlsx"
    ft_running_input_filename = './resources/flexTool-v2.0/InputData/' + "model_to_run.xlsx"
    copy(ft_template, ft_template_copy_filename)

    writer = pd.ExcelWriter(ft_template_copy_filename, engine='openpyxl', mode="a", if_sheet_exists="overlay")
    writer.workbook = openpyxl.load_workbook(ft_template_copy_filename)

    print("Extracting data...")
    ft_nodes.to_excel(writer, sheet_name='gridNode', header=False, index=False, startrow=1, startcol=0)
    writer.close()
    print("Extracting done!")
    copy(ft_template_copy_filename,ft_running_input_filename)
    return ft_template_copy_filename

def run(country, maed_scenario, maed_results,maed_years, ft_year):
    ft_template ="./resources/Mapping/Templates/template.xlsx"
    first_col=2
    last_col=first_col+len(maed_years)-1
    maed_outputs=read_maed_results(maed_results, maed_years, first_col, last_col)
    print(maed_outputs)
    start_year = 2015
    end_year=2070
    df= calc_demand(maed_outputs, maed_years, start_year, end_year)
    total_elec_demand=df[ft_year].values
    reserve_margin=1.15
    ft_nodes=flextool_nodes(country, ft_template, total_elec_demand, reserve_margin)
    extract_data(country, maed_scenario, ft_template, ft_nodes)

if __name__ == "__main__":
    country = input("Enter a country: ")
    maed_scenario = input("Enter a scenario (BAU, MS, NZ): ")
    ft_year = int(input("Enter a simulation year from 2015 to 2070 for FlexTool: "))
    maed_results=f"./resources/maed-2.0.0/maed_results.xlsx"
    config = read_configuration("maed")
    maed_years = config["years"]
    maed_years=[2022, 2025, 2030, 2035, 2040, 2045, 2050]
    run(country, maed_scenario, maed_results, maed_years, ft_year)