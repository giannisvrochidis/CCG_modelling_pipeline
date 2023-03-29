import pandas as pd
import os
import subprocess
import numpy as np
import json
import shutil
from utils import read_configuration, format_path
import openpyxl

#
def read_dataframe(years, data_out, sheet_name):
    df=  pd.read_excel(data_out,sheet_name=sheet_name)
    df.index=df.iloc[:,0]
    return df

def extract_data(country, scenario, data_out, economic, freight, intercity, urban, transport_international):
    print("Preparing to extract data...")

    writer = pd.ExcelWriter(data_out, engine='openpyxl', mode="a", if_sheet_exists="overlay")
    writer.workbook = openpyxl.load_workbook(data_out)

    economic_demography, economic_gdp = economic
    transport_freight_generation, transport_freight_intensity, transport_freight_modal = freight
    transport_intercity_factors, transport_intercity_intensity, transport_intercity_modal = intercity
    transport_urban_factors, transport_urban_intensity, transport_urban_modal = urban

    print("Extracting data...")
    economic_demography.to_excel(writer, sheet_name='economic_demography', header=True, index=False)
    economic_gdp.to_excel(writer, sheet_name='economic_gdp', header=True, index=False)
    transport_freight_generation.to_excel(writer, sheet_name='transport_freight_generation', header=True, index=False)
    transport_freight_intensity.to_excel(writer, sheet_name='transport_freight_intensity', header=True, index=False)
    transport_freight_modal.to_excel(writer, sheet_name='transport_freight_modal', header=True, index=False)
    transport_intercity_factors.to_excel(writer, sheet_name='transport_intercity_factors', header=True, index=False)
    transport_intercity_intensity.to_excel(writer, sheet_name='transport_intercity_intensity', header=True, index=False)
    transport_intercity_modal.to_excel(writer, sheet_name='transport_intercity_modal', header=True, index=False)
    transport_urban_factors.to_excel(writer, sheet_name='transport_urban_factors', header=True, index=False)
    transport_urban_intensity.to_excel(writer, sheet_name='transport_urban_intensity', header=True, index=False)
    transport_urban_modal.to_excel(writer, sheet_name='transport_urban_modal', header=True, index=False)
    transport_international.to_excel(writer, sheet_name='transport_international', header=True, index=False)

    writer.close()
    print("Extracting done!")

def run(country, scenario, years):
    config = read_configuration("maed")
    historic_years=[y for y in years if y <=2022]
    projection_years=[y for y in years if y >=2022]

    data_template = "./resources/maed-2.0.0/inputs/maedd_template_TSDK.xlsx"
    data_out ="./resources/maed-2.0.0/inputs/models/"+country+'_'+scenario+'.xlsx'
    shutil.copy(data_template,data_out)
    tsdk_file = "./inputs/MAED Transport Starter Kits/TSDK_"+country+'.xlsx'
    historic_data = pd.read_excel(tsdk_file, sheet_name='Historical')
    projection_data= pd.read_excel(tsdk_file, sheet_name='Projection - '+scenario)

    #Socioeconomic tabs filling
    economic_demography = read_dataframe(years, data_out, 'economic_demography')
    economic_gdp = read_dataframe(years, data_out, 'economic_gdp')
    transport_freight_generation = read_dataframe(years, data_out, 'transport_freight_generation')
    transport_freight_modal = read_dataframe(years, data_out, 'transport_freight_modal')
    transport_freight_intensity = read_dataframe(years,data_out,'transport_freight_intensity')
    transport_intercity_factors = read_dataframe(years,data_out,'transport_intercity_factors')
    transport_intercity_modal = read_dataframe(years,data_out,'transport_intercity_modal')
    transport_intercity_intensity = read_dataframe(years,data_out,'transport_intercity_intensity')
    transport_urban_factors = read_dataframe(years,data_out,'transport_urban_factors')
    transport_urban_modal = read_dataframe(years,data_out,'transport_urban_modal')
    transport_urban_intensity = read_dataframe(years,data_out,'transport_urban_intensity')
    transport_international = read_dataframe(years,data_out,'transport_international')

    economic_demography.loc['Population',historic_years]=historic_data.loc[(historic_data.Variable=='Population') & (historic_data.Type=='All'),historic_years].values[0]
    economic_demography.loc['Population growth rate',historic_years]=historic_data.loc[(historic_data.Variable=='Population growth') & (historic_data.Type=='All'),historic_years].values[0]
    economic_demography.loc['Urban Population',historic_years]=historic_data.loc[(historic_data.Variable=='Population') & (historic_data.Type=='Urban'),historic_years].values[0]
    economic_demography.loc['Rural Population',historic_years]=historic_data.loc[(historic_data.Variable=='Population') & (historic_data.Type=='Rural'),historic_years].values[0]

    economic_demography.loc['Population growth rate',projection_years]=projection_data.loc[(projection_data.Variable=='Population growth') & (projection_data.Type=='All'),projection_years].values[0]
    economic_demography.loc['Rural Population',projection_years]=projection_data.loc[(projection_data.Variable=='Population') & (projection_data.Type=='Rural'),projection_years].values[0]
    economic_demography.loc['Urban Population',projection_years]=projection_data.loc[(projection_data.Variable=='Population') & (projection_data.Type=='Urban'),projection_years].values[0]

    economic_demography.loc['Population in cities with public transport',historic_years+projection_years]=25
    economic_demography.loc['Potential Labour Force',historic_years+projection_years]=50
    economic_demography.loc['Participating Labour Force',historic_years+projection_years]=50
    economic_demography.loc['Person/ rural Household',historic_years+projection_years]=3
    economic_demography.loc['Person/ urban Household',historic_years+projection_years]=3

    economic_gdp.loc['GDP', historic_years]=historic_data.loc[(historic_data.Variable=='GDP') & (historic_data.Type=='All'),historic_years].values[0]
    economic_gdp.loc['GDP Growth rate',historic_years]=historic_data.loc[(historic_data.Variable=='GDP growth') & (historic_data.Type=='All'),historic_years].values[0]
    economic_gdp.loc['Agriculture',historic_years]=historic_data.loc[(historic_data.Variable=='GDP') & (historic_data.Type=='Agriculture'),historic_years].values[0]
    economic_gdp.loc['Construction',historic_years]=historic_data.loc[(historic_data.Variable=='GDP') & (historic_data.Type=='Construction'),historic_years].values[0]
    economic_gdp.loc['Mining',historic_years]=historic_data.loc[(historic_data.Variable=='GDP') & (historic_data.Type=='Mining'),historic_years].values[0]
    economic_gdp.loc['Manufacturing',historic_years]=historic_data.loc[(historic_data.Variable=='GDP') & (historic_data.Type=='Manufacturing'),historic_years].values[0]
    economic_gdp.loc['Service',historic_years]=historic_data.loc[(historic_data.Variable=='GDP') & (historic_data.Type=='Service'),historic_years].values[0]
    economic_gdp.loc['Energy',historic_years]=historic_data.loc[(historic_data.Variable=='GDP') & (historic_data.Type=='Energy'),historic_years].values[0]

    economic_gdp.loc['GDP Growth rate',projection_years]=projection_data.loc[(projection_data.Variable=='GDP growth') & (projection_data.Type=='All'),projection_years].values[0]
    economic_gdp.loc['Agriculture',projection_years]=projection_data.loc[(projection_data.Variable=='GDP') & (projection_data.Type=='Agriculture'),projection_years].values[0]
    economic_gdp.loc['Construction',projection_years]=projection_data.loc[(projection_data.Variable=='GDP') & (projection_data.Type=='Construction'),projection_years].values[0]
    economic_gdp.loc['Mining',projection_years]=projection_data.loc[(projection_data.Variable=='GDP') & (projection_data.Type=='Mining'),projection_years].values[0]
    economic_gdp.loc['Manufacturing',projection_years]=projection_data.loc[(projection_data.Variable=='GDP') & (projection_data.Type=='Manufacturing'),projection_years].values[0]
    economic_gdp.loc['Service',projection_years]=projection_data.loc[(projection_data.Variable=='GDP') & (projection_data.Type=='Service'),projection_years].values[0]
    economic_gdp.loc['Energy',projection_years]=projection_data.loc[(projection_data.Variable=='GDP') & (projection_data.Type=='Energy'),projection_years].values[0]
    economic_gdp.loc['Total',historic_years+projection_years]=economic_gdp.loc['Agriculture',historic_years+projection_years]+economic_gdp.loc['Construction',historic_years+projection_years]+economic_gdp.loc['Mining',historic_years+projection_years]+economic_gdp.loc['Manufacturing',historic_years+projection_years]+economic_gdp.loc['Service',historic_years+projection_years]+economic_gdp.loc['Energy',historic_years+projection_years]
    
    ##GDP growth rate correction
    hist_gdp=historic_data.loc[(historic_data.Variable=='GDP growth') & (historic_data.Type=='All'),:]
    hist_gdp=hist_gdp.reset_index()
    proj_gdp=projection_data.loc[(projection_data.Variable=='GDP growth') & (projection_data.Type=='All'),:]
    proj_gdp=proj_gdp.reset_index()
    joined_gdp=(hist_gdp.T.append(proj_gdp.T)).T
    for i in range(0,len(years)):
        if years[i]==years[0]:
            economic_gdp.loc['GDP Growth rate',years[i]]=historic_data.loc[(historic_data.Variable=='GDP growth') & (historic_data.Type=='All'),years[i]].values[0]
        else:
            economic_gdp.loc['GDP Growth rate',years[i]]=joined_gdp.loc[0,list(range(years[i-1]+1,years[i]+1))].mean()


    for year in years:
        if np.isnan(economic_gdp.loc['GDP',year]):
            economic_gdp.loc['GDP',year]=economic_gdp.iloc[economic_gdp.index.get_loc('GDP'),economic_gdp.columns.get_loc(year)-1]*(1+economic_gdp.loc['GDP Growth rate',year]/100)**(year-economic_gdp.columns[economic_gdp.columns.get_loc(year)-1])
        if np.isnan(economic_demography.loc['Population',year]):   
            economic_demography.loc['Population',year]=economic_demography.iloc[economic_demography.index.get_loc('Population'),economic_demography.columns.get_loc(year)-1]*(1+economic_demography.loc['Population growth rate',year]/100)**(year-economic_demography.columns[economic_demography.columns.get_loc(year)-1])

    economic_gdp.loc['GDP per capita',historic_years+projection_years]=economic_gdp.loc['GDP',historic_years+projection_years]/economic_demography.loc['Population',historic_years+projection_years]
    economic_demography.loc['Number of urban Households',historic_years+projection_years]=economic_demography.loc['Population',historic_years+projection_years]*economic_demography.loc['Urban Population',historic_years+projection_years]/100/economic_demography.loc['Person/ urban Household',historic_years+projection_years]
    economic_demography.loc['Number of rural Households',historic_years+projection_years]=economic_demography.loc['Population',historic_years+projection_years]*economic_demography.loc['Rural Population',historic_years+projection_years]/100/economic_demography.loc['Person/ rural Household',historic_years+projection_years]
    economic_demography.loc['Active Labour Force',historic_years+projection_years]=economic_demography.loc['Population',historic_years+projection_years]*economic_demography.loc['Potential Labour Force',historic_years+projection_years]/100*economic_demography.loc['Participating Labour Force',historic_years+projection_years]/100
    economic_demography.loc['Population inside Large Cities',historic_years+projection_years]=economic_demography.loc['Population',historic_years+projection_years]*economic_demography.loc['Population in cities with public transport',historic_years+projection_years]/100

    #Transport freight demand tab
    transport_freight_generation.loc['Agriculture',historic_years]=1#historic_data.loc[(historic_data.Variable=='Freight Demand') & (historic_data.Type=='Agriculture'),historic_years].values[0]
    transport_freight_generation.loc['Construction',historic_years]=1#historic_data.loc[(historic_data.Variable=='Freight Demand') & (historic_data.Type=='Construction'),historic_years].values[0]
    transport_freight_generation.loc['Mining',historic_years]=1#historic_data.loc[(historic_data.Variable=='Freight Demand') & (historic_data.Type=='Mining'),historic_years].values[0]
    transport_freight_generation.loc['Manufacturing',historic_years]=1#historic_data.loc[(historic_data.Variable=='Freight Demand') & (historic_data.Type=='Manufacturing'),historic_years].values[0]
    transport_freight_generation.loc['Service',historic_years]=1#historic_data.loc[(historic_data.Variable=='Freight Demand') & (historic_data.Type=='Service'),historic_years].values[0]
    transport_freight_generation.loc['Energy',historic_years]=1#historic_data.loc[(historic_data.Variable=='Freight Demand') & (historic_data.Type=='Energy'),historic_years].values[0]
    transport_freight_generation.loc['Agriculture',projection_years]=1#projection_data.loc[(projection_data.Variable=='Freight Demand') & (projection_data.Type=='Agriculture'),projection_years].values[0]
    transport_freight_generation.loc['Construction',projection_years]=1#projection_data.loc[(projection_data.Variable=='Freight Demand') & (projection_data.Type=='Construction'),projection_years].values[0]
    transport_freight_generation.loc['Mining',projection_years]=1#projection_data.loc[(projection_data.Variable=='Freight Demand') & (projection_data.Type=='Mining'),projection_years].values[0]
    transport_freight_generation.loc['Manufacturing',projection_years]=1#projection_data.loc[(projection_data.Variable=='Freight Demand') & (projection_data.Type=='Manufacturing'),projection_years].values[0]
    transport_freight_generation.loc['Service',projection_years]=1#projection_data.loc[(projection_data.Variable=='Freight Demand') & (projection_data.Type=='Service'),projection_years].values[0]
    transport_freight_generation.loc['Energy',projection_years]=1#projection_data.loc[(projection_data.Variable=='Freight Demand') & (projection_data.Type=='Energy'),projection_years].values[0]
    
    transport_freight_generation.loc['Base value',historic_years+projection_years]=transport_freight_generation.sum(axis='rows')*economic_gdp.loc['GDP',historic_years+projection_years]/1000
    
    #Transport freight modal and intensity tabs
    freight_means=transport_freight_modal.index.values
    
    for tra in freight_means:
        transport_freight_modal.loc[tra, historic_years]=1 #historic_data.loc[(historic_data.Variable=='FILLME' & historic_data.Sub-type==tra.split(' ')[0]) & (historic_data.Fuel==tra.split(' ')[1]),historic_years].values[0]
        transport_freight_modal.loc[tra, projection_years]=1 #projection_data.loc[(projection_data.Variable=='FILLME' & projection_data.Sub-Type==tra.split(' ')[0]) & (projection_data.Fuel==tra.split(' ')[1]),projection_years].values[0]
        transport_freight_intensity.loc[tra, historic_years]=1 #historic_data.loc[(historic_data.Variable=='FILLME' & historic_data.Sub-type==tra.split(' ')[0]) & (historic_data.Fuel==tra.split(' ')[1]),historic_years].values[0]
        transport_freight_intensity.loc[tra, projection_years]=1 #projection_data.loc[(projection_data.Variable=='FILLME' & projection_data.Sub-Type==tra.split(' ')[0]) & (projection_data.Fuel==tra.split(' ')[1]),projection_years].values[0]
        if tra == freight_means[len(freight_means)-1]:
            transport_freight_modal.loc[tra, historic_years+projection_years]=transport_freight_modal.iloc[0:len(freight_means)-1].sum()
            transport_freight_modal.loc[tra, historic_years+projection_years]=100-transport_freight_modal.loc[tra, historic_years+projection_years]

    # Transport intercity tabs filling
    transport_intercity_factors.loc['Distance travelled',historic_years]=1#historic_data.loc[(historic_data.Variable=='Distance travelled') & (historic_data.Type=='Intercity'),historic_years].values[0]
    transport_intercity_factors.loc['Car ownership',historic_years]=1#historic_data.loc[(historic_data.Variable=='Car ownership')].values[0]
    transport_intercity_factors.loc['Distance travelled by car',historic_years]=1#historic_data.loc[(historic_data.Variable=='Distance travelled') & (historic_data.Sub-type=='Cars'),historic_years].values[0]
    transport_intercity_factors.loc['Cars',historic_years]=1#historic_data.loc[(historic_data.Variable=='Distance travelled') & (historic_data.Sub-type=='Cars'),historic_years].values[0]
    transport_intercity_factors.loc['Air Plane',historic_years]=1#historic_data.loc[(historic_data.Variable=='Distance travelled') & (historic_data.Sub-type=='Airplanes'),historic_years].values[0]

    transport_intercity_factors.loc['Distance travelled',projection_years]=1#projection_data.loc[(projection_data.Variable=='Distance travelled') & (projection_data.Type=='Intercity'),projection_years].values[0]
    transport_intercity_factors.loc['Car ownership',projection_years]=1#projection_data.loc[(projection_data.Variable=='Car ownership')].values[0]
    transport_intercity_factors.loc['Distance travelled by car',projection_years]=1#projection_data.loc[(projection_data.Variable=='Distance travelled') & (projection_data.Sub-type=='Cars'),projection_years].values[0]
    transport_intercity_factors.loc['Cars',projection_years]=1#projection_data.loc[(projection_data.Variable=='Load factors') & (projection_data.Sub-type=='Cars'),projection_years].values[0]
    transport_intercity_factors.loc['Air Plane',projection_years]=1#projection_data.loc[(projection_data.Variable=='Load factors') & (projection_data.Sub-type=='Airplanes'),projection_years].values[0]

    intercity_means=transport_intercity_intensity.index.values
    modal_cars_index=0
    modal_public_index=transport_intercity_modal.index.get_loc('Modal split of public intercity transportation')
    
    for tra in intercity_means:
        transport_intercity_intensity.loc[tra, historic_years]=1 #historic_data.loc[(historic_data.Variable=='FILLME' & historic_data.Sub-type==tra.split(' ')[0]) & (historic_data.Fuel==tra.split(' ')[1]),historic_years].values[0]
        transport_intercity_intensity.loc[tra, projection_years]=1 #projection_data.loc[(projection_data.Variable=='FILLME' & projection_data.Sub-Type==tra.split(' ')[0]) & (projection_data.Fuel==tra.split(' ')[1]),projection_years].values[0]
        transport_intercity_modal.loc[tra, historic_years]=1 #historic_data.loc[(historic_data.Variable=='FILLME' & historic_data.Sub-type==tra.split(' ')[0]) & (historic_data.Fuel==tra.split(' ')[1]),historic_years].values[0]
        transport_intercity_modal.loc[tra, projection_years]=1 #projection_data.loc[(projection_data.Variable=='FILLME' & projection_data.Sub-Type==tra.split(' ')[0]) & (projection_data.Fuel==tra.split(' ')[1]),projection_years].values[0]
        transport_intercity_factors.loc[tra, historic_years]=1 #historic_data.loc[(historic_data.Variable=='FILLME' & historic_data.Sub-type==tra.split(' ')[0]) & (historic_data.Fuel==tra.split(' ')[1]),historic_years].values[0]
        transport_intercity_factors.loc[tra, projection_years]=1 #projection_data.loc[(projection_data.Variable=='FILLME' & projection_data.Sub-Type==tra.split(' ')[0]) & (projection_data.Fuel==tra.split(' ')[1]),projection_years].values[0]


    transport_intercity_modal.iloc[modal_cars_index,2:]=transport_intercity_modal.iloc[modal_cars_index+1:modal_public_index,2:].sum()
    transport_intercity_modal.iloc[modal_public_index,2:]=transport_intercity_modal.iloc[modal_public_index+1:len(transport_intercity_modal),2:].sum()

    #Transport urban tabs filling
    transport_urban_factors.loc['Distance travelled',historic_years]=1#historic_data.loc[(historic_data.Variable=='Distance travelled') & (historic_data.Type=='urban'),historic_years].values[0]
    transport_urban_factors.loc['Distance travelled',projection_years]=1#projection_data.loc[(projection_data.Variable=='Distance travelled') & (projection_data.Type=='urban'),projection_years].values[0]

    urban_means=transport_urban_intensity.index.values

    for tra in urban_means:
        transport_urban_intensity.loc[tra, historic_years]=1 #historic_data.loc[(historic_data.Variable=='FILLME' & historic_data.Sub-type==tra.split(' ')[0]) & (historic_data.Fuel==tra.split(' ')[1]),historic_years].values[0]
        transport_urban_intensity.loc[tra, projection_years]=1 #projection_data.loc[(projection_data.Variable=='FILLME' & projection_data.Sub-Type==tra.split(' ')[0]) & (projection_data.Fuel==tra.split(' ')[1]),projection_years].values[0]
        transport_urban_modal.loc[tra, historic_years]=1 #historic_data.loc[(historic_data.Variable=='FILLME' & historic_data.Sub-type==tra.split(' ')[0]) & (historic_data.Fuel==tra.split(' ')[1]),historic_years].values[0]
        transport_urban_modal.loc[tra, projection_years]=1 #projection_data.loc[(projection_data.Variable=='FILLME' & projection_data.Sub-Type==tra.split(' ')[0]) & (projection_data.Fuel==tra.split(' ')[1]),projection_years].values[0]
        transport_urban_factors.loc[tra, historic_years]=1 #historic_data.loc[(historic_data.Variable=='FILLME' & historic_data.Sub-type==tra.split(' ')[0]) & (historic_data.Fuel==tra.split(' ')[1]),historic_years].values[0]
        transport_urban_factors.loc[tra, projection_years]=1 #projection_data.loc[(projection_data.Variable=='FILLME' & projection_data.Sub-Type==tra.split(' ')[0]) & (projection_data.Fuel==tra.split(' ')[1]),projection_years].values[0]


    transport_urban_modal.iloc[modal_cars_index,2:]=transport_urban_modal.iloc[modal_cars_index+1:modal_public_index,2:].sum()
    transport_urban_modal.iloc[modal_public_index,2:]=transport_urban_modal.iloc[modal_public_index+1:len(transport_intercity_modal),2:].sum()

    #International tab filling
    transport_international.loc['Constant',historic_years]=0 #Anything?
    transport_international.loc['Constant',projection_years]=0 #Anything?
    transport_international.loc['Variable',historic_years]=0 #Anything?
    transport_international.loc['Variable',projection_years]=0 #Anything?

    #Printing
    economic= economic_demography, economic_gdp
    freight = transport_freight_generation, transport_freight_intensity, transport_freight_modal
    intercity= transport_intercity_factors, transport_intercity_intensity, transport_intercity_modal
    urban= transport_urban_factors, transport_urban_intensity, transport_urban_modal
    extract_data(country, scenario, data_out, economic, freight, intercity, urban, transport_international)

    return data_out



if __name__ == "__main__":
    country = 'Vietnam' #input("Enter a country: ")
    scenario = 'BAU'#input("Enter a scenario (BAU, MS, NZ): ")
    data_out=run(country, scenario, years)
