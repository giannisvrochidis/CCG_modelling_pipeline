import pandas as pd
from pandas import DataFrame
import numpy as np
import os
from utils import read_configuration,format_path
import openpyxl

def read_data(flextool_input):
    data=pd.DataFrame(index=range(0,8760), columns=['Demand','Wind','Solar','Hydro'])
    flextool_input=os.getcwd()+'\\resources\\flexTool-v2.0\\InputData\\model_to_run.xlsx'
    data.loc[:,'Demand']=pd.read_excel(flextool_input,sheet_name='ts_energy').loc[2:,'elec'].values
    cf=pd.read_excel(flextool_input,sheet_name='ts_cf').loc[1:,:].reset_index()
    inflow=pd.read_excel(flextool_input,sheet_name='ts_inflow').loc[1:,:].reset_index()
    caps=pd.read_excel(flextool_input,sheet_name='units').loc[:,['unit type','cf profile','inflow','capacity (MW)']].reset_index()
    caps.index=caps['unit type']
    data.loc[:,'Wind']=caps.loc['PWRWND001','capacity (MW)']*cf[caps.loc['PWRWND001','cf profile']]+caps.loc['PWRWND002','capacity (MW)']*cf[caps.loc['PWRWND002','cf profile']]+caps.loc['PWRWND001S','capacity (MW)']*cf[caps.loc['PWRWND001S','cf profile']]
    data.loc[:,'Solar']=caps.loc['PWRSOL001','capacity (MW)']*cf[caps.loc['PWRSOL001','cf profile']]+caps.loc['PWRCSP002','capacity (MW)']*cf[caps.loc['PWRCSP002','cf profile']]+caps.loc['PWRSOL001S','capacity (MW)']*cf[caps.loc['PWRSOL001S','cf profile']]
    data.loc[:,'Hydro']=caps.loc['PWRHYD001','capacity (MW)']*inflow[caps.loc['PWRHYD001','inflow']]+caps.loc['PWRHYD002','capacity (MW)']*inflow[caps.loc['PWRHYD002','inflow']]+caps.loc['PWRHYD002','capacity (MW)']*inflow[caps.loc['PWRHYD002','inflow']]
    return data

def week_locator(position):
     for x in range(0,8760,168):
        if x >position:
            week_start=x-168
            break
        elif x==position:
            week_start=x
            break
     return week_start
 
def ts_creator(ts, week, lenght):
    if week+lenght>8759:
        ts[week:8759]=1
    else:
        ts[week:week+lenght]=1     
        
    return ts
        

def select_weeks(ts, flextool_input,data):

    """
    Read data from .csv file and create the relevant time series for the analysis
    """
    Demand=pd.to_numeric(data.Demand)
    VRE=pd.to_numeric(data.Wind+data.Solar+data.Hydro)
    VRE_Demand_ratio=pd.to_numeric(VRE/Demand)
    Net_Load=pd.to_numeric((Demand-VRE).abs())
    ramp_ts=pd.to_numeric(pd.Series(np.zeros(8760)))
    for i in range(8760):
        if i==0:
            ramp_ts[i]=0
        else:
            ramp_ts[i]=(Net_Load[i]-Net_Load[i-1])
    
    ramp_ts=ramp_ts.abs()
    
        
    """
    1. Peak demand week
    """
    
    max_value_D=Demand.max()
    max_position_D=Demand.idxmax()
    
    week_start_1=week_locator(max_position_D)
         
    ts=ts_creator(ts,week_start_1,168)    
    
    """
    2. Lowest demand week
    """
    min_value_D=Demand.min()
    min_position_D=Demand.idxmin()
    
    week_start_2=week_locator(min_position_D)
    
          
    ts=ts_creator(ts,week_start_2,168)
       
    """
    3. Highest VRE
    """
    max_value_VRE=VRE.max()
    max_position_VRE=VRE.idxmax()
    
    week_start_3=week_locator(max_position_VRE)
    
    ts=ts_creator(ts,week_start_3,168)
    
    """
    4. Lowest VRE
    """
    min_value_VRE=VRE.min()
    min_position_VRE=VRE.idxmin()
    
    week_start_4=week_locator(min_position_VRE)
    
    ts=ts_creator(ts,week_start_4,168)
    
    """
    5. Highest VRE/Demand ratio
    """
    max_value_ratio=VRE_Demand_ratio.max()
    max_position_ratio=VRE_Demand_ratio.idxmax()
    
    week_start_5=week_locator(max_position_ratio)
    
    ts=ts_creator(ts,week_start_5,168)
    
    """
    6. Lowest VRE/Demand ratio
    """
    min_value_ratio=VRE_Demand_ratio.min()
    min_position_ratio=VRE_Demand_ratio.idxmin()
    
    week_start_6=week_locator(min_position_ratio)
    
    ts=ts_creator(ts,week_start_6,168)
    
    """
    7. Highest Net Demand
    """
    max_value_net=Net_Load.max()
    max_position_net=Net_Load.idxmax()
    
    week_start_7=week_locator(max_position_net)
    
    ts=ts_creator(ts,week_start_7,168)
    
    """
    8. Lowest net demand
    """
    min_value_net=Net_Load.min()
    min_position_net=Net_Load.idxmin()
    
    week_start_8=week_locator(min_position_net)
                
    ts=ts_creator(ts,week_start_8,168)
    
    """
    9. Maximum ramp of net load
    """
    max_value_ramp=ramp_ts.max()
    max_position_ramp=ramp_ts.idxmax()
    
    week_start_9=week_locator(max_position_ramp)   
    
    ts=ts_creator(ts, week_start_9,168)
    return ts
    
    """
    Writing results into a new csv file
    """
def write_weeks(ts,flextool_input):
    ts=ts.astype(int)
    writer = pd.ExcelWriter(flextool_input, engine='openpyxl', mode="a", if_sheet_exists="overlay")
    writer.workbook = openpyxl.load_workbook(flextool_input)
    ts.to_excel(writer, sheet_name='ts_time', header=False, index=False, startrow=1, startcol=1)
    ts.to_excel(os.getcwd()+'\\resources\\flexTool-v2.0\\InputData\\periods_selection.xlsx')
    writer.close()

    
    
def run(command):
    config = read_configuration("flextool")
    flextool_input=os.getcwd()+'\\resources\\flexTool-v2.0\\InputData\\'+config['modelling_process_options']['active_input_files']
    ts=pd.Series(np.zeros(8760))
    if command=='Full year':
        ts[:]=1
    else:
        data=read_data(flextool_input)
        ts=select_weeks(ts,flextool_input,data)    
    write_weeks(ts,flextool_input)
    
if __name__ == '__main__':
    command = input("Do you want the simulation horizon to include the whole year or representative weeks? For whole year press f...")
    run(command)



