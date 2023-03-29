import download_sand
import pandas as pd
import xlwings as xw
from utils import read_configuration, format_path

def read_maed_results_IEA(maed_results, maed_years, first_col, last_col):
    wb1 = xw.Book(maed_results)
    sheet_1 = wb1.sheets['Sheet1']
    range_1 = sheet_1.range((  1, first_col), (1, last_col))
    maed_outputs=pd.DataFrame(range_1.value)
    maed_outputs.index=maed_years
    maed_outputs=maed_outputs.T
    maed_outputs.loc['TRACAR',:]=sheet_1.range((335, first_col), (335, last_col)).value+(pd.DataFrame(sheet_1.range((410, first_col), (412, last_col)).value).fillna(0).sum().T).values
    maed_outputs.loc['TRABUS',:]=(pd.DataFrame(sheet_1.range((339, first_col), (342, last_col)).value).fillna(0).sum().T).values+(pd.DataFrame(sheet_1.range((413, first_col), (415, last_col)).value).fillna(0).sum().T).values+sheet_1.range((409, first_col), (409, last_col)).value
    maed_outputs.loc['INDELC',:]=sheet_1.range((264, first_col), (264, last_col)).value
    maed_outputs.loc['INDHEHmult',:]=(pd.DataFrame(sheet_1.range((206, first_col), (207, last_col)).value).fillna(0).sum().T).values+sheet_1.range((209, first_col), (209, last_col)).value
    maed_outputs.loc['INDHEH',:]=sheet_1.range((187, first_col), (187, last_col)).value*maed_outputs.loc['INDHEHmult',:]
    maed_outputs=maed_outputs.drop(index=[0,'INDHEHmult'])
    maed_outputs.loc['INDHEL',:]=(pd.DataFrame(sheet_1.range((93, first_col), (113, last_col)).value).fillna(0).sum().T).values-sheet_1.range((97, first_col), (97, last_col)).value-sheet_1.range((104, first_col), (104, last_col)).value-sheet_1.range((111, first_col), (111, last_col)).value
    maed_outputs.loc['RESCKN',:]=pd.DataFrame(sheet_1.range((661, first_col), (661, last_col)).value).T.values-(sheet_1.range((658, first_col), (658, last_col)).value)
    maed_outputs.loc['RESHEL',:]=pd.DataFrame(sheet_1.range((646, first_col), (646, last_col)).value).T.values+sheet_1.range((654, first_col), (654, last_col)).value-sheet_1.range((642, first_col), (642, last_col)).value-sheet_1.range((650, first_col), (650, last_col)).value
    maed_outputs.loc['RESELC',:]=sheet_1.range((751, first_col), (751, last_col)).value
    maed_outputs.loc['COMELC',:]=sheet_1.range((677, first_col), (677, last_col)).value
    maed_outputs.loc['COMHEL',:]=pd.DataFrame(sheet_1.range((741, first_col), (741, last_col)).value).T.values-(sheet_1.range((738, first_col), (738, last_col)).value)
    wb1.close()
    return maed_outputs

def read_maed_results_TSDK(maed_results, maed_years, first_col, last_col):
    wb1 = xw.Book(maed_results)
    sheet_1 = wb1.sheets['Sheet1']
    range_1 = sheet_1.range((  1, first_col), (1, last_col))
    maed_outputs=pd.DataFrame(range_1.value)
    maed_outputs.index=maed_years
    maed_outputs=maed_outputs.T
    maed_outputs.loc['TRACAR',:]=sheet_1.range((339, first_col), (339, last_col)).value+(pd.DataFrame(sheet_1.range((447, first_col), (449, last_col)).value).fillna(0).sum().T).values
    maed_outputs.loc['TRABUS',:]=(pd.DataFrame(sheet_1.range((348, first_col), (350, last_col)).value).fillna(0).sum().T).values+(pd.DataFrame(sheet_1.range((442, first_col), (444, last_col)).value).fillna(0).sum().T).values
    maed_outputs.loc['INDELC',:]=sheet_1.range((258, first_col), (258, last_col)).value
    maed_outputs.loc['INDHEHmult',:]=(pd.DataFrame(sheet_1.range((200, first_col), (201, last_col)).value).fillna(0).sum().T).values+sheet_1.range((203, first_col), (203, last_col)).value
    maed_outputs.loc['INDHEH',:]=sheet_1.range((181, first_col), (181, last_col)).value*maed_outputs.loc['INDHEHmult',:]
    maed_outputs=maed_outputs.drop(index=[0,'INDHEHmult'])
    maed_outputs.loc['INDHEL',:]=(pd.DataFrame(sheet_1.range((88, first_col), (106, last_col)).value).fillna(0).sum().T).values-sheet_1.range((98, first_col), (98, last_col)).value-sheet_1.range((91, first_col), (91, last_col)).value-sheet_1.range((105, first_col), (105, last_col)).value
    maed_outputs.loc['RESCKN',:]=pd.DataFrame(sheet_1.range((584, first_col), (584, last_col)).value).T.values-(sheet_1.range((581, first_col), (581, last_col)).value)
    maed_outputs.loc['RESHEL',:]=pd.DataFrame(sheet_1.range((569, first_col), (569, last_col)).value).T.values+sheet_1.range((577, first_col), (577, last_col)).value-sheet_1.range((565, first_col), (565, last_col)).value-sheet_1.range((573, first_col), (573, last_col)).value
    maed_outputs.loc['RESELC',:]=sheet_1.range((525, first_col), (525, last_col)).value
    maed_outputs.loc['COMELC',:]=sheet_1.range((653, first_col), (653, last_col)).value
    maed_outputs.loc['COMHEL',:]=pd.DataFrame(sheet_1.range((643, first_col), (643, last_col)).value).T.values-(sheet_1.range((640, first_col), (640, last_col)).value)
    wb1.close()
    return maed_outputs

def prepare_sand_df(maed_outputs, maed_years, start_year, end_year, data_source_path):
    df=pd.read_excel(data_source_path, sheet_name='Parameters')
    for row in maed_outputs.index:
        parameter=pd.Series(list(maed_outputs.loc[row,:]))
        ind=maed_outputs.loc[row,:].name
        parameter.index=maed_years
        parameter_out=pd.Series(range(start_year, end_year+1))
        parameter_out.index=parameter_out.values
        parameter_out.loc[:]=None
        parameter_out.loc[maed_years]=parameter
        parameter_out=parameter_out.interpolate(method='linear')
        parameter_out=pd.DataFrame(parameter_out).T
        parameter_out=parameter_out.fillna(parameter[maed_years[0]])
        if ind in (['INDELC','COMELC','RESELC']):
            df.loc[(df.FUEL==ind) & (df.Parameter=='SpecifiedAnnualDemand'),["%02d" %i for i in range(2015,2071)]]=parameter_out.values
        else:
            df.loc[(df.FUEL==ind) & (df.Parameter=='AccumulatedAnnualDemand'),["%02d" %i for i in range(2015,2071)]]=parameter_out.values
    return df

def write_sand(df, data_source_path):

    wb = xw.Book(data_source_path)
    sheet = wb.sheets['Parameters']
    sheet.clear()
    sheet.range('A1').value = df.columns.values
    sheet.range('A2').value = df.values
    last_row = sheet.cells.last_cell.row
    last_col = sheet.cells.last_cell.column
    data_range = sheet.range((1, 1), (last_row, 66))
    table_name = "Table1" 
    sheet.api.ListObjects.Add(1, data_range.api, 0, 1, table_name)
    wb.save(data_source_path)
    wb.close()
    return data_source_path


def run(country, scenario, maed_results,maed_years, selected_option):
    model_dir_path, data_source_path = download_sand.run(country, scenario)
    first_col=2
    last_col=first_col+len(maed_years)-1
    if selected_option == 'IEA':
        maed_outputs=read_maed_results_IEA(maed_results, maed_years, first_col, last_col)
    else:
        maed_outputs=read_maed_results_TSDK(maed_results, maed_years, first_col, last_col)
    start_year = 2015
    end_year=2070
    df= prepare_sand_df(maed_outputs, maed_years, start_year, end_year, data_source_path)
    write_sand(df, data_source_path)
    return model_dir_path, data_source_path

if __name__ == "__main__":
    country = input("Enter a country: ")
    scenario = input("Select OSeMOSYS scenario (Base, NZ, LC, FF): ")
    maed_results=f"./resources/maed-2.0.0/maed_results.xlsx"
    config = read_configuration("maed")
    maed_years = config["years"]
    selected_option, _ = pick(["Transport Starter Kits", "IEA"], "Which MAED template do you want to use?", indicator='=>')
    maed_years=[2022, 2025, 2030, 2035, 2040, 2045, 2050]
    run(country, scenario, maed_results,maed_years,selected_option)