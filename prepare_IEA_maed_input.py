import pandas as pd
import xlwings as xw
from utils import format_path
import shutil

def run(country):
    app = xw.App(visible=False)
    map=pd.read_csv(format_path(f"./resources/maed-2.0.0//inputs/mapping.csv"))
    for i in range(0,len(map)):

        new_input_file_sheet_name=map.loc[i,'output_sheet']
        maed_inputs_sheet_name=map.loc[i,'input_sheet']

        path = f"./resources/maed-2.0.0//inputs/"

        # Define the source and destination file paths and create new input file
        maed_inputs_path = path+map.loc[i,'input_book']
        IEA_template_path = path+map.loc[i,'output_book']

        input_file_path="./resources/maed-2.0.0/inputs/models/"+country+'_IEA'+'.xlsx'
        input_file=shutil.copy(IEA_template_path,input_file_path)
        # Define the range to copy from the MAED_inputs file

        range_1 = map.loc[i,'input_range']

        # Define the range to paste to in the Combined file

        range_2 = map.loc[i,'output_range']

        # Load the source and destination workbooks
        wb1 = xw.Book(maed_inputs_path)
        # Open the destination workbook
        wb2 = xw.Book(input_file)

        # Get the source range
        sheet_1 = wb1.sheets[maed_inputs_sheet_name]
        range_1 = sheet_1.range(range_1)

        # Get the destination range
        sheet_2 = wb2.sheets[new_input_file_sheet_name]
        range_2 = sheet_2.range(range_2)

        # Copy values from the source range to the destination range
        range_2.value = range_1.value

        # Close the workbooks
        wb1.close()
        wb2.save()
        wb2.close()

    app.kill()
    return input_file

if __name__ == "__main__":
    country = input("Enter a country: ")
    run(country)