import pandas as pd
import shutil

sectors_to_prop = {
    "Industry":"AAA_2",
    "Transport":"S_2",
    "Services":"S2_2",
    "Households":"S1_2",
    "Agriculture":"S3_2"
}

input_path = "./MAEDEL-coefficients.xlsx"
template_path = "./maedel_template.xlsx"
output_path= "./maedel_input.xlsx"
shutil.copy(template_path, output_path)

input_df = pd.read_excel(input_path, sheet_name=None)
output_df = pd.read_excel(output_path, sheet_name=None)

for input_sheet_name, input_sheet_df in input_df.items():
    input_params = input_sheet_name.split("_")
    sector_in, type = input_params[0], input_params[1]
    if type != 'daily' and type != 'hourly':
        output_sheet_df = input_sheet_df
        output_sheet_name = input_sheet_name
        if type == 'weekly': output_sheet_name = f"coefficient_{type}-{sectors_to_prop[sector_in]}"
        
    else:
        year = input_sheet_df.iloc[0, 0]
        column_df = input_sheet_df.iloc[1:, 1:].stack().reset_index(drop=True).to_frame(name=year)
        output_sheet_name = f"coefficient_{type}-{sectors_to_prop[sector_in]}-{year}"
        output_sheet_df = pd.concat([output_df[output_sheet_name], column_df], axis=1)
    with pd.ExcelWriter(output_path, engine="openpyxl", mode="a") as writer:
        writer.book.remove(writer.book[output_sheet_name])
        output_sheet_df.to_excel(writer, sheet_name=output_sheet_name, index=False)



