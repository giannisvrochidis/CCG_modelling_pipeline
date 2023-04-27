#%%
import pandas as pd

input_path = "./MAED_inputs.xlsx"
output_path = "./MAED_IEA_template.xlsx"

input_dict = pd.read_excel(input_path, sheet_name=None)
output_dict = pd.read_excel(output_path, sheet_name=None)

for output_sheet_name, output_sheet_data in output_dict.items():
    output_properties = output_sheet_data['Property']
    for output_row_index, property in output_properties.items():
        for input_sheet_name, input_sheet_data in input_dict.items():
            try: input_row_index = input_sheet_data.loc[input_sheet_data.iloc[:, 1] == property].index[0]
            except: continue
            property_row = input_sheet_data.iloc[input_row_index, 1:]
            output_sheet_data.iloc[output_row_index, 1:] = property_row
    with pd.ExcelWriter(output_path, engine="openpyxl", mode="a") as writer:
        writer.book.remove(writer.book[output_sheet_name])
        output_sheet_data.to_excel(writer, sheet_name=output_sheet_name, index=False)

# %%
