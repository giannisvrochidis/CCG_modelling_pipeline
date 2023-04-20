import maed
import osemosys
import flextool
import sdk_to_ft
import download_sand
import data_to_FT
import maed_config
from pick import pick
import maed_to_sand
from download_sand import clone_sdk_before_run
from time import strftime


# CLI Menu
maed_option = 'MAED'
osemosys_option = 'Osemosys'
flextool_option = 'FlexTool'
title = 'Please choose which parts of the pipeline you want to execute: '
options = [
    f'{maed_option} -> {osemosys_option} -> {flextool_option}',
    f'{maed_option} -> {osemosys_option}',
    f'{osemosys_option} -> {flextool_option}',
    maed_option,
    osemosys_option,
    flextool_option,
    'Exit'
]
selected_option, _ = pick(options, title, indicator='=>')

# Handle menu selection
if selected_option == 'Exit': 
    exit()

country = input("Enter a country: ")

if selected_option == f'{maed_option} -> {osemosys_option} -> {flextool_option}':
    
    output_dir = f"./runs/MAED-OSeMOSYS-FlexTool/{country}_{strftime('%Y-%m-%d_%H-%M-%S')}"

    #MAED
    maed_type, selected_option, maed_scenario, maed_years = maed_config.run(country)
    maed_results =maed.run(country, maed_type, selected_option, maed_scenario, maed_years, output_dir)
    
    #OSeMOSYS
    scenario = input("Select OSeMOSYS scenario (Base, NZv1, NZv2, LCv1, LCv2, FF): ")
    model_dir_path, data_source_path=maed_to_sand.run(country, scenario, maed_results,maed_years,selected_option,output_dir)
    output_csv=osemosys.run(country, scenario, model_dir_path, data_source_path)
    # output_csv=f"./testing/SDK_dummy_results.csv"

    #FlexTool 2.0
    sdk_to_ft.run(country, scenario, output_csv, data_source_path, output_dir)
    print(f'FlexTool will now start. You can save the output file which will automatically open when FlexTool finishes in the directory: {output_dir}')
    flextool.run()

elif selected_option == f'{osemosys_option} -> {flextool_option}':
    
    output_dir = f"./runs/OSeMOSYS-FlexTool/{country}_{strftime('%Y-%m-%d_%H-%M-%S')}"

    #OSeMOSYS
    scenario = input("Select OSeMOSYS scenario (Base, NZv1, NZv2, LCv1, LCv2, FF): ")
    model_dir_path, data_source_path = download_sand.run(country, scenario, output_dir)
    output_csv=osemosys.run(country, scenario, model_dir_path, data_source_path)
    # output_csv=f"./testing/SDK_dummy_results.csv"

    #FlexTool 2.0
    sdk_to_ft.run(country, scenario, output_csv, data_source_path, output_dir)
    print(f'FlexTool will now start. You can save the output file which will automatically open when FlexTool finishes in the directory: {output_dir}')
    flextool.run()
    
elif selected_option == f'{maed_option} -> {osemosys_option}':
    
    output_dir = f"./runs/MAED-OSeMOSYS/{country}_{strftime('%Y-%m-%d_%H-%M-%S')}"

    #MAED
    maed_type, selected_option, maed_scenario, maed_years = maed_config.run(country)
    maed_results =maed.run(country, maed_type, selected_option, maed_scenario, maed_years, output_dir)
    
    #OSeMOSYS
    scenario = input("Select OSeMOSYS scenario (Base, NZv1, NZv2, LCv1, LCv2, FF): ")
    model_dir_path, data_source_path=maed_to_sand.run(country, scenario, maed_results,maed_years,selected_option,output_dir)
    output_csv=osemosys.run(country, scenario, model_dir_path, data_source_path)    

elif selected_option == flextool_option:
    scenario = input("Select SDK scenario (Base, NZv1, NZv2, LCv1, LCv2, FF): ")
    data_to_FT.run(country, scenario)
    flextool.run()
elif selected_option == osemosys_option:
    output_dir = f"./runs/OSeMOSYS/{country}_{strftime('%Y-%m-%d_%H-%M-%S')}"
    #OSeMOSYS
    scenario = input("Select OSeMOSYS scenario (Base, NZv1, NZv2, LCv1, LCv2, FF): ")
    model_dir_path, data_source_path=maed_to_sand.run(country, scenario, maed_results,maed_years,selected_option,output_dir)
    output_csv=osemosys.run(country, scenario, model_dir_path, data_source_path)

elif selected_option == maed_option:        
        
    output_dir = f"./runs/MAED/{country}_{strftime('%Y-%m-%d_%H-%M-%S')}"

    #MAED
    maed_type, selected_option, maed_scenario, maed_years = maed_config.run(country)
    maed_results =maed.run(country, maed_type, selected_option, maed_scenario, maed_years, output_dir)