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
    #MAED
    maed_type, selected_option, maed_scenario, maed_years = maed_config.run(country)
    maed_results =maed.run(country, maed_type, selected_option, maed_scenario, maed_years)
    
    #OSeMOSYS
    scenario = input("Select OSeMOSYS scenario (Base, NZv1, NZv2, LCv1, LCv2, FF): ")
    model_dir_path, data_source_path=maed_to_sand.run(country, scenario, maed_results,maed_years,selected_option)
    print(data_source_path)
    # output_csv=run(country, scenario, model_dir_path, data_source_path)
    output_csv=f"./testing/SDK_dummy_results.csv"

    #FlexTool 2.0
    sdk_to_ft.run(country, scenario, output_csv, data_source_path)
    flextool.run()

elif selected_option == f'{osemosys_option} -> {flextool_option}':
    #OSeMOSYS
    scenario = input("Select OSeMOSYS scenario (Base, NZv1, NZv2, LCv1, LCv2, FF): ")
    model_dir_path, data_source_path = download_sand.run(country, scenario)
    output_csv=osemosys.run(country, scenario, model_dir_path, data_source_path)
    # output_csv=f"./testing/SDK_dummy_results.csv"

    #FlexTool 2.0
    sdk_to_ft.run(country, scenario, output_csv, data_source_path)
    flextool.run()
    
elif selected_option == f'{maed_option} -> {osemosys_option}':
    maed_type, selected_option, maed_scenario, maed_years = maed_config.run(country)
    maed_results =maed.run(country, maed_type, selected_option, maed_scenario, maed_year,selected_option)

    scenario = input("Select OSeMOSYS scenario (Base, NZv1, NZv2, LCv1, LCv2, FF): ")
    model_dir_path, data_source_path=maed_to_sand.run(country, scenario, maed_results,maed_years)
    output_csv=osemosys.run(country, scenario, model_dir_path, data_source_path)
    # output_csv=f"./testing/SDK_dummy_results.csv"
elif selected_option == flextool_option:
    scenario = input("Select SDK scenario (Base, NZv1, NZv2, LCv1, LCv2, FF): ")
    data_to_FT.run(country, scenario)
    flextool.run()
elif selected_option == osemosys_option:
    scenario = input("Select OSeMOSYS scenario (Base, NZv1, NZv2, LCv1, LCv2, FF): ")
    model_dir_path, data_source_path = download_sand.run(country, scenario)
    osemosys.run(country, scenario, model_dir_path, data_source_path)

elif selected_option == maed_option:        
    maed_type, selected_option, maed_scenario, maed_years = maed_config.run(country)
    maed_results =maed.run(country, maed_type, selected_option, maed_scenario, maed_years)