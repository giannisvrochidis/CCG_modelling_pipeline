from utils import read_configuration, open_work_book, get_sheet, set_checkbox_value, set_cell_values, run_excel_macro, format_path
from shutil import copy
import PeriodsSelection

settings_and_filters_cells = {
    "leave_out_nodes": "C3",
    "leave_out_grids": "C4",
    "time_series_filter": "C5",
    "model_file": "C6",
    "solver": "C7",
    "clp_option": "C8",
    "clear_res_folder": "C9",
    "use_wtee": "C10",
    "max_parallel": "C11",
    "input_folder": "C12",
    "time_series_folder": "C13",
    "results_folder": "C14",
    "plot_start": "C15",
    "plot_length": "C16",
}

modelling_process_settings_cells = {
    "active_input_files": "B15",
    "active_scenarios": "F15",
}

modelling_process_settings_map_to_checkbox = {
    "leave_results_file_open": "CheckBox1",
    "import_results_after_optim": "CheckBox3",
    "plots_in_results_file": "CheckBox2",
    "parallel_calculation": "CheckBox5",
    "run_in_background": "CheckBox4",
}


def configure_modelling_process_options(sheet, settings):
    for setting_name, value in settings.items():
        if (setting_name in modelling_process_settings_map_to_checkbox):
            checkbox_name = modelling_process_settings_map_to_checkbox[setting_name]
            set_checkbox_value(sheet, checkbox_name, value)
        else:
            initial_cell = modelling_process_settings_cells[setting_name]
            set_cell_values(sheet, initial_cell, value, "column")


def configure_settings_and_filters(sheet, settings):
    for name, value in settings.items():
        initial_cell = settings_and_filters_cells[name]
        set_cell_values(sheet, initial_cell, value, "row")


def configure_settings(wb, config):
    configure_modelling_process_options(
        get_sheet(wb, "Sensitivity scenarios"),
        config["modelling_process_options"]
    )
    configure_settings_and_filters(
        get_sheet(wb, "Settings and filters"),
        config["settings_and_filters"]
    )


def run():
    print("\n---------- FlexTool v2.0 ----------\n")
    config = read_configuration("flextool")
    template_path = "./resources/flexTool-v2.0/flexTool_template.xlsm"
    path = "./resources/flexTool-v2.0/flexTool.xlsm"
    copy(template_path, path)
    print("Opening FlexTool...")
    command = pick(["Full year", "Representative weeks"], "Do you want the simulation horizon to include the whole year or representative weeks?", indicator='=>')[0]
    PeriodsSelection.run(command)
    wb = open_work_book(path)
    print("Setting up FlexTool...")
    configure_settings(wb, config)
    print("Start writting time series and run...")
    run_excel_macro(wb, "run_scenarios_module.write_ts_and_run")
    wb.save()
    wb.close()
    print("FlexTool started!")


if __name__ == "__main__":
    run()
