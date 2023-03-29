from datetime import datetime
from utils import read_configuration, run_excel_macro, execute_program, write_to_file, open_work_book, format_path
import pandas as pd
import numpy as np
import sys
from pathlib import Path


def extract_data_from_xls(data_source):
    print("Start extracting data from data source...")
    try:
        with open_work_book(data_source) as wb:
            run_excel_macro(wb, "Module1.writefile")
    except Exception as exc:
        print("Error extracting data:")
        print(str(exc))
        return False
    print("Extracting data done!")
    return True


def run_glpsol(path, data_file_name, model_file, lp_file_name):
    print("Starting glpsol. Please wait...")
    glpsol_path = format_path(path, "Utils", "glpsol.exe")
    glpsol_args = "--check -m \"{0}\" -d \"{1}\" --wlp \"{2}\"".format(
        model_file, data_file_name, lp_file_name
    )
    return run_process(".", glpsol_path, glpsol_args)


def run_cbc(path, input_file_name, output_file_name, cbc_ratio):
    print("Starting CBC. Please wait...")
    cbc_path = format_path(path, "CBC", "bin", "cbc.exe")
    cbc_ratio_option = "ratio " + cbc_ratio if cbc_ratio else ""
    cbc_args = "\"{0}\" {1} solve -solu \"{2}\"".format(
        input_file_name, cbc_ratio_option, output_file_name
    )
    return run_process(".", cbc_path, cbc_args)


def run_process(dir_name, file_name, args):
    try:
        print("Running {0} {1}".format(file_name, args))
        completed, logs = execute_program(dir_name, file_name, args)
        print("Printing logs...\n")
        print(logs.decode("utf-8"))
        return completed
    except Exception as exc:
        print("Error " + str(exc))


def check_for_all_zeros(df):
    df.loc[:, (df == 0).all(axis=0)] = np.nan
    return df

def replace_result_value(df):
    df[df.loc[:, ~df.isnull().all()].iloc[:, -2].name] = np.nan
    return df

def python_converter(country, scenario, input, output_dir):
    output_filename = f"{country}_{scenario}_processed_results.csv"
    output_dir = r"{}".format(output_dir)
    input = r"{}".format(input)

    columns = [
        "index",
        "Variable",
        "Dim1",
        "Dim2",
        "Dim3",
        "Dim4",
        "Dim5",
        "Dim6",
        "Dim7",
        "Dim8",
        "Dim9",
        "Dim10",
        "ResultValue",
    ]

    osemosys_output = pd.read_csv(
        input, names=columns, sep="\(|,|\)|[ \t]{1,}", engine="python"
    )
    osemosys_output = osemosys_output[osemosys_output["index"] != "Optimal"]

    osemosys_clean = osemosys_output.groupby("Variable").apply(
        lambda x: check_for_all_zeros(x)
    )

    osemosys_clean["ResultValue"] = osemosys_clean.ffill(axis=1).iloc[:, -1]

    osemosys_cleaned = osemosys_clean.groupby("Variable").apply(
        lambda x: replace_result_value(x)
    )

    osemosys_cleaned = osemosys_cleaned.drop("index", axis=1)

    output_directory = Path(output_dir) / Path(output_filename)

    # output_directory = '"{}"'.format(output_directory)
    osemosys_cleaned.to_csv(
        output_directory,
        index=False,
    )
    return output_directory

def run(country, scenario, model_dir_path, data_source_path):
    print("\n---------- OSEMOSYS ----------\n")

    config = read_configuration("osemosys")

    path = format_path("./resources/clicSAND-1.1")
    model_folder = model_dir_path if model_dir_path else format_path(config["model_folder"])
    data_source = data_source_path if data_source_path else format_path(config["data_source"])
    model_file = format_path(config["model_file"])
    cbc_ratio = config["cbc_ratio"]

    data_file_name = format_path(data_source + ".txt")
    lp_file_name = format_path(data_source + ".lp")
    results_file_name = format_path(data_source + ".results.txt")

    print("Data file: " + data_file_name)
    print("Model file: " + model_file)
    print("GLPSOL Output file: " + lp_file_name)
    print("Results file: " + results_file_name)
    print()

    try:
        result = False
        result = extract_data_from_xls(data_source)
        if result: result = run_glpsol(path, data_file_name, model_file, lp_file_name)
        if result: result = run_cbc(path, lp_file_name, results_file_name, cbc_ratio)
        input = results_file_name
        output_dir = model_folder
        output_csv = python_converter(country, scenario, input, output_dir)
        print("OSEMOSYS Done!")

    except Exception as exc:
        print(str(exc) + "Error Running Model")
    return output_csv

if __name__ == "__main__":
    run(None, None)
