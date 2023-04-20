import pandas as pd
import csv
import os
from math import inf
from shutil import copy
from urllib import request, error
from fnmatch import fnmatch
import warnings
import numpy as np
import json 
from datetime import datetime



# %% [markdown]
def download_sdk(country, sdk_urls_location, sdk_location):
    try:
        sdk_urls_file = open(sdk_urls_location)
        sdk_urls = { c: u for c, u in csv.reader(sdk_urls_file) }
        url_sdk = sdk_urls[country]
    except:
        raise Exception("Downloading the SDK failed. Make sure that there is a SDK for the given country and scenario.")
    try:
        print("Downloading SDK...")
        request.urlretrieve(url_sdk.replace(" ", "%20"), sdk_location)
        print("Downloading done!")
    except error.HTTPError:
        raise Exception("Downloading the SDK failed. Make sure that the SDK's url and export-location are valid.")


def get_sdk_path(country, scenario):
    sdk_filename_pattern = country + "*" + scenario + "*"
    sdk_folder = './inputs/OSeMOSYS Starter Kits/'
    sdk_file = None
    for file in os.listdir(os.path.join(sdk_folder)):
        if fnmatch(file, sdk_filename_pattern):
            sdk_file = file
            break
    return sdk_folder + sdk_file if sdk_file else None


def find_sdk_file_path(country, scenario):
    warnings.simplefilter(action='ignore', category=UserWarning)
    sdk_path = get_sdk_path(country, scenario)
    if sdk_path: 
        return sdk_path
    else:
        try:
            sdk_urls_location = f"./resources/Mapping/download_links/{scenario}_links.csv"
            sdk_location = f"./inputs/OSeMOSYS Starter Kits/{country}_{scenario}_SAND.xlsm"
            download_sdk(country, sdk_urls_location, sdk_location)
            sdk_path = get_sdk_path(country, scenario)
            return sdk_path
        except: return(print("Failed to find SDK."))


def clone_sdk_before_run(country, scenario, output_dir, sdk_file):
    try:
         if not os.path.exists(output_dir): os.makedirs(output_dir)
         runnable=copy(sdk_file,output_dir)
         return output_dir, runnable
    except:
        print('Failed to create OSeMOSYS running folder')


def run(country, scenario, output_dir):
    print("\n----------Downloading OSeMOSYS Starter Data Kits ----------\n")
    if not country or not scenario: raise Exception("You need to provide a country and a scenario.")
    sdk_file_path=find_sdk_file_path(country, scenario)
    new_dir_path, new_file_path=clone_sdk_before_run(country, scenario, output_dir, sdk_file_path)
    return output_dir, new_file_path


if __name__ == "__main__":
    country = input("Enter a country: ")
    scenario = input("Select OSeMOSYS scenario (Base, NZv1, NZv2, LCv1, LCv2, FF): ")
    output_dir = f"./runs/OSeMOSYS/{country}_{scenario}_{strftime('%Y-%m-%d_%H-%M-%S')}"
    model_dir_path, data_source_path = run(country, scenario)
    print("Model Folder Path:", model_dir_path, "\n")
    print("Data Source Path:", data_source_path, "\n")
    print(pd.read_excel(data_source_path, sheet_name='Parameters'))
    