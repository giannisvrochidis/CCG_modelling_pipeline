import requests
import subprocess
import pandas as pd
import os
import json
from utils import read_configuration, format_path
import prepare_maed_TSDK, prepare_IEA_maed_input
import maed_config
import webbrowser
from shutil import copytree
from time import strftime
from io import StringIO
from shutil import copy
from pick import pick

PORT = "8765"

# Convert from intermediate excel to json

def create_data_from_template(df, sheet_name):
    years = list(df.columns[2:])
    total_props = df.shape[0]
    if sheet_name=='industry_manufacturing_penetrat': sheet_name='industry_manufacturing_penetration'
    try:
        info = sheet_name.split("-")
        id = info[0]
        client = info[1]
    except: 
        client = ""
    
    data = { "SID": "1" }
    for p in range(total_props):
        for year in years:
                prop = str(df['Property'][p]).replace("{Y}", str(year)).replace("{C}", str(client))
                value = str(df[year][p])
                if ("." in prop):
                    parent_prop, child_prop = prop.split(".")
                    if(parent_prop not in data):
                        data[parent_prop] = { "SID": f"{client}_{year}" }
                    data[parent_prop][child_prop] = value
                else:
                    data[prop] = value

    if client and "weekly" not in sheet_name: 
        data["SID"] = f"{client}_{year}"

    return json.dumps(data), id


# API Calls

def prepend_base(endpoint):
    return "http://localhost:" + PORT + endpoint

def get_php_session_id():
    check_login_url = prepend_base("/auth/login/checklogin.php")
    return requests.get(check_login_url).cookies["PHPSESSID"]


def get_cookies(phpsessid, maed_type, case_study_name):
    return {
        "PHPSESSID": phpsessid,
        "maedtype": maed_type,
        "titlecs": case_study_name,
        "decimal":"5",
        "id":"",
        "l":"0",
        "langCookie":"en",
    }


def edit_maedd_general_info(years, cookies):
    url = prepend_base(f"/app/geninf/maedd_geninf.php")
    data = {
        "id": "1",
        "studyName": cookies["titlecs"],
        "Year": ",".join(map(str, years)),
        "populationunit": "Million",
        "Gdpunit": "Billion",
        "Currency": "US$",
        "energyunit": "PJ",
        "punit": "Million",
        "funit": "Million",
        "Desc": None,
        "action": "update"
    }
    requests.post(url, data=data, cookies=cookies)

def edit_maedel_general_info(years, cookies):
    url = prepend_base(f"/app/geninf/maedel_geninf.php")
    data = {
        "id": "1",
        "studyName": cookies["titlecs"],
        "Year": ",".join(years),
        "Desc": None,
        "action": "update"
    }
    requests.post(url, data=data, cookies=cookies)


def edit_maed_data(maed_input, cookies):
    df = pd.read_excel(maed_input, None)
    url = prepend_base(f"/app/data/{cookies['maedtype']}_data.php") 
    for sheet_name in df.keys():
        data, id = create_data_from_template(df[sheet_name], sheet_name)
        data={"data": data, "datanotes":None, "id":id, "action":"edit"}
        requests.post(url, data=data, cookies=cookies)

def calculate(cookies):
    url = prepend_base(f"/app/calculation/{cookies['maedtype']}_calculation.php")
    requests.post(url, cookies=cookies)


def download_results(cookies):
    export_results_url = prepend_base(f"/app/results/{cookies['maedtype']}_results_export.php")
    requests.post(export_results_url, cookies=cookies)
    excel_results_url = prepend_base(f"/app/results/{cookies['maedtype']}_results_excel.php")
    return requests.get(excel_results_url, cookies=cookies).content


# Results

def export_results(results_path, results):
    open(format_path(results_path), 'wb').write(results)

def duplicate_template(country, scenario, maed_type, current_filename):
    template_dir_path = format_path(f"./resources/maed-2.0.0/project/storage/{maed_type}/data/projects/admin/")
    new_filename = f"{current_filename}_{country}_{scenario}_{strftime('%Y-%m-%d_%H-%M-%S')}"
    copytree(format_path(f"{template_dir_path}/{current_filename}"), format_path(f"{template_dir_path}/{new_filename}"))
    return new_filename

def execute_maed_scenario(years, results_path, maed_input, maed_type, case_study_name):
    phpsessid = get_php_session_id()
    cookies = get_cookies(phpsessid, maed_type, case_study_name)
    
    print("Importing data to MAED...")
    if (maed_type == "maedd"): edit_maedd_general_info(years, cookies)
    elif (maed_type == "maedel"): edit_maedel_general_info(years, cookies)
    else: raise Exception("Invalid maed type. Choose between: 'maedd', 'maedel'")
    edit_maed_data(maed_input, cookies)

    print("Importing done!")
    # input("Press Enter to continue...")
    print("Calculating and exporting results...")
    calculate(cookies)
    results = download_results(cookies)
    export_results(results_path, results)

def create_copy(file_path, output_filename):
    copied_file_path = f"./runs/{output_filename}"
    os.makedirs(os.path.dirname(copied_file_path), exist_ok=True)
    copy(file_path, copied_file_path)

def choose_maed_input_menu(country, scenario, maed_type, years, selected_option):
    maed_input=None   
    if selected_option == "Transport Starter Kits":
        maed_input=prepare_maed_TSDK.run(country, scenario, years)
        if maed_input is None: maed_input = f"./resources/maed-2.0.0//inputs/{maed_type}_TSDK_template.xlsx"
        current_filename = "TSDK_template"
    elif selected_option == "IEA":
        maed_input = prepare_IEA_maed_input.run(country)
        if maed_input is None: maed_input = f"./resources/maed-2.0.0//inputs/{maed_type}_IEA_template.xlsx"
        current_filename = "IEA_template"

    return maed_input, current_filename

def run(country, maed_type, selected_option, scenario, years):
    print("\n---------- MAED ----------\n")
    config = read_configuration("maed")
    #years = config["years"]     
    path = "./resources/maed-2.0.0"
    maed_input, current_filename = choose_maed_input_menu(country, scenario, maed_type, years, selected_option)
    copied_filename = f"./MAED/{maed_type}/{country}_{scenario}_{strftime('%Y-%m-%d_%H-%M-%S')}"
    create_copy(maed_input, copied_filename+"/maed_input.xlsx")
    results_path = "./resources/maed-2.0.0/maed_results.xlsx"
    print("Starting MAED...")
    logfile = open('./resources/maed-2.0.0/maed_py.log.txt', 'w+')
    logfile.flush()  
    subprocess.Popen(f"server.bat {PORT}", shell=True, cwd=path, stdout=logfile, stderr=logfile)
    case_study_name = duplicate_template(country, scenario, maed_type,current_filename)
    execute_maed_scenario(years, results_path, maed_input, maed_type, case_study_name)
    results_file=create_copy(results_path, copied_filename+"/maed_results.xlsx")
    if pick(["yes", "no"], "Do you want to open the MAED interface?", indicator='=>')[0] == "yes":
        webbrowser.open(prepend_base(f"/app.html#/GeneralInformation/{case_study_name}")) # Comment out to avoid opening the browser
    print("MAED done!")
    return results_path


if __name__ == "__main__":
    country = input("Enter a country: ")
    maed_type, selected_option, scenario, years = maed_config.run(country)
    run(country, maed_type, selected_option, scenario, years)