import json
from pathlib import Path
import subprocess
import xlwings as xw


# Paths
def format_path(*path):
    return str(Path(*path))

# Processes
def execute_program(dir_name, file_name, args):
    path = format_path(dir_name, file_name)
    proc = subprocess.run(path + " " + args, shell=False, capture_output=True)
    logs = proc.stdout + b" " + proc.stderr
    return proc.returncode == 0, logs

# Bytes
def stringify_bytes(string):
    try:
        string = string.decode()
    except (UnicodeDecodeError, AttributeError):
        pass
    return string

# Files
def write_to_file(dir_name, file_name, lines):
    path = format_path(dir_name, file_name)
    with open(path, "w") as file:
        file.write(stringify_bytes("\n".join(lines)))

# JSON
def read_configuration(tool):
    with open(format_path('settings.json')) as config_file:
        return json.load(config_file)[tool]


# Excel
orientations = {
    "row": {"expand": "right"},
    "column": {"expand": "down", "transpose": True},
}


def open_work_book(path):
    return xw.Book(format_path(path))

def close_work_book(book):
    book.close()

def run_excel_macro(wb, macro_name):
    execute_macro = wb.macro(macro_name)
    try: execute_macro()
    except: pass


def get_sheet(wb, sheet_name):
    return wb.sheets[sheet_name]


def set_cell_values(sheet, cell_coords, setting_value, direction):
    selected_cells = sheet[cell_coords]
    selected_cells.options(**orientations[direction]).value = setting_value


def set_checkbox_value(sheet, checkbox_name, checkbox_value):
    selected_checkbox = sheet.api.OLEObjects(checkbox_name)
    selected_checkbox.Object.Value = checkbox_value
