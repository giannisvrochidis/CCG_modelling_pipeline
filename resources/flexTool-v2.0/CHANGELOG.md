# Changelog
All **notable** changes to this project are documented here.

## [2.0] - 2020-03-06
Improved support for different Excel language versions and regional settings
Bug fix: several small bugs in outputs and model crashes due to special circumstances

## [2.0.alpha] - 2020-01-30
Bug fix: Input files miscalculate time series and lets storage to generate for free at the last time step. template.xlsm has been fixed and instructions for other files here: https://gitlab.vtt.fi/FlexTool/FlexTool/issues/48
Bug fix: Empty spaces in unit_type names will now work.
Name change: master.xlsm is now flexTool.xlsm

## [1.3] - 2020-01-12
VRE units are now real units (they are now individual units that can be dispatched within availability limits)
Minimum generation, maximum generation and minimum online limits are available as time series using ts_unit.

## [1.2] - 2019-11-20
Minimum uptimes and downtimes
Made colors work when sliding the time in the results shhet
Reserves for conversion units
Demand_increase and conversion units included in ramp constraints
Improved master.xlsx behaviour during model execution (less windows opening and closing)

## [1.1] - 2019-03-08
Implementation of nodeGroups and unitGroups 
Improving parallel calculations

## [1.0] - 2018-12-14
Below is the old changelog from the master.xlsx. Flextool.mod has not had any changelog before this.

v 0.64
Scenario activation bug fixed

V 0.63
Bug related to long scenario names fixed

v 0.62
Open result excel in full screen

v 0.61
Changed the warning when all parallel runs are not ready
Added category names for the duration curve plots

v 0.60
"export time series" checks last row from A-column
Added unit (s) to the output of solve time

v 0.59
import results bug fixed

v 0.58
bug in master rename fixed
"Parameter not found" warning message updated

v 0.57
Updated filenames and scenarionames
Updated 'sensitivity definitions' cell colors to match input data files

v 0.56
Improved cell comments in the 'sensitivity definitions'
Updated 'sensitivity definitions' cell colors to match input data files

v. 0.55
Second node added to transfers_invest_plot
Small text changes

v. 0.54
Dropdown list added to scenario defitions first column
Add row buttons to scenario definitions sheet
minimized comments fixed, autosize when workbook is opened
"Settings and filters" sheet protected (no password)
"Sensitivity scenarios" sheet protected (no password)
Name of this excel file can be changed

v 0.53
Added error handler when opening input file
bug fixed in onlineunit plot
case sensitivity bug fixed in result file header
Headers added to plots

v 0.52
Hyperlinks to explanation excel
sheetsform updated to work better with small amount of sheets

v 0.51
Error message added. At least one scenario and input file have to be active. 

v 0.50
Updated parameter names and comments on 'sensitivity definitions' sheet
Updated colors on each sheet

v 0.47
"No inflow" issue fixed
Sheetsform updated
Write parameters module updated
Duration plot y scale modified
Warning if the sensitivity definition does not find anything to replace or if parameter is not found 
Plot start and length parameters added
Input file name removed from results if only one input file
input file and scenario name written in summary sheets


v 0.46
Correcting comments and updating visual outlook

v 0.45
convert_sol modified to use less memory

v 0.44
combine scenarios with different grids and nodes
possible to use comma as decimal delimiter in excel
Sheet selection window updated

v 0.43
Clean-up

v 0.42
Data and settings combined

v 0.41
Code cleaning

v 0.40
"Export time series" button removed from data sheet
"overwrite ts" checkbox option removed
"write ts and run" button added
"Batch file created" message removed

v. 0.39
sensitivity and input file scenarios combined

v 0.38
folders changed
export single time series option removed

v 0.37
Text aligment in merged cells changed to up and left
Match rows in summary sheets between scenarios
code cleaning

v 0.36
Result import start after last case if "number of parallel calculations" < "number of scenarios"
sensitivity definitions updated
Ramp_xh_grid_plot moved to flexibility tab in sheets selection window (sheetsform updated)

v 0.35
Missing parameter bug fixed in sensitivity scenarios

v 0.34
Old "Input file scenarios" log files will be removed before optimisation
Possibility to open log file if optimisation failed

v. 0.33
cost_unittype plot updated

v. 0.32
improved error and optimisation ready handling
ImportRes.vbs updated

v. 0.31
colombia input data + sensitivity definitions in master.xlsx updated

v. 0.30
onlineUnit_plot works with different number of columns

v. 0.29
stacked area chart works now with different number of columns

Change log:
v 0.28
duration & durationRamp plot y limits updated
Hardcoded column names replaced by column numbers in create_duration_plot function

v. 0.27
added parameters loss_of_reserves_penalty and lack_of_capacity_penalty to sensitivity definitions
updated formatting in Sensitivity definitions and sensitivity scenarios
updated information at sensitivity scenarios

v. 0.26
output.txt name bug fixed

v. 0.25
loss_of_load.csv renamed to events.csv
Last row of solver output will be copied to summary

v. 0.24
unit list bug fixed
"_I" or "_D" file selection updated
Sheet order updated to work with two summary files
"create_plot" data availability check updated
Try to import rampRoom files only if available

v. 0.23
Try to add plot only if corresponding data sheet exists
Results file list updated

v. 0.22
Delete *.txt before optimisation
Sheet ordering in results excel
Input file scenarios working again 
Code cleaning

v. 0.21
bug in batch file fixed
printResults bug fixed
error handling added to remove_empty_cols function
stacked area and column charts will be created only if data is available
Sheet selection form is resizable, but form content is not
Result excel resize code changed

v. 0.20
variable datatypes changed in create_plot module

v. 0.19
Result excel resize added again
output txt names corrected
summary results updated
no onlineunit plot if invest only

v. 0.18
minor Colombia case updates

v. 0.17
Result excel resize removed

v. 0.16
sheetsform update: jump to next listbox with arrow keys (not working perfecly)
sheetsform default location to right side
sheetsform resize based on screen resolution
Output txt filenames changed

v. 0.15
"Cancel" button added to "File not found" error message
"Retry-cancel" added if parallel calculation and result file not found

v. 0.14
Some arbitrary limits removed from code
IsInArray function updated
readGridNode function updated

v. 0.13
time series export bug fixed

v. 0.12
Negative grid and node filters
Sheet selection window file updated (SheetsForm.frm)
time series plot x-limits updated

v. 0.11
Getting started updated
storageContent import and plot added
onlineUnit import and plot added
time series plot x-limits updated
units_invest and transfers_invest from _I files when available
import summary files (_D and _I if available)
Sheet selection window file updated (SheetsForm.frm)
same y-axis: node_plot: reserve_requirement_MW  reserve_conventional_MW ja reserve_VRE_MW

v. 0.10
Sheet selection window removed from master.xlsm
remove empty columns from results, not tested with multiple scenarios
