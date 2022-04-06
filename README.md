# Libre Expenses Calculator

A libre document to calculate the normalized home expenses using python macros on `LibreOffice >= 7.2`

## Usage

1. The macros were created as user scripts. To be able to run them, the custom_scripts.py needs to be placed into LibreOffice **user scripts folder** *if it does not exist, create it*.  

   User scripts folder path is as follows:
    - **Linux**:
    `/home/user/.config/libreoffice/4/user/Scripts/python/` 

    - **Windows**:
    `%APPDATA%\LibreOffice\4\user\Scripts\python`

    - **macOS**:
    `/Users/user/Library/Application Support/LibreOffice/4/user/Scripts/python`

2. Open `expenses.ods`, with macros enabled
3. Click `Recalculate monthlies` button.  

*Read [this help page](https://help.libreoffice.org/latest/en-US/text/sbasic/python/python_locations.html) to learn more about where Python scripts are located.*  
*This repo is inspired by [LibreOffice Conference 2021](https://github.com/rafaelhlima/LibOCon_2021_SFCalc)*

