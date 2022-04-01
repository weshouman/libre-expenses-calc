# Libre Expenses Calculator

A libre document to calculate the normalized home expenses using python macros on LibreOffice >= 7.2

## Usage

The macros were created as user scripts. To be able to run them, the custom_scripts.py needs to be placed into the LibreOffice user scripts folder in your machine.  

- On **Linux** machines, the folder is:
  `/home/user/.config/libreoffice/4/user/Scripts/python/` (if it does not exist, create it)

- On **Windows** machines, the folder is:
  `%APPDATA%\LibreOffice\4\user\Scripts\python`

- On **macOS** machines, the folder is:
  `/Users/user/Library/Application Support/LibreOffice/4/user/Scripts/python`

Read [this help page](https://help.libreoffice.org/latest/en-US/text/sbasic/python/python_locations.html) to learn more about where Python scripts are located.  

Open `expenses.ods`, with macros enabled, and click the `Recaculate monthlies` button.  


*This repo is inspired by [LibreOffice Conference 2021](https://github.com/rafaelhlima/LibOCon_2021_SFCalc)*

