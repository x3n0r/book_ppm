# book_ppm
Booking time for ppm

## requirements
### windows
- install edge webdriver https://developer.microsoft.com/de-de/microsoft-edge/tools/webdriver/ ( take stable-version ) and place it under C:\Users\a3270\AppData\Local\Microsoft\WindowsApps with the name MicrosoftWebDriver.exe
- installed python on windows
### python modules
These modules can be installed via Windows Powershell and the command `pip install --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org <name>`
- requests
- selenium
- urllib3
- PySimpleGUI

## How to
- open book_ppm_gui.pyw  
![Screenshot1](https://github.com/x3n0r/book_ppm/assets/33955757/06dbd3c0-6792-45eb-8681-53a0f05abeba)
- Select Year
- Select Month
- Input content from Excel list e.g.  
PRO0001998 - 22201165 - ABSi Project Budget	 3,37   
PRO0024238 - ABS Core Budget	 59,21   
PRO0027515 - 22206001 - Az France	 3,37  
- Decide if you want to get the entries directly written within the browser (checkbox UseBrowserUi)
    if the checkbox is selected Browser will open and input the entries
    if it is NOT selected an Outputwindow will open where you can see the output.
- Select Submit
- Answer question if you have any Absence days(without holidays)
- When you click yes a new Window will appear where you can select all the days where you want to input an absence


## Settings
Settings can be deleted by deleting the file `%LocalAppData%\PySimpleGUI\settings\settings.json`
