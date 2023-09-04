# book_ppm
Booking time for ppm

## requirements
### windows
- installed python on windows
### python modules
These modules can be installed via Windows Powershell and the command `pip install --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org <name>`
- requests
- selenium
- urllib3
- PySimpleGUI
- 
## How to
- open book_ppm_gui.pyw
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

