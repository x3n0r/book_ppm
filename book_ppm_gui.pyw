# Install something
# pip install --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org PySimpleGUI
# PySimpleGUI, requests, urllib3, selenium
import PySimpleGUI as sg

import calendar
from calendar import Calendar, monthrange
from datetime import datetime
import os.path
import os
from book_ppm_version import version as book_ppm_version
from book_ppm import main
from book_ppm import VersionUpdateNeeded, getGitHubRelease, getGitHubBody, getGitHubDownload
from book_ppm import output_date_to_stdout
from book_ppm import testText
from book_ppm_settings import *
import json

ALL_FONTS = sg.Text.fonts_installed_list()

# Size for function name (spacers)
NAME_SIZE           = 15

settings = None

# function which return a simplegui text where Spacers are set
def name(name, size=NAME_SIZE):
    dots = size-len(name)-2
    return sg.Text(name + ' ' + '•'*dots, size=(size,1), justification='r',pad=(0,0), font=(settings.get('FONT', DEFAULT_FONT),DEFAULT_FONTSIZE))

# create the main window
def make_main_window(theme=None):

    months = list(calendar.month_name)
    #First entry is empty remove that
    months.pop(0)
    year = datetime.now().year
    years = list()

    #year + and - 1 for range
    for x in range(year-1, year+2):
        years.append(str(x))

    layout = [
        [ sg.MenubarCustom([['File', ['Settings', 'Update','Exit']]],  k='-CUST MENUBAR-',p=0)],
        [ sg.Sizer(50,25)],
        [ name('Year') , sg.Spin(years,initial_value=str(year), s=(15,2), readonly=True,key='-inputYear-') ],
        [ name('Month'), sg.Spin(months,initial_value=calendar.month_name[datetime.now().month], s=(15,2), readonly=True,key='-inputMonth-') ],
        [ sg.Multiline(s=(50,5),default_text='',key='-inputText-')],
        [ sg.HSep()],
        [ sg.Checkbox('UseBrowserUi',default=(not DEBUG),key='-useBrowserUI-')],
        [ sg.Button('Submit',key="Submit") ],
        [ sg.Sizer(50,25)],
    ]

    window = sg.Window('PPM Booker', layout, keep_on_top=True, use_custom_titlebar=sg.MenubarCustom,resizable=True,alpha_channel=DEFAULT_TRANSPERENCY)

    return window

# window for output when brosergui is not used
def make_output_window(output):

    output_date_to_stdout(output)
    treedata = sg.TreeData()

    weekIndex=0
    for date in output:
        date_string = f"{date:%Y-%m-%d %A}"
        treedata.Insert('',date_string,date_string,[])
        for project in output[date]:
            if project == 'MAX':
                continue
            treedata.Insert(date_string,date_string+project,project,f"{output[date][project]:>5.2f}",[])
        ##Output a newline after a workweek
        if date.isoweekday() == 5:
            weekIndex+=1
            treedata.Insert('','weekend'+str(weekIndex),'',[])

    layout = [
        [sg.Tree(treedata, headings=['Hours'], num_rows=30,col0_width=30, key="-TREE-")],
    ] 
    window = sg.Window('Book PPM Output', layout, keep_on_top=True, use_custom_titlebar=sg.MenubarCustom,resizable=True,finalize=True )

 
    return window

#create settings window
def make_settings_window():

    savedTheme = sg.theme()
    savedFont = settings.get('FONT', DEFAULT_FONT)

    frame_l = [
        [sg.Text('List of InBuilt Themes')],
        [sg.Text('Please Choose a Theme  to see Demo window')],
        [sg.Listbox(values = sg.theme_list(),
            size =(30, 20),
            key ='-LISTTHEME-',
            enable_events = True)
        ],
    ]

    layout_l = [
        [ sg.Frame('Theme', frame_l) ],
    ]

    frame_r = [
        [sg.Text("Please Choose a Font to see as 'Demo Text'")],
        [sg.Text('Demo Text',
            size=(40, 2),
            relief=sg.RELIEF_GROOVE,
            font=(settings.get('FONT', DEFAULT_FONT),DEFAULT_FONTSIZE),
            key='-text-',
        )],
        [sg.Listbox(ALL_FONTS, size=(30, 20), change_submits=True, key='-LISTFONT-')],
    ]

    layout_r = [
        [ sg.Frame('Font', frame_r) ],
    ]

    layout_button = [
        [sg.Button('Save')],
        [sg.Button('Exit')],
    ]

    layout = [
        [sg.Col(layout_l, p=0), sg.Col(layout_r, p=0)],
        [layout_button]
    ]

    useTheme = savedTheme
    useFont = savedFont

    window = sg.Window('Theme List', layout,keep_on_top=True)
    while True:  
        event, values = window.read()
        
        if event in (None, 'Exit'):
            #reuse theme and font which were saved during init
            useTheme = savedTheme
            useFont = savedFont
            break
        if event == "-LISTTHEME-":
            #set theme ( is set for everything all other windows as well)
            sg.theme(values['-LISTTHEME-'][0])
            #open popup with new theme
            popup_return = sg.popup_get_text('This is {}'.format(values['-LISTTHEME-'][0]),keep_on_top=True)
            #when clicked on ok then it is not None
            if popup_return is not None:
                #set theme close window and return new theme
                sg.theme(savedTheme)
                window.close()
                return [values['-LISTTHEME-'][0], useFont]
        if event == "-LISTFONT-":
            #get textelement -text- and update with new font
            text_elem = window['-text-']
            text_elem.update(font=(values['-LISTFONT-'][0], DEFAULT_FONTSIZE))
            useFont = values['-LISTFONT-'][0]
        if event == "Save":
            break
        
    #revert theme (if not done last window freeze)
    sg.theme(savedTheme)
    window.close()
    return [sg.theme(), useFont]

# window for output when brosergui is not used
def make_absence_window(inputMonth=8,inputYear=2023):
    inputDate=datetime(int(inputYear), int(inputMonth), 1)

    checkBoxArray1=[]
    checkBoxArray2=[]
    checkBoxArrayActual=checkBoxArray1
    count=0
    for d in [x for x in Calendar().itermonthdates(inputYear, inputMonth) if x.month == inputMonth]:
        if d.isoweekday() > 5:
            continue
        if count >= 15:
            checkBoxArrayActual = checkBoxArray2
        print(d)
        checkBoxArrayActual.append([sg.Checkbox(d,default=False,key='-'+str(d)+'-')])
        count+=1

    #for d in [x for x in Calendar().itermonthdates(inputYear, inputMonth) if x.month == 7]:

    layoutL = checkBoxArray1
    layoutR = checkBoxArray2

    layout = [
        [sg.Col(layoutL, p=0), sg.Col(layoutR, p=0)],
        [sg.Button('Ok')]
    ]

    window = sg.Window('Absence Days', layout, keep_on_top=True, use_custom_titlebar=sg.MenubarCustom,resizable=False,finalize=True,size=(220,550) )

    return window

def make_update_window():
    layoutL = [
        [ name('Actual Release', size=20) , sg.Text(book_ppm_version) ],
        [ name('GitHub Release', size=20) , sg.Text(getGitHubRelease()) ],
        [ sg.Button('Update', disabled = not VersionUpdateNeeded())  ],
        [ sg.Button('Ok') ],
    ]

    layoutR = [
        [ sg.Text("Latest/Newest Changelog:") ],
        [ sg.Multiline(s=(50,5),default_text=getGitHubBody(), disabled=True, size=(100,100))],
    ]

    layout = [
        [sg.Col(layoutL, p=0), sg.Col(layoutR, p=0)],
    ]

    window = sg.Window('Auto Updater', layout, keep_on_top=True, use_custom_titlebar=sg.MenubarCustom,resizable=False,finalize=True, size=(800,300) )

    return window

def main():
    global settings
    #simple gui settings saved to %LocalAppData%\PySimpleGUI with filename saved in SETTINGS_FILE
    settings = sg.UserSettings(filename=SETTINGS_FILE)
    settings.load()
    #set default theme when no settings saved
    sg.theme(settings.get('THEME', DEFAULT_THEME))

    window = make_main_window()

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Settings':
            #when clicked on settings menu entry
            (theme, font) = make_settings_window()
            if theme != sg.theme() or font != settings.get('FONT', DEFAULT_FONT):
                window.close()
                sg.theme(theme)
                settings['THEME'] = theme
                settings['FONT'] = font
                settings.save()
                window = make_main_window()

        if event == 'Update':
            update_window = make_update_window()
            while True:
                event, values = update_window.read()
                if event == sg.WIN_CLOSED or event == 'Ok':
                    update_window.close()
                    break
                if event == 'Update':
                    getGitHubDownload()
                    update_window.close()
                    window.close()
                    break

        if event  == 'Submit': 
            year = int(values['-inputYear-'])
            month = datetime.strptime(values['-inputMonth-'], '%B').month 
            text = values['-inputText-']
            useBrowserUI = values['-useBrowserUI-']

            if text == "" and not DEBUG:
                continue

            answer = sg.popup_yes_no("Do you have any absence day to book(except holidays)?",  title="YesNo", keep_on_top=True)
            #print ("You clicked", answer)
            clickedAbsence=[]
            winClosed = False
            if answer == 'Yes':
                absence_window = make_absence_window(inputYear=year,inputMonth=month)
                while True:
                    event, values = absence_window.read()
                    if event == sg.WIN_CLOSED or event == 'Exit':
                        winClosed = True
                        break
                    if event == "Ok":
                        #print("GetAll Clicked Items")

                        for d in [x for x in Calendar().itermonthdates(year, month) if x.month == month]:
                            if d.isoweekday() > 5:
                                continue    
                            value = values['-'+str(d)+'-']
                            if value:
                                clickedAbsence.append(str(d))
                        break
            if winClosed:
                continue

            #run book_ppm program
            if text != "" or DEBUG:
                if DEBUG:
                    text=testText
                output = main(inputYear=year,inputMonth=month,inputText=text,useUI=useBrowserUI,absence=clickedAbsence)
                make_output_window(output)

    window.close()

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        sg.popup_error_with_traceback('Error happened. Here is some info:', e)
