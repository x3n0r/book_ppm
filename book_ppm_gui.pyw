# Install something
# pip install --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org PySimpleGUI
# PySimpleGUI, requests, urllib3, selenium
import PySimpleGUI as sg

import calendar
from datetime import datetime
import os.path
import os
from book_ppm import main
from book_ppm import output_date_to_stdout
from book_ppm import testText
from book_ppm_settings import *
import json

ALL_FONTS = sg.Text.fonts_installed_list()

# Size for function name (spacers)
NAME_SIZE           = 15
# Default font and fontsize when there are no settings
DEFAULT_FONT        = 'Courier'
DEFAULT_FONTSIZE    = 10

# function which return a simplegui text where Spacers are set
def name(name):
    dots = NAME_SIZE-len(name)-2
    return sg.Text(name + ' ' + '•'*dots, size=(NAME_SIZE,1), justification='r',pad=(0,0), font=(settings.get('FONT', DEFAULT_FONT),DEFAULT_FONTSIZE))

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
        [ sg.MenubarCustom([['File', ['Settings','Exit']]],  k='-CUST MENUBAR-',p=0)],
        [ sg.Sizer(50,25)],
        [ name('Year') , sg.Spin(years,initial_value=str(year), s=(15,2), readonly=True,key='-inputYear-') ],
        [ name('Month'), sg.Spin(months,initial_value=calendar.month_name[datetime.now().month], s=(15,2), readonly=True,key='-inputMonth-') ],
        [ sg.Multiline(s=(50,5),default_text='',key='-inputText-')],
        [ sg.HSep()],
        [ sg.Checkbox('UseBrowserUi',default=(not DEBUG),key='-useBrowserUI-')],
        [ sg.Button('Submit',key="Submit") ],
        [ sg.Sizer(50,25)],
    ]

    window = sg.Window('The PySimpleGUI Element List', layout, keep_on_top=True, use_custom_titlebar=sg.MenubarCustom,resizable=True,alpha_channel=DEFAULT_TRANSPERENCY)

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

    if event  == 'Submit':
        #run book_ppm program
        year = int(values['-inputYear-'])
        month = datetime.strptime(values['-inputMonth-'], '%B').month 
        text = values['-inputText-']
        useBrowserUI = values['-useBrowserUI-']
        if text != "" or DEBUG:
            if DEBUG:
                text=testText
            output = main(inputYear=year,inputMonth=month,inputText=text,useUI=useBrowserUI)
            make_output_window(output)

window.close()