#scheiße finden rechtsklick kopieren und xpath
#erster montag im monat => im dropdown //*[@id="datetimepicker_timesheet"]/span[2]
#/html/body/div[2]/div/div[1]/table/tbody/tr[2]/td[2]/div
#laden
#=>

import time
import datetime
import calendar
import json
import os
import zipfile
import shutil
import subprocess
import sys

import requests
from urllib3.exceptions import InsecureRequestWarning
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains

from book_ppm_settings import *
from book_ppm_version import version as book_ppm_version
from book_ppm_version import versiontuple

#disable InsecureRequestWarning: Unverified HTTPS request is being made to host 'date.nager.at'. Adding certificate verification is strongly advised. warning
# See: https://urllib3.readthedocs.io/en/latest/user-guide.html#ssl to solve this issue
#urllib3.disable_warnings()
# Suppress only the single warning from urllib3 needed
requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)

#Initialization for holiday API
PROJECT_URL     = 'https://aztech.service-now.com/tcp'
holiday_api     = f'https://date.nager.at/api/v3/publicholidays/'
holidays = []
HOLIDAYSSET = False

gitHubJsonResponse = None

#Debug string for explicit run
testText="""PRO0001998 - 22201165 - ABSi Project Budget	 3,37
PRO0024238 - ABS Core Budget	 159,21
PRO0027515 - 22206001 - Az France	 3,37"""

#testText="""PRO0001998 - 22201165 - ABSi Project Budget	 3,37"""


#column definition of ppm website
colDefinition={
    "Project Number" : 0,
    "Project Name" : 1,
    "Project/Task" :2,
    "Accounting Element" :3,
    "Remaining Hours" :4,
    "Correction" :5,
    "State" :6,
    "sunday" :7,
    "monday" :8,
    "tuesday" :9,
    "wednesday" :10,
    "thursday" :11,
    "friday" :12,
    "saturday" :13,
    "Total" :14,
}

def debugLogger(msg):
    if LOGGER:
        print(msg)

def __setup_driver():
    if DRIVER == 'edge':
        #TODO EdgeOptions?
        return  webdriver.Edge()
    elif DRIVER == 'chrome':
        return webdriver.Firefox()
    elif DRIVER == 'firefox':
        return webdriver.Firefox()
    elif DRIVER == 'safari':
        return webdriver.Safari()

def setup_driver():
    driver = __setup_driver()
    return driver

def __init_holidays(date, absence):
    #init holiday api so we do not need to call it 1000 times 
    global holidays
    global HOLIDAYSSET
    HOLIDAYSSET = True
    holidaysTemp = []
    holiday_api_actual_year = holiday_api
    holiday_api_actual_year += f'{date.year}/{COUNTRY_CODE}'
    r = requests.get(holiday_api_actual_year, verify = False)
    if r.status_code == 200:
        holidaysTemp = r.json()
    else:
        raise Exception(f"Cannot fetch holidays for {COUNTRY_CODE} for year {year}")

    for holiday in holidaysTemp:
        date_format = '%Y-%m-%d'
        date_obj = datetime.datetime.strptime(holiday["date"], date_format)
        
        if date.month == date_obj.month:
            holidays.append(holiday)
    
    for day in absence:
        date_format = '%Y-%m-%d'
        date_obj = datetime.datetime.strptime(day, date_format)
        if date.month == date_obj.month:
            dayTemp = {
                "date": str(date_obj.date()),
                "localName": "Urlaub",
                "name": "Absence",                
            }
            holidays.append(dayTemp)
    #print(json.dumps(holidays, indent=4))
    #{
    #    "date": "2023-08-15",
    #    "localName": "Maria Himmelfahrt",
    #    "name": "Assumption Day",
    #    "countryCode": "AT",
    #    "fixed": true,
    #    "global": true,
    #    "counties": null,
    #    "launchYear": null,
    #    "types": [
    #        "Public"
    #    ]
    #}

def __is_holiday(date):
    if not HOLIDAYSSET:
        raise Exception(f"Holidays not set, but they are already needed.")

    #check if date is a holiday
    for holiday in holidays:
        if str(date.date()) == holiday['date']:
            return True

    return False

def get_first_business_day(year, month):
    d = datetime.datetime(int(year), int(month), 1)
    #greater den 1 and smaler 5 is mon tue wed thu fri   0 + 6 + 7 is sun sat sun
    if d.isoweekday() >= 1 and d.isoweekday() <= 5 and not __is_holiday(d):
        return d
    return get_next_business_day(d)

def get_next_business_day(dt):
    dt += datetime.timedelta(days=1)
    #check if weekand or holiday
    while dt.isoweekday() > 5 or __is_holiday(dt):
        dt += datetime.timedelta(days=1)
    return dt

#Deprecated output_date_to_stdout used
def output_project_to_stdout(output):
    for project in output:
        print(project)
        for date in output[project]:
            date_string = f"{date:%Y-%m-%d %A}"
            print(f"{date_string:<22} {output[project][date]:.2f} hours")
            #Output a newline after a workweek
            if date.isoweekday() == 5:
                print("")
        print()

#Deprecated generate_with_leading_date_name used
def generate_with_leading_project_name(inputText,inputDate,start_day):
    output = {}
    for line in inputText.splitlines():

        inputText = line.split(" ")
        project_name=inputText[0]
        project_duration=float(inputText[len(inputText)-1].replace(",","."))
        output[project_name] = {}

        while True:
            daily_duration = 8.0 if project_duration >= 8 else project_duration
            project_duration -= daily_duration

            if output[project_name].get(start_day) is not None:
                output[project_name][start_day] += daily_duration
            else:
                output[project_name][start_day] = daily_duration

            next_day = get_next_business_day(start_day)
            if int(next_day.month) > int(start_day.month):
                start_day = get_first_business_day(inputDate.year, inputDate.month)
            else:
                start_day = next_day
            if project_duration <= 0:
                break
    return output

#output to stdout
def output_date_to_stdout(output):
    for date in output:
        date_string = f"{date:%Y-%m-%d %A}"
        print(f"{date_string:<22}")
        for project in output[date]:
            if project == 'MAX':
                continue
            print(f"\t{project:<12} {output[date][project]:>5.2f} hours")
        ##Output a newline after a workweek
        if date.isoweekday() == 5:
            print("")

# Leading date was easier then leading project to get it right
def generate_with_leading_date(inputText,inputDate,start_day):
    output = {}

    projects = {}
    projectsIndex = 0

    for holiday in holidays:
        date_format = '%Y-%m-%d'
        date_obj = datetime.datetime.strptime(holiday["date"], date_format)

        output[date_obj] = {}
        output[date_obj]['MAX'] = 7.7
        output[date_obj]['Absence'] = 7.7

    #split input into projectnumber and hours save it in projects
    for line in inputText.splitlines():
        inputText = line.split(" ")
        project_name=inputText[0]
        project_duration=float(inputText[len(inputText)-1].replace(",","."))
        projects[project_name] = project_duration
        projectsIndex += 1

    #get first projectnumber/name
    projectsIndexKey = next(iter(projects))
    #init output
    output[start_day] = {}
    output[start_day]['MAX'] = 0.0
    MAX = 8.0
    while True:
        max_duration = MAX
    
        #substract already "booked" on this date (from other projects)
        max_duration -= output[start_day]['MAX'] 

        daily_duration = max_duration if projects[projectsIndexKey] >= max_duration else projects[projectsIndexKey]
        projects[projectsIndexKey] -= daily_duration

        if output[start_day].get(projectsIndexKey) is not None:
            # when something is already inside
            output[start_day][projectsIndexKey] += daily_duration
        else:
            #this is the first entry
            output[start_day][projectsIndexKey] = daily_duration

        #add the duration to have it for other dates
        output[start_day]['MAX'] += daily_duration

        #when there are no hours left on this project delete it and break if there is no project left else get next project
        if projects[projectsIndexKey] <= 0:
            del projects[projectsIndexKey]
            if len(projects) == 0:
                break 
            projectsIndexKey = next(iter(projects))
        # when the maximum of this date is reached
        if output[start_day]['MAX'] >= MAX:
            #fetch next date
            next_day = get_next_business_day(start_day)
            #when the date is in the next month
            if int(next_day.month) > int(start_day.month):
                #add to maximum possible per date and fetch first date of month
                MAX += 4.0
                if MAX > 24:
                    raise Exception(f"You are not allowed to book more than 24 hours a day!")

                start_day = get_first_business_day(inputDate.year, inputDate.month)
            else:
                #set next day and init it 
                start_day = next_day
                if not output.get(start_day):
                    output[start_day] = {}
                    output[start_day]['MAX'] = 0.0

    return dict(sorted(output.items()))

#Load the page and wait for a specific element in the browser
def load_page_and_wait_till_loaded(driver):
    driver.get(f'{PROJECT_URL}')
    assert 'Time Sheet Portal' in driver.title
    delay = 10 # seconds
    try:
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CLASS_NAME, f"my-timesheet.pull-left.ng-scope")))
        print("Page is ready!")
    except TimeoutException:
        print("Loading took too much time!")

#click on datetimepicker
def browser_calendar_click_on(driver):
    debugLogger("Called browser_calendar_click_on")
    if not driver:
        return
    d_field = driver.find_element(By.XPATH, f"//*[@id=\"datetimepicker_timesheet\"]/span[2]")
    d_field.click()

#click on next week arrow and wait till loaded
def browser_calendar_click_next_week(driver):
    debugLogger("Called browser_calendar_click_next_week")
    if not driver:
        return
    try:
        d_field = driver.find_element(By.XPATH, f"//*[@id=\"x7a8ac6aa0b032200acc30e7363673acb\"]/div/div[2]/button[2]")
    except:
        raise Exception(f"Could not find next week button.")

    d_field.click()
    waitForSelectLoading(driver)

#click on absence button and wait till loaded
def add_absence(driver):
    debugLogger("Called add_absence")
    if not driver:
        return
    try:
        d_field = driver.find_element(By.XPATH, f"//*[@id=\"other\"]/div[5]/div[2]/button[2]")
    except:
        raise Exception(f"Could not find absence button.")

    d_field.click()
    waitForSelectLoading(driver) 

#select year and month from datetimepicker
def browser_calendar_select_calendar_year_and_month(driver,inputYear,inputMonth):
    browser_calendar_click_on(driver)

    datepicker_days_table = driver.find_element(By.CSS_SELECTOR, f"body > div.bootstrap-datetimepicker-widget.dropdown-menu.picker-open.bottom > div > div.datepicker-days > table")
    datepicker_month_table = driver.find_element(By.CSS_SELECTOR, f"body > div.bootstrap-datetimepicker-widget.dropdown-menu.picker-open.bottom > div > div.datepicker-months > table")
    datepicker_year_table = driver.find_element(By.CSS_SELECTOR, f"body > div.bootstrap-datetimepicker-widget.dropdown-menu.picker-open.bottom > div > div.datepicker-years > table")
    
    datepicker_days_table.find_element(By.TAG_NAME, f"thead").find_element(By.CLASS_NAME, f"picker-switch").click()
    waitForSelectLoading(driver) 
    datepicker_month_table.find_element(By.TAG_NAME, f"thead").find_element(By.CLASS_NAME, f"picker-switch").click()
    waitForSelectLoading(driver) 

    datepicker_year_table_body = datepicker_year_table.find_element(By.TAG_NAME, f"tbody")
    for row in datepicker_year_table_body.find_elements(By.TAG_NAME, "tr"):
        for span in row.find_elements(By.TAG_NAME, "td")[0].find_elements(By.TAG_NAME, "span"):
            if span.text == str(inputYear):
                span.click()
                break

    datepicker_month_table_body = datepicker_month_table.find_element(By.TAG_NAME, f"tbody")
    for row in datepicker_month_table_body.find_elements(By.TAG_NAME, "tr"):
        for span in row.find_elements(By.TAG_NAME, "td")[0].find_elements(By.TAG_NAME, "span"):
            if span.text == str(calendar.month_abbr[inputMonth]):
                span.click()
                break

#select a specific day from datetimepicker
def browser_calendar_select_day(driver,selectDate):
    debugLogger("Called browser_calendar_select_day")
    if not driver:
        return
    browser_calendar_click_on(driver)

    datepicker_days_table = driver.find_element(By.CSS_SELECTOR, f"body > div.bootstrap-datetimepicker-widget.dropdown-menu.picker-open.bottom > div > div.datepicker-days > table")
    datepicker_days_table_body = datepicker_days_table.find_element(By.TAG_NAME, f"tbody")

    found = False
    #for day in datepicker_days_table_body.find_elements(By.XPATH, f"//td[(@class, 'day') and not (@class='old') and not (@class='disabled')]"):
    for day in datepicker_days_table_body.find_elements(By.XPATH, f"//td[@class='day']"):
        if day.text == str(selectDate.day):
            found = True
            day.click()
            break

    if not found:
        raise Exception(f"Day of Browser {str(selectDate)} is disabled.")

    waitForSelectLoading(driver) 

#wait till loading screen is gone
def waitForSelectLoading(driver):
    delay = 20 # seconds
    try:
        print("Wait for Loading!")
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, f"body > div.sp-page-root.page.flex-column.sp-can-animate > div.sp-page-loader-mobile.visible-xs.visible-sm.sp-loading-indicator.la-sm.invisible")))
        print("Loading is finished!")
    except TimeoutException:
        print("Loading took too much time!")

#COLUMNS USED IN BROWSER SITE
#Project Number
#Project Name
#Project/Task
#Accounting Element
#Remaining Hours
#Correction
#State
#Sun 13   7
#Mon 14
#Tue 15
#Wed 16
#Thu 17
#Fri 18
#Sat 19  13
#Total

def DoubleOrSingleClick(driver, dictDateHours, cols):
    doubleClicked = False
    for date in dictDateHours:
        if not doubleClicked:
            doubleClicked = True
            debugLogger("double click perform "+ str(date))
            ACTION = ActionChains(driver)
            ACTION.double_click(cols[colDefinition[date.strftime('%A').lower()]]).perform()
            ACTION.reset_actions()
        else:
            debugLogger("normal click perform "+ str(date))
            cols[colDefinition[date.strftime('%A').lower()]].click()
        input_field = cols[colDefinition[date.strftime('%A').lower()]].find_element(By.XPATH, f"//input[@id=\"{date.strftime('%A').lower()}\"]")
        input_field.clear()
        input_field.send_keys(str(dictDateHours[date]))
        input_field.send_keys()

#for every browser row get the columns and look if it is the right project.
#when yes double click on first given date and set hours
# go through whole list to set all the hours for specifc date on this project
def TableClickThroughProject(driver, project_name, dictDateHours):
    debugLogger("Called TableClickThroughProject")
    debugLogger(dictDateHours)
    if not driver:
        return

    browser_table = driver.find_element(By.ID, f"tc-grid-table")
    browser_table_body = browser_table.find_element(By.TAG_NAME, f"tbody")
    timesheet_pull_left = browser_table_body.find_element(By.XPATH, f"//div[@class='my-timesheet pull-left ng-scope']")

    rows=driver.find_elements(By.XPATH, f"//*[@id=\"tc-grid-table\"]/tbody/tr")
    for row in rows :
        cols= row.find_elements(By.TAG_NAME, f"td")
        if project_name == "Absence":
            if cols[colDefinition['Project/Task']].text == project_name:
                debugLogger("ABSENCE!!!!")
                DoubleOrSingleClick(driver, dictDateHours, cols)
        else:
            if cols[colDefinition['Project Number']].text == project_name:
                debugLogger("normal Date")
                DoubleOrSingleClick(driver, dictDateHours, cols)

def isInSameWeek(lastDate, date):
    if not lastDate:
        return True
    return lastDate.isocalendar()[1] == date.isocalendar()[1]

#convert leading date to project so we can click from left to right on specific project 
def convertLeadingDateToProjectOnTempDict(driver, output, useUI):
    switchWeek = False
    add_absence(driver)
    test = {}
    lastDate = None
    debugLogger(output)
    debugLogger("LOOP START")
    for date in output:
        debugLogger("CHeck if last date")
        if lastDate:
            calculatedWeeks = ( date.isocalendar()[1] - lastDate.isocalendar()[1] )
            if calculatedWeeks > 1:
                for project in test:
                    TableClickThroughProject(driver, project ,test[project])
                browser_calendar_select_day(driver, date)
                add_absence(driver)
                test = {}
            elif calculatedWeeks == 1 or date.isoweekday() == 5:
                for project in test:
                    TableClickThroughProject(driver, project ,test[project])
                browser_calendar_click_next_week(driver)
                add_absence(driver)
                test = {}
            elif calculatedWeeks == 0:
                pass
            else:
                raise Exception("WTF happend here with the isocalendar weeks. Could not find calculated weekdays")
                
        #if switchWeek:
        #    switchWeek = False
        #    browser_calendar_click_next_week(driver)
        #    add_absence(driver)
        #    test = {}
        for project in output[date]:
            if project == 'MAX':
                continue
            if test.get(project) == None:
                test[project] = {}
            test[project][date] = output[date][project]
        #if date.isoweekday() == 5:
        #    switchWeek = True
        #    for project in test:
        #        TableClickThroughProject(driver, project ,test[project])
        lastDate = date

    debugLogger("LOOP FINISHED")
    #after output is gone submit last elements in test
    for project in test:
        TableClickThroughProject(driver, project ,test[project])

#Setup browser stuff only if useUI is TRUE
def main(inputMonth=8,inputYear=2023,inputText=testText,useUI=False,absence=[]):

    #date_format = '%Y-%m-%d'
    #date_obj = datetime.datetime.strptime(holiday["date"], date_format)
    debugLogger("LOOP START")
    if DEBUG:
        absence=["2023-08-03","2023-08-11"]

    inputDate=datetime.datetime(int(inputYear), int(inputMonth), 1)

    __init_holidays(inputDate, absence)  

    start_day = get_first_business_day(inputDate.year, inputDate.month)

    if True == useUI:
        driver = setup_driver()
        load_page_and_wait_till_loaded(driver)
        browser_calendar_select_calendar_year_and_month(driver,inputYear,inputMonth)
        browser_calendar_select_day(driver, start_day)

    output = generate_with_leading_date(inputText,inputDate,start_day)

    if True == useUI:
        convertLeadingDateToProjectOnTempDict(driver,output,useUI)
        driver.close()

    return output

def getGitHubResponse():
    global gitHubJsonResponse
    r = requests.get("https://api.github.com/repos/x3n0r/book_ppm/releases/latest", verify=False)
    gitHubJsonResponse = r.json()
    debugLogger(gitHubJsonResponse)

def VersionUpdateNeeded():
    if not gitHubJsonResponse:
        getGitHubResponse()
    debugLogger(gitHubJsonResponse["name"] + " > " + book_ppm_version)
    return versiontuple(gitHubJsonResponse["name"]) > versiontuple(book_ppm_version)

def getGitHubRelease():
    if not gitHubJsonResponse:
        getGitHubResponse()
    return gitHubJsonResponse["name"]

def getGitHubBody():
    if not gitHubJsonResponse:
        getGitHubResponse()
    return gitHubJsonResponse["body"]

def getGitHubDownload():
    if not gitHubJsonResponse:
        getGitHubResponse()
    print(gitHubJsonResponse["zipball_url"])
    actualDirectory = os.getcwd()
    downloadDirectory = actualDirectory + "\\" + "downloads"
    if os.path.isdir(downloadDirectory):
        shutil.rmtree(downloadDirectory)
    downloadFullPath = downloadDirectory + "\\" + "book_ppm_" + gitHubJsonResponse["name"] + ".zip"
    if not os.path.isdir(downloadDirectory):
        os.makedirs(downloadDirectory)
        debugLogger("created folder : " + downloadDirectory)
    else:
        debugLogger("folder already exists: " + downloadDirectory)
    r = requests.get(gitHubJsonResponse["zipball_url"], allow_redirects=True, verify=False)
    open(downloadFullPath, 'wb').write(r.content)
    shutil.unpack_archive(downloadFullPath, downloadDirectory)
    createdFolder = downloadDirectory + "\\" + next(os.walk(downloadDirectory))[1][0]
    destinationFolder = actualDirectory
    if TESTUPDATE:
        destinationFolder = downloadDirectory 
    shutil.copytree(createdFolder, destinationFolder, dirs_exist_ok=True)
    print("open new folder")
    openApp=destinationFolder + "\\" + "book_ppm_gui.pyw"
    print(openApp)
    subprocess.Popen([sys.executable, openApp])

if __name__ == '__main__':
    if DEBUG:
        output = main()
        if output:
            print(output)
