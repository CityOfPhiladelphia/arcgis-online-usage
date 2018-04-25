"""
Script: AGO Usage Script
Version: 3.0.0
Created: 4/15/2017
Created By: Tim Haynes
Updated: 4/24/2018
Updated By: Tim Haynes

Summary: Script pulls basic information about, including the number of views (requests) over the previous 60 days, for
each item in the PHL ArcGIS Online Organization. Information is then filtered and formatted into an easy to use report.
"""

# region Import Libraries
from datetime import date, datetime, timedelta
import csv
import logging
import requests
import sys
import time
import traceback
import win32com.client as win32
from configparser import ConfigParser
import os
# endregion

# region Functions & Classes


# Simple class for attributing rows in the Department CSV
class Departments:
    def __init__(self, row):
        self.user = row[0].strip()
        self.department = row[1].strip()


# Read department list csv / populate dictionary of users and departments / end script on bad input
def read_departmentlist():
    f = open(departmentCSV, 'r')
    i = 0
    try:
        reader = csv.reader(f)
        for row in reader:
            if i == 0:
                i += 1
                continue
            r = Departments(row)
            if r.department in listDepartments:
                dictDepartment[r.user] = r.department
            else:
                print('Invalid department warning for user: ' + r.user)
                print('    - ' + r.department + ' does not exist in the "listDepartments" variable')
                log.warning('Invalid department warning for user: ' + r.user + ' = ' + r.department)
                sys.exit(1)
            i += 1
        print('Dictionary of content owners {User: Department}:')
        print(dictDepartment)
    except IOError:
        print('Error opening ' + departmentCSV, sys.exc_info()[0])
        errorhandler('The script failed in the read_departmentlist function')
        sys.exit(1)
    f.close()


# Request token from ArcGIS Online
def tokengenerator():
    try:
        tokenrequest = requests.post(urlOrg + '/sharing/rest/generateToken?', data={'username': user, 'password': password, 'referer': 'https://www.arcgis.com', 'f': 'json', 'expiration': 20160})
        return tokenrequest.json()['token']
    except:
        errorhandler('The script failed in the tokengenerator function')
        sys.exit(1)


# Create a dictionary of all AGO Users that have logged in {user: last login}
def listagousers():
    try:
        r = requests.get('{0}/sharing/rest/portals/self/users?start=1&num=10&sortField=fullname&sortOrder=asc&f=json&token={1}'.format(urlOrg, token))
        numusers = r.json()['total']
        if numusers % 100 > 0:
            _range = (round(numusers / 100)) + 1
        else:
            _range = (round(numusers / 100))
        start = 1
        for iterUsers in range(_range):
            r = requests.get('{0}/sharing/rest/portals/self/users?start={1}&num=100&sortField=fullname&sortOrder=asc&f=json&token={2}'.format(urlOrg, start, token))
            jsonusers = r.json()['users']
            start += 100
            for user in jsonusers:
                if user['lastLogin'] != -1:
                    dictAGOUsers[user['username']] = time.strftime('%m/%d/%Y', time.localtime(int(user['lastLogin'])/1000))
        print('List of all {0} AGO Users in the org:'.format(len(dictAGOUsers)))
        print(dictAGOUsers)
    except:
        errorhandler('The script failed in the listagousers function')
        sys.exit(1)

# Simple class for attributing items as they are read
class Items:
    def __init__(self, json):
        self.type = json['type']
        self.id = json['id']
        self.title = json['title']
        self.name = json['name']
        self.access = json['access']
        self.url = '{0}/home/item.html?id={1}'.format(urlOrg, self.id)
        self.created = time.strftime('%m/%d/%Y', time.localtime(int(json['created'])/1000))
        self.modified = time.strftime('%m/%d/%Y', time.localtime(int(json['modified'])/1000))
        self.size = round((int(json['size']) / 1024 ** 2.0), 2)


# Create a list of each users items, in and out of folders, and acquire important item details at same time
def itemlister():
    try:
        r = requests.get('{0}/sharing/rest/content/users/{1}?f=json&token={2}'.format(urlOrg, owner, token))
        listfolders = r.json()['folders']
        items = r.json()['items']
        itemscraper(items, False, '', token)
        for folder in listfolders:
            r = requests.get('{0}/sharing/rest/content/users/{1}/{2}?f=json&token={3}'.format(urlOrg, owner, folder['id'], token))
            items = r.json()['items']
            itemscraper(items, True, folder['title'], token)
    except:
        errorhandler('The script failed in the itemlister function')
        sys.exit(1)


# Acquire important item details (used in itemlister() function)
def itemscraper(itemlist, foldertrigger, foldertitle, tokenkey):
    try:
        for _item in itemlist:
            token = tokenkey
            j = Items(_item)
            attempts = 0
            while True:
                try:
                    if j.type == 'Service Definition':
                        break
                    elif j.type == 'Feature Service':
                        r = requests.get('{0}/sharing/rest/portals/fLeGjb7u4uXqeF9q/usage?f=json&startTime={1}&endTime={2}&period={3}&vars=num&groupby=name&etype=svcusg&name={4}&stype=features&token={5}'.format(urlOrg, startTime, endTime, period, j.name, token))
                        annualburn = (j.size / 10) * 2.4 * 12
                        itemwriter(r, j, annualburn, foldertrigger, foldertitle)
                    elif j.type == 'Map Service':
                        r = requests.get('{0}/sharing/rest/portals/fLeGjb7u4uXqeF9q/usage?f=json&startTime={1}&endTime={2}&period={3}&vars=num&groupby=name&etype=svcusg&name={4}&stype=tiles&token={5}'.format(urlOrg, startTime, endTime, period, j.name, token))
                        annualburn = (j.size / 1024) * 1.2 * 12
                        itemwriter(r, j, annualburn, foldertrigger, foldertitle)
                    else:
                        r = requests.get('{0}/sharing/rest/portals/fLeGjb7u4uXqeF9q/usage?f=json&startTime={1}&endTime={2}&period={3}&vars=num&groupby=name&name={4}&token={5}'.format(urlOrg, startTime, endTime, period, j.id, token))
                        annualburn = (j.size / 1024) * 1.2 * 12
                        itemwriter(r, j, annualburn, foldertrigger, foldertitle)
                except KeyError:
                    print(traceback.format_exc())
                    time.sleep(10)
                    token = tokengenerator()
                    print('Regenerated Token')
                    continue
                except:
                    print(traceback.format_exc())
                    attempts += 1
                    if attempts <= 3:
                        print('Sleeping 120')
                        time.sleep(120)
                        print('Trying again...')
                        continue
                    else:
                        print('Max attempts recorded on item ID: {0}'.format(j.id))
                        print('...Skipping to next item.')
                        break
                break
    except:
        errorhandler('The script failed in the itemscraper function')
        sys.exit(1)


# Write item details to a CSV (used in itemscraper() function)
def itemwriter(viewlist, itemobject, itemburn, foldertrigger, foldertitle):
    try:
        views = 0
        for bin in viewlist.json()['data']:
            for num in bin['num']:
                views += int(num[1])
        if foldertrigger:
            writeCSV.writerow([itemobject.title, itemobject.type, itemobject.id, owner, department, views, itemobject.access, foldertitle, itemobject.url, itemobject.created, itemobject.modified, itemobject.size, itemburn, login])
        else:
            writeCSV.writerow([itemobject.title, itemobject.type, itemobject.id, owner, department, views, itemobject.access, '(Home)', itemobject.url, itemobject.created, itemobject.modified, itemobject.size, itemburn, login])
    except:
        errorhandler('The script failed in the itemwriter function')
        sys.exit(1)


# Simple class for attributing items as they are read
class ScrapedItems:
    def __init__(self, row):
        self.title = row[0].strip()
        self.type = row[1].strip()
        self.id = row[2].strip()
        self.owner = row[3].strip()
        self.department = row[4].strip()
        self.views = row[5].strip()
        self.access = row[6].strip()
        self.folder = row[7].strip()
        self.url = row[8].strip()
        self.created = row[9].strip()
        self.modified = row[10].strip()
        self.size = row[11].strip()
        self.annualcreditburn = row[12].strip()
        self.lastlogin = row[13].strip()

def writeallattributes(row):
    try:
        return '"' + row.type + '","' + row.id + '","' + row.owner + '","' + row.department + '","' + row.views + '","' + row.access + '","' + row.folder + '","' + row.url + '","' + row.created + '","' + row.modified + '","' + row.size + '","' + row.annualcreditburn + '","' + row.lastlogin + '"'
    except:
        errorhandler('The script failed in the writeallattributes function')

def csvwriter(title, dictionary, tabtext, sortcolumn='M', lastcolumn='N', burncolumn='M'):
    try:
        with open(os.path.join(subreportFolder, title + '.csv'), 'w', newline='') as fp:
            writeCSV = csv.writer(fp, delimiter=',', quoting=csv.QUOTE_ALL)
            if lastcolumn == 'N':
                writeCSV.writerow(['Title', 'Type', 'ID', 'Owner', 'Department', 'Views', 'Access', 'Folder', 'URL', 'Created', 'Modified','Size', 'Annual Credit Burn', 'Last Login'])
            else:
                writeCSV.writerow(['Department', 'Annual Credit Burn'])
            for p in dictionary.items():
                fp.write('"%s",%s\n' % p)
        booktitle = excelApplication.Workbooks.Open(os.path.join(subreportFolder, title + '.csv'))
        sheettitle = booktitle.Worksheets(title)
        sheettitle.UsedRange.Sort(Key1=sheettitle.Range('{0}1'.format(sortcolumn)), Order1=2, Orientation=1)
        sheettitle.UsedRange.Copy()
        tabtitle = bookReport.Worksheets(tabtext)
        tabtitle.Paste(tabtitle.Range('a3'))
        booktitle.Close(os.path.join(subreportFolder, title + '.csv'))
        tabtitle.Columns(burncolumn).NumberFormat = '#,##0.00'
        tabtitle.Columns('A:{0}'.format(lastcolumn)).AutoFit()
        tabtitle.Range('A3:{0}3'.format(lastcolumn)).Font.Bold = True
        tabtitle.Select()
        tabtitle.Range('A3').Select()
        tabtitle.ListObjects.Add().TableStyle = 'TableStyleLight9'
    except:
        errorhandler('The script failed in the csvwriter function')


# This function formats and prints error handling / logging
def errorhandler(logstring='Script failed.'):
    log.critical(logstring)
    log.critical(traceback.format_exc())
    print(traceback.format_exc())


# region Config Parser
try:
    scriptDirectory = os.path.dirname(__file__)
    config = ConfigParser()
    config.read(os.path.join(scriptDirectory, 'Config_AGO_Usage.cfg'))
except:
    print('Could not read config file')
    print(traceback.format_exc())
    sys.exit(1)
# endregion

# region Set-Up Log
try:
    print('Configuring log file...')
    log_file_path = os.path.join(scriptDirectory, 'Log', config['Logging']['loggingFile'])
    log = logging.getLogger('AGO_USAGE')
    log.setLevel(logging.INFO)
    hdlr = logging.FileHandler(log_file_path)
    hdlr.setLevel(logging.INFO)
    hdlrFormatter = logging.Formatter('%(name)s - %(levelname)s - %(asctime)s - %(message)s', '%m/%d/%Y  %I:%M:%S %p')
    hdlr.setFormatter(hdlrFormatter)
    log.addHandler(hdlr)
    log.info('Script Started...')
    log.info('Log configured.')
except:
    print('Could not properly set up the log file')
    print(traceback.format_exc())
    sys.exit(1)
# endregion

# region Declare Script Variables
try:
    print('Setting parameters...')
    urlOrg = config['OrganizationCredentials']['urlOrg']
    user = config['OrganizationCredentials']['admin']
    password = config['OrganizationCredentials']['password']
    directory = config['Directories']['mainFolder']
    rawOutputFolder = os.path.join(directory, 'RawOutput')
    reportFolder = os.path.join(directory, 'Reports')
    subreportFolder = os.path.join(directory, 'SubReports')
    departmentCSV = os.path.join(scriptDirectory, config['FileNames']['csvDepartment'] + '.csv')
    reportDate = str(date.today().strftime('%Y%m%d'))
    itemCSVFileName = '{0}_{1}.csv'.format(config['FileNames']['csvItem'], reportDate)
    itemCSV = os.path.join(rawOutputFolder, itemCSVFileName)
    listDepartments = list(config['Departments']['listDepartments'].split('\n'))
    reportTemplate = '{0}.xlsx'.format(config['FileNames']['xlsxTemplate'])
    report = '{0}_{1}.xlsx'.format(config['FileNames']['xlsxReport'], reportDate)
    dictDepartment = {}
    dictAGOUsers = {}
    endTime = int(time.time()) * 1000
    startTime = endTime - (int(config['Query']['days']) * 86400000)
    period = '1d'
except:
    errorhandler('The script failed in the Setting Parameters region.')
    sys.exit(1)
# endregion

# region Write item details to CSV
try:
    log.info('Writing item details to CSV...')
    read_departmentlist()
    token = tokengenerator()
    listagousers()
    with open(itemCSV, 'w', newline='') as openCSV:
        writeCSV = csv.writer(openCSV, delimiter=',', quoting=csv.QUOTE_ALL)
        writeCSV.writerow(['Title', 'Type', 'ID', 'Owner', 'Department', 'Views', 'Access', 'Folder', 'URL', 'Created', 'Modified', 'Size', 'Annual Credit Burn', 'Last Login'])
        for k, v in dictAGOUsers.items():
            owner = k
            login = v
            if owner in dictDepartment:
                department = dictDepartment[owner]
            else:
                department = 'UNKNOWN'
            print('Processing {0}'.format(owner))
            itemlister()
except:
    errorhandler('The script failed in the Write item details to CSV region')
    sys.exit(1)
# endregion


# region Reporter
try:
    log.info('Creating reports...')
    with open(itemCSV, 'r') as openCSV:
        reader = csv.reader(openCSV)
        dictAllItems = {}
        dictDepartmentBurn = {}
        for dept in listDepartments:
            dictDepartmentBurn[dept] = 0
        dictEnterpriseItems = {}
        dictNoLogin = {}
        dictFSPrivate = {}
        dictFSShared = {}
        dictWebMap = {}
        dictWebApplication = {}
        dictOther = {}
        n = True
        for row in reader:
            if n:
                n = False
                continue
            else:
                s = ScrapedItems(row)
                if s.department:
                    dictDepartmentBurn[s.department] += float(s.annualcreditburn)
                else:
                    dictDepartmentBurn['UNKNOWN'] += float(s.annualcreditburn)
                if s.department == 'ENTERPRISE':
                    if s.type == 'Map Service':
                        dictEnterpriseItems[s.title] = writeallattributes(s)
                    else:
                        dictEnterpriseItems[s.title + ' '] = writeallattributes(s)
                else:
                    if s.type == 'Map Service':
                        dictAllItems[s.title] = writeallattributes(s)
                    else:
                        dictAllItems[s.title + ' '] = writeallattributes(s)
                if datetime.strptime(str(s.lastlogin), '%m/%d/%Y') < (datetime.today() - timedelta(days=60)):
                    dictNoLogin[s.title] = writeallattributes(s)
                if s.access == 'private' and s.type == 'Feature Service':
                    dictFSPrivate[s.title] = writeallattributes(s)
                elif s.access != 'private' and s.type == 'Feature Service':
                    dictFSShared[s.title] = writeallattributes(s)
                elif s.type == 'Web Map':
                    dictWebMap[s.title] = writeallattributes(s)
                elif s.type == 'Web Mapping Application':
                    dictWebApplication[s.title] = writeallattributes(s)
                else:
                    dictOther[s.title] = writeallattributes(s)

    excelApplication = win32.gencache.EnsureDispatch('Excel.Application')
    excelApplication.Visible = True
    excelApplication.DisplayAlerts = False
    bookTemplate = excelApplication.Workbooks.Open(os.path.join(scriptDirectory, reportTemplate))
    bookTemplate.SaveAs(os.path.join(reportFolder, report))
    bookTemplate.Close(os.path.join(scriptDirectory, reportTemplate))
    bookReport = excelApplication.Workbooks.Open(os.path.join(reportFolder, report))
    tabIndex = bookReport.Worksheets('Index')

    csvwriter(title='AllItems', dictionary=dictAllItems, tabtext='Non-Enterprise Items')
    csvwriter(title='DepartmentBurn', dictionary=dictDepartmentBurn, tabtext='Credit Burn by Department', sortcolumn='B', lastcolumn='B', burncolumn='B')
    csvwriter(title='EnterpriseItems', dictionary=dictEnterpriseItems, tabtext='Enterprise Items')
    csvwriter(title='NoLogin', dictionary=dictNoLogin, tabtext='No Recent Logins')
    csvwriter(title='FSPrivate', dictionary=dictFSPrivate, tabtext='Feature Services - Private')
    csvwriter(title='FSShared', dictionary=dictFSShared, tabtext='Feature Services - Shared')
    csvwriter(title='WebMap', dictionary=dictWebMap, tabtext='Web Maps', sortcolumn='F')
    csvwriter(title='WebApplication', dictionary=dictWebApplication, tabtext='Web Applications', sortcolumn='F')
    csvwriter(title='Other', dictionary=dictOther, tabtext='Other Items')
    tabIndex.Cells(1, 1).Value = str(tabIndex.Cells(1, 1).Value) + ' - ' + str(date.today().strftime('%m/%d/%Y'))
    tabIndex.Select()
    bookReport.Save()
    log.info('Script complete.')
except:
    errorhandler('The script failed in the Subreporter region')
    sys.exit(1)
# endregion
