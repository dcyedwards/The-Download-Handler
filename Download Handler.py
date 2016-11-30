#! python3
# --------------------------------------------------------------------------------------------------------------------------------------
#  Script: Manual Gas updates - File downloader
# --------------------------------------------------------------------------------------------------------------------------------------
#  Author: David Edwards
# --------------------------------------------------------------------------------------------------------------------------------------
#  Date: Written on the thirteenth day of the fifth month of the two-thousandth and sixteenth year of our Lord. (13/05/2016)
# --------------------------------------------------------------------------------------------------------------------------------------
#  Purpose: To download all the gas demand and storage data for the European markets of interest that I download manually  
#           because it's EXCRUCIATINGLY BORING. Heaven knows how many times I've almost fallen asleep doing this.
#           Also this script might possibly cost me my job but hey.
# --------------------------------------------------------------------------------------------------------------------------------------
# Note: To modify this script you will almost certainly need to have some knowledge of HTML, Javascript, regular expressions (regex(es)) 
#       and BeautifulSoup4 in Python. You don't need to know NumPy, Pandas, Petl or anything like that so don't worry okay.
# --------------------------------------------------------------------------------------------------------------------------------------
# Pseudocode:
# ----------
# 1.0 - Necessary Modules
# 2.0 - General Settings
#       - 2.1 - Raw Folder Directory
#       - 2.2 - Proxy Settings
#       - 2.3 - Download Date of Data Settings
#       - 2.4 - Setting File Paths
#               - 2.4.1 - User Downloads Folder
#               - 2.4.2 - Firefox MIME Profile Directory
#       - 2.5 - Setting Firefox browser prefernces
# 3.0 Download Functions
#       - 3.1 - function1: Dutch Border Flows 
#       - 3.2 - function2: Dutch Flows Aggregated
#       - 3.3 - function3: UK Interconnectors
#       - 3.4 - function4: UK LDZ Actuals
#       - 3.5 - function5: UK LDZ Offtake
#       - 3.6 - function6: UK Industrial Offtake Energy
#       - 3.7 - function7: Northern Ireland Flows 
#       - 3.8 - function8: UK DECC
#       - 3.9 - function9: UK Storage
#       - 3.10 - function10: Exports to Ireland
#       - 3.11 - function11: UK Grain LNG
#       - 3.12 - function12: UK South Hook/Dragon LNG/Langeled/BBL
#       - 3.13 - function13: French Demand
#       - 3.14 - function14: French GDF
#       - 3.15 - function15: French TIGF
#       - 3.16 - function16: German BAFA
#       - 3.17 - function17: German IEA
#       - 3.18 - function18: German DESTATIS
#       - 3.19 - function19: All Countries - GIE storage
#       - 3.20 - function20: Dutch Statline
#       - 3.21 - function21: Dutch IEA
#       - 3.22 - function22: Norwegian Production
#       - 3.23 - function23: Norwegian Gas Exports
#       - 3.24 - function24: Belgian Loenhout Inventories Storage
#       - 3.25 - function25: Italian Demand Dgerm
#       - 3.26 - function26: Spanish Demand
#       - 3.27 - function27: Indigenous Production (Norway, Netherlands, UK)
#       - 3.28 - function28: Belgian Inventories LNG
#       - 3.29 - function29: French Sendout
#       - 3.30 - function30: Montoir recorded flows
#       - 3.31 - function31: Fos Tonkin recorded flows
#       - 3.32 - function32: Fos Cavaou recorded flows
#       - 3.33 - function33: Netherlands Services Gate 
# 4.0 - The function_Dictionary dictionary
# 5.0 - The downloadHandler function
#       5.1 - Invoking downloadHandler
# 6.0 - Wrapping Up
#       6.1 - Raw Folder cleanup
#       6.2 - File Migration
# 7.0 - End
# ---------------------------------------------------------------------------------------------------------------------------------------------

# 1.0 - Necessary Modules

import glob, os, win32com.client, getpass, pythoncom, time, datetime, shutil, sys, bs4, requests, re
from bs4 import BeautifulSoup
from dateutil.relativedelta import relativedelta
from threading import Timer
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotVisibleException
from selenium.webdriver.common.by import By
import selenium.webdriver.support.ui as ui
from selenium.webdriver.support.wait import WebDriverWait

# 2.0 - General Settings
# 2.1 - Raw Folder Directory
desktop = 'C:\\Users\\'+ getpass.getuser() +'\Desktop'
if os.getcwd() != desktop:
        os.chdir(desktop)
print('Raw data files saved here: ' + os.getcwd())

# 2.2 - Proxy Settings
#http_proxy = "http://lonproxy:8080"                     # As an aside, I'll specify the proxy server setting(s)... 
#proxyDict = {"http" : http_proxy}                       # ...for my request module to use later on.

# 2.3 - Data Download Date Settings
t = datetime.datetime.today()                           # Creating a folder for these files using today's date
today = t.strftime('%Y%m%d')
          
currentFolder = os.getcwd()+'\\'+today                  # I love this os module.

t1 = datetime.datetime.today()                          # Creating variable for wanted dates
t2 = t1 + relativedelta(months=-2)
t3 = t2.strftime('%d/%m/%Y')

try:
    os.makedirs(currentFolder)                          # Creating and testing for the existence of my current download folder
    print('Folder '+today+' created')
    os.chdir(currentFolder)
    print(os.getcwd())
except FileExistsError:                                 # If it exists, delete and recreate it
    shutil.rmtree(currentFolder)
    os.makedirs(currentFolder)
    print('Folder '+today+' deleted, then re-created')
    os.chdir(currentFolder)
    print(os.getcwd())
   
print(os.path.exists(os.getcwd()))                      # Just checking the current working directory

# 2.4 - File paths
# 2.4.1 - User Downloads Folder
cwdRawFolder = os.getcwd()
#os.chdir("C:\\Python Downloads")
dwnloadFolder = os.getcwd()
DestatisExcel = os.path.expanduser(cwdRawFolder)

# 2.4.2 - Firefox Mime Profile Directory
print('\tAccessing the Firefox download profile on the network. Expect some delay.\n')

# 2.5 - Setting Firefox browser prefernces
myMimeType = "C:\\Users\\"+ getpass.getuser() +"\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\d93TZgtL.default" # Specifying the MIME type I saved in my Firefox profile
prf = webdriver.FirefoxProfile(myMimeType) # Using my MIME type - Get in there (MIME = Multipurpose Internet Mail Extensions)
prf.set_preference("browser.download.folderlist", 2)
prf.set_preference("browser.download.manager.showWhenStarting", False)
prf.set_preference("broswer.download.dir", dwnloadFolder)
prf.set_preference("browswer.helperApps.neverAsk.saveToDisk","text/csv, text/plain, application/pdf, application/vnd.ms-excel")
prf.set_preference("pdfjs.disabled", "true")
br = webdriver.Firefox(prf)
br.implicitly_wait(10) # in seconds

# 3.0 Download Functions

# 3.1 - function1: Dutch Border Flows
def function1(): 
        print('1)\tDownloading Dutch storage data...')
        try:
                br.get('http://dataport.gastransportservices.nl/default.aspx?ReportPath=%2fTransparency%2fFlowAggregated&ReportTitle=FlowAggregated')
                drpArrow = br.find_element_by_xpath(".//input[@alt='Select a value' and @title='Select a value']").click()
                selectAllflows = br.find_element_by_xpath(".//input[@id='ReportViewerControl_ctl04_ctl03_divDropDown_ctl00' and @type='checkbox']").click()
                drpArrow2 = br.find_element_by_xpath(".//input[@alt='Select a value' and @title='Select a value']").click()
                gasday = br.find_element_by_xpath(".//select[@id='ReportViewerControl_ctl04_ctl17_ddValue']/option[@value='2']").click()    # Gas Day
                timeFlows = br.find_element_by_xpath(".//select[@id='ReportViewerControl_ctl04_ctl09_ddValue']/option[@value='1']").click() # 06:00
                timeBox = br.find_element_by_id('ReportViewerControl_ctl04_ctl07_txtValue').clear()                                         # Clear Date
                flowsToday = t + relativedelta(months=-2)
                flowsToday2 = flowsToday.strftime('%d-%m-%Y')                                                                               # Date creation                                                                                                                         
                timeBox = br.find_element_by_id('ReportViewerControl_ctl04_ctl07_txtValue').clear()
                timeBox2 = br.find_element_by_id('ReportViewerControl_ctl04_ctl07_txtValue')
                timeBox2.send_keys(flowsToday2)
                applybtn = br.find_element_by_xpath(".//input[@id='ReportViewerControl_ctl04_ctl00']").click()
                time.sleep(10)
                exportBtn = WebDriverWait(br, 30).until(
                        EC.presence_of_element_located((By.XPATH, ".//img[@id='ReportViewerControl_ctl05_ctl04_ctl00_ButtonImgDown']"))).click()
                time.sleep(10)
                excelExp = WebDriverWait(br, 30).until(
                        EC.presence_of_element_located((By.XPATH, ".//a[@title='Excel']"))).click()
        except ElementNotVisibleException:
                br.refresh()
                time.sleep(10)
                a = br.switch_to.alert
                a.accept()
                time.sleep(10)
                exportBtn = WebDriverWait(br, 30).until(
                        EC.presence_of_element_located((By.XPATH, ".//img[@id='ReportViewerControl_ctl05_ctl04_ctl00_ButtonImgDown']"))).click()
                time.sleep(10)
                excelExp = WebDriverWait(br, 30).until(
                        EC.presence_of_element_located((By.XPATH, ".//a[@title='Excel']"))).click()
        except:
                print('\n\tWebsite Error: Dutch Flows Aggregated storage dtata not downloaded.\n\tPossibly due to a website timeout.\n')
                

# 3.2 - function2: Dutch Flows Aggregated     
def function2():
        tab1 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
        try:
                br.get('http://www.gasunietransportservices.nl/en/dataport-pages/borderpoints/flow-per-networkpoint')
                borderDateStart = br.find_element_by_xpath(".//input[@id='ReportViewerControl_ctl04_ctl07_txtValue' and @type='text']").clear()
                flowsToday = t + relativedelta(months=-2)
                flowsToday2 = flowsToday.strftime('%d-%m-%Y')
                borderDateStart2 = br.find_element_by_xpath(".//input[@id='ReportViewerControl_ctl04_ctl07_txtValue' and @type='text']").send_keys(flowsToday2)
                borderNetwork = br.find_element_by_xpath(".//input[@alt='Select a value' and @title='Select a value']").click()
                time.sleep(20)
                borderSelectAllFlows = br.find_element_by_xpath(".//input[@id='ReportViewerControl_ctl04_ctl05_divDropDown_ctl00' and @type='checkbox']").click()
                borderNetwork2 = br.find_element_by_xpath(".//input[@alt='Select a value' and @title='Select a value']").click()
                borderGasDay = br.find_element_by_xpath(".//select[@id='ReportViewerControl_ctl04_ctl17_ddValue']/option[@value='2']").click()
                applybtnBorder = br.find_element_by_xpath(".//input[@id='ReportViewerControl_ctl04_ctl00']").click()
                time.sleep(10)
                exportBtnBorder = br.find_element_by_xpath(".//img[@id='ReportViewerControl_ctl05_ctl04_ctl00_ButtonImgDown']").click()
                time.sleep(10)
                excelExpBorder = WebDriverWait(br, 30).until(
                        EC.presence_of_element_located((By.XPATH, ".//a[@title='Excel']"))).click()
                print('\tDutch storage data downloaded.')
        except ElementNotVisibleException:
                br.refresh()
                time.sleep(10)
                a = br.switch_to.alert
                a.accept()
                time.sleep(10)
                exportBtnBorder = br.find_element_by_xpath(".//img[@id='ReportViewerControl_ctl05_ctl04_ctl00_ButtonImgDown']").click()
                time.sleep(10)
                excelExpBorder = WebDriverWait(br, 30).until(
                        EC.presence_of_element_located((By.XPATH, ".//a[@title='Excel']"))).click()
                print('\tDutch storage data downloaded.')
        except:
                print('\n\tWebsite Error: Dutch Border Points storage dtata not downloaded.\n\tPossibly due to a website timeout.\n')

# 3.3 - function3: UK Interconnectors
def function3():
        try:
                br.get('http://marketinformation.natgrid.co.uk/gas/DataItemExplorer.aspx')             # Navigating to website
                time.sleep(3)
                br.switch_to.frame(br.find_element_by_tag_name("iframe"))                              # Switching to iframe to access HTML elements
                print('2)\tDownloading UK Interconnectors...')                        
                natgrid1 = br.find_element_by_xpath(".//img[@alt='Expand Demand']").click()
                natgrid2 = br.find_element_by_xpath(".//img[@alt='Expand Exit Point Actuals']").click()
                natgrid3 = br.find_element_by_xpath(".//img[@alt='Expand Interconnector']").click()
                natgrid4 = br.find_element_by_xpath(".//img[@alt='Expand Energy']").click()
                natgrid5 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn41Nodes input[id='tvDataItemn43CheckBox']"))).click()
                natgrid6a = WebDriverWait(br, 15).until(# Reports
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn13 [alt='Expand Reports']"))).click()
                natgrid6b = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn86 [alt='Expand NTS Physical Entry End Of Day (NTSEOD)']"))).click()
                natgrid7b = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn86Nodes input[id='tvDataItemn201CheckBox']"))).click()
                natgrid6 = WebDriverWait(br, 15).until(# Date Selection
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_chkLatest']"))).click()
                natgrid7 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_rdoApplicableFor']"))).click()
                natgrid8 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_txtSpecifyFromDate']")))
                natgrid8.clear()
                natgrid8.send_keys(t3)
                natgrid9 = WebDriverWait(br, 15).until(# Export
                        EC.presence_of_element_located((By.XPATH, "//a[@id='lbtnCSVDaily']"))).click()
                print('\tUK Interconnectors downloaded.')
        except:
                print('\n\tWebsite Error: UK Interconnectors not downloaded.\n')

# 3.4 - function4: UK LDZ Actuals
def function4():
        try:
                print('3)\tDownloading UK LDZ Actuals...')
                tab2 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://marketinformation.natgrid.co.uk/gas/DataItemExplorer.aspx')
                time.sleep(5)
                br.switch_to.frame(br.find_element_by_tag_name("iframe"))
                LDZact1 = WebDriverWait(br, 15).until(# Demand
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn2 [alt='Expand Demand']"))).click()
                LDZact2 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn19 [alt='Expand LDZ Actual']"))).click()
                allDM = 53                      # This variable refers to the first LDZ checkbox values in the HTML for which I am interested.
                allNDM = 79                     # This one refers to the first NDM checkbox values in the HTML for which I am also very interested.
                LDZact3 = WebDriverWait(br, 15).until(# Selecting all DM(LDZ(D + 1))
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn25 [alt='Expand DM']"))).click()
                while True:                     # A simple loop to select the LDZ values I want
                        if allDM > 77:
                                break
                        DM = WebDriverWait(br, 15).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn25Nodes input[id='tvDataItemn" + str(allDM) + "CheckBox']"))).click()
                        allDM += 2
                ndmLDZ = WebDriverWait(br, 15).until(# Selecting all NDM(LDZ(D + 1))
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn26 [alt='Expand NDM']"))).click()
                while True:# Repetition can be your friend
                        if allNDM > 103:
                                break
                        NDM = WebDriverWait(br, 15).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR,"#tvDataItemn19Nodes input[id='tvDataItemn" + str(allNDM) + "CheckBox']"))).click()
                        allNDM += 2
                LDZact4 = WebDriverWait(br, 15).until(                                                              # Date Selection
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_chkLatest']"))).click()
                LDZact5 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_rdoApplicableFor']"))).click()
                LDZact6 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_txtSpecifyFromDate']")))
                LDZact6.clear()
                LDZact6.send_keys(t3)
                LDZact7 = WebDriverWait(br, 15).until(                                                               # Export
                        EC.presence_of_element_located((By.XPATH, "//a[@id='lbtnCSVDaily']"))).click()
                print('\tLDZ Actuals downloaded.')
        except:
                print('\n\tWebsite Error: UK LDZ Actuals data not downloaded.\n')

# 3.5 - function5: UK LDZ Offtake
def function5():
        try:
                print('4)\tDownloading UK LDZ Offtake...')
                tab3 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://marketinformation.natgrid.co.uk/gas/DataItemExplorer.aspx')
                time.sleep(3)
                br.switch_to.frame(br.find_element_by_tag_name("iframe"))
                offtake1 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, ".//img[@alt='Expand Demand']"))).click()
                offtake2 = WebDriverWait(br, 15).until(
                EC.presence_of_element_located((By.XPATH, ".//img[@alt='Expand Exit Point Actuals']"))).click()
                offtake3 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, ".//img[@alt='Expand LDZ Offtake']"))).click()
                offtake4 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, ".//input[@name='tvDataItemn41CheckBox']"))).click()
                offtakedate1 = WebDriverWait(br, 15).until(                                                         # Date Selection
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_chkLatest']"))).click()
                offtakedate2 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_rdoApplicableFor']"))).click()
                offtakedate3 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_txtSpecifyFromDate']")))
                offtakedate3.clear()
                offtakedate3.send_keys(t3)
                offtakedate4 = WebDriverWait(br, 15).until(                                                         # Export
                        EC.presence_of_element_located((By.XPATH, "//a[@id='lbtnCSVDaily']"))).click()
                print('\tLDZ Offtake downloaded.')
        except:
                print('\n\tWebsite Error: UK LDZ Offtake data not downloaded.\n')

# 3.6 - function6: UK Industrial Offtake Energy
def function6():
        try:
                print('5)\tDownloading Industrial Offtake Energy...')
                tab3 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://marketinformation.natgrid.co.uk/gas/DataItemExplorer.aspx')
                time.sleep(3)
                br.switch_to.frame(br.find_element_by_tag_name("iframe"))
                pwr = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn2 [alt='Expand Demand']"))).click()
                pwr2 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn18 [alt='Expand Exit Point Actuals']"))).click()
                pwr3 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn25 [alt='Expand Industrial Offtake']"))).click()
                pwr4 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn25Nodes input[id='tvDataItemn41CheckBox']"))).click()
                print('\tIndustrial Offtake Energy downloaded..')
                print('6)\tDownloading NTS Power Station Energy...')
                pwr5 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn28 [alt='Expand NTS Power Station']"))).click()
                pwr6 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn28Nodes input[id='tvDataItemn44CheckBox']"))).click()
                pwrDwnld = WebDriverWait(br, 15).until(                                                             # Date Selection
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_chkLatest']"))).click()
                pwrDwnld2 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_rdoApplicableFor']"))).click()
                pwrDwnld3 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_txtSpecifyFromDate']")))
                pwrDwnld3.clear()
                pwrDwnld3.send_keys(t3)
                pwrDwnld4 = WebDriverWait(br, 15).until(                                                            # Export
                        EC.presence_of_element_located((By.XPATH, "//a[@id='lbtnCSVDaily']"))).click()
                print('\tNTS Power Station Energy downloaded.')
        except:
                print('\n\tWebsite Error: Industrial Offtake Energy data not downloaded.\n')

# 3.7 - function7: Northern Ireland Flows  
def function7():
        try:
                print('7)\tDownloading Northern Ireland flows at Stranraer...')
                tab4 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://www.gasnetworks.ie/en-IE/Gas-Industry/Transparency/Transportation-Montly-Reports/2016-Reports/')
                time.sleep(3)
                def MonthGrabber():     # This function selects the month I'm interested in.
                        for i in range(1,12):
                                res = requests.get('http://www.gasnetworks.ie/en-IE/Gas-Industry/Transparency/'
                                                   'Transportation-Montly-Reports/2016-Reports/', proxies = proxyDict)
                                res.raise_for_status()
                                soup = bs4.BeautifulSoup(res.text,"html.parser") # Reading all the HTML on the page.
                                tag = soup.select('p > a')                       # Selecting all 'a' elements which are direct children/descendants of 'p'.
                                try:
                                        soupTitle = BeautifulSoup(str(tag[i]),"html.parser") 
                                        soupT = soupTitle.find_all('a', title=True)
                                        tag = soupT[0]['title']                  # Extracting the 'title' attributes from my elements.
                                        tag = tag.split(' Report', 1)[0]         # Stripping away the ' Report' string
                                except IndexError:                               # Error handling
                                        soupTitle = BeautifulSoup(str(tag[i-1]),"html.parser")
                                        soupT = soupTitle.find_all('a', title=True)
                                        tag = soupT[0]['title']
                                        a = tag.split(' Report', 1)[0]
                                        a = str(tag)
                                        print('\tGasnetworks.ie data month: '+a)
                                        gasnet = br.find_element_by_link_text(a)
                                        gasnet.send_keys(Keys.CONTROL + Keys.RETURN)                          
                                        break                                   # Once the condition is met, break/end
                MonthGrabber() # Calling my function
                print('\tNorthern Ireland flows downloaded.')
        except:
                print('\n\tWebsite Error: Northern Ireland flows not downloaded.\n')

# 3.8 - function8: UK DECC 
def function8(): 
        try:
                print('8)\tDownloading DECC gas production and supply...')
                tab5 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('https://www.gov.uk/government/statistics/gas-section-4-energy-trends')
                DECC = br.find_element_by_link_text('Natural gas production and supply (ET 4.2)').click()
                print('\tDECC gas production and supply downloaded.')
        except:
                print('\n\tWebsite Error: DECC gas production and supply not downloaded.\n')

# 3.9 - function9: UK Storage 
def function9(): 
        try:
                print('9)\tDownloading UK storage data...')
                tab6 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://marketinformation.natgrid.co.uk/gas/DataItemExplorer.aspx')
                time.sleep(3)
                br.switch_to.frame(br.find_element_by_tag_name("iframe")) # Not forgetting to switch to the iframe
                ukStorage = WebDriverWait(br, 15).until(                  # Simple HTML element selection
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn15 [alt='Expand Storage']"))).click()
                ukStocks = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn22 [alt='Expand Stock Levels']"))).click()
                ukStocksLng = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn22Nodes input[id='tvDataItemn24CheckBox']"))).click()
                ukStocksMed = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn22Nodes input[id='tvDataItemn25CheckBox']"))).click()
                ukStocksSrt = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn22Nodes input[id='tvDataItemn26CheckBox']"))).click()
                ukStorageDates = WebDriverWait(br, 15).until(            # Date Selection
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_chkLatest']"))).click()
                ukStorageDates2 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_rdoApplicableFor']"))).click()
                ukStorageDates3 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_txtSpecifyFromDate']")))
                ukStorageDates3.clear()
                ukStorageDates3.send_keys(t3)
                ukStorageExp = WebDriverWait(br, 15).until(              # Export
                        EC.presence_of_element_located((By.XPATH, "//a[@id='lbtnCSVDaily']"))).click()
                print('\tUK Storage data downloaded')
        except:
                print('\n\tWebsite Error: UK storage data not downloaded.\n')

# 3.10 - function10: Exports to Ireland
def function10():
        try:
                print('10)\tDownloading Exports to Ireland data')
                tab7 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://marketinformation.natgrid.co.uk/gas/DataItemExplorer.aspx')
                time.sleep(3)
                br.switch_to.frame(br.find_element_by_tag_name("iframe"))
                Ireland = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn2 [alt='Expand Demand']"))).click()
                Ireland2 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn18 [alt='Expand Exit Point Actuals']"))).click()
                Ireland3 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn26 [alt='Expand Interconnector']"))).click()
                Ireland4 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn41 [alt='Expand Energy']"))).click()
                Ireland5 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn41Nodes input[id='tvDataItemn45CheckBox']"))).click()
                IrelandDates = WebDriverWait(br, 15).until( # Date Selection
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_chkLatest']"))).click()
                IrelandDates2 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_rdoApplicableFor']"))).click()
                IrelandDates3 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_txtSpecifyFromDate']")))
                IrelandDates3.clear()
                IrelandDates3.send_keys(t3)
                IrelandExp = WebDriverWait(br, 15).until( # Exporting
                        EC.presence_of_element_located((By.XPATH, "//a[@id='lbtnCSVDaily']"))).click()
                print('\tExports to Ireland downloaded.')
        except:
                print('\n\tWebsite Error: Exports to Ireland Data not downloaded.\n')

# 3.11 - function11: UK Grain LNG
def function11():
        try:
                print('11)\tDownloading GRAIN LNG data...')
                tab8 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('https://extranet.nationalgrid.com/Grain')
                currentYear = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.ID, 'CurrentYearSection'))).click()
                entries = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, ".//select[@name='tabCurrentYear_length']/option[@value='-1']"))).click()
                export = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, ".//a[@id='btnExportCurrentYear']"))).click()
                print('\tGRAIN LNG data downloaded.')
        except:
                print('\n\tWebsite Error: GRAIN LNG data not downloaded.\n')

# 3.12 - function12: UK South Hook/Dragon LNG/Langeled/BBL  
def function12():
        try:
                print('12)\tDownloading South Hook/Dragon LNG/ Langeled/BBL data...')
                tab9 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://marketinformation.natgrid.co.uk/gas/DataItemExplorer.aspx')
                time.sleep(3)
                br.switch_to.frame(br.find_element_by_tag_name("iframe"))
                supplies = WebDriverWait(br, 15).until( # More HTML element selection
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn16 [alt='Expand Supplies']"))).click()
                supplies2 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn19 [alt='Expand Daily Actuals (Physical)']"))).click()
                supplies3 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn24 [alt='Expand Energy']"))).click()
                bacton = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn24Nodes input[id='tvDataItemn40CheckBox']"))).click()
                dragon = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn24Nodes input[id='tvDataItemn56CheckBox']"))).click()
                easington = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn24Nodes input[id='tvDataItemn66CheckBox']"))).click()
                southHook = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "#tvDataItemn24Nodes input[id='tvDataItemn94CheckBox']"))).click()
                grainDates = WebDriverWait(br, 15).until( # Date Selection
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_chkLatest']"))).click()
                grainDates2 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_rdoApplicableFor']"))).click()
                grainDates3 = WebDriverWait(br, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@id='ctrlDateTime_txtSpecifyFromDate']")))
                grainDates3.clear()
                grainDates3.send_keys(t3)
                grainExp = WebDriverWait(br, 15).until( # Exporting...
                        EC.presence_of_element_located((By.XPATH, "//a[@id='lbtnCSVDaily']"))).click()
                print('\tSouth Hook/Dragon LNG/ Langeled/BBL data downloaded.')
        except:
                print('\n\tWebsite Error: South Hook/Dragon LNG/Langeled/BBL data not downloaded.\n')

# 3.13 - function13: French Demand
def function13():
        try:
                print('13)\tDownloading French demand...')
                tab10 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://www.statistiques.developpement-durable.gouv.fr/donnees-ligne/r/pegase.html')
                frDemand = br.find_elements_by_css_selector("li.spip a") # Searching for 'Importations, productions, consommations par énergie, en unité propre' link (mensuelles)
                frD = str(frDemand[4].text)
                frDemand2 = br.find_element_by_link_text(frD).click()
                time.sleep(4)
                br.switch_to_window(br.window_handles[1])                                       # switching to new pop-up window
                frDemand3 = br.find_element_by_partial_link_text('Pégase - Gaz naturel, '
                                                                 'approvisionnement et consommation en France, en TWh PCS').click()
                frDemand4 = br.find_element_by_link_text('Période').click()
                frDemand5 = br.find_element_by_xpath(".//img[@name='SetDimensionOrderButton']").click() # Configuring the dimension order...
                time.sleep(3)
                libelle = br.find_element_by_xpath(".//option[@value='1']").click()
                oneMoveDim = br.find_element_by_id("Move10").click()
                period = br.find_element_by_xpath(".//option[@value='2']").click()
                oneMoveRight = br.find_element_by_id("Move01").click()
                applyFr = br.find_element_by_xpath(".//input[@id='ApplyBtn']").click()         # Downloading...
                afficher = br.find_element_by_xpath(".//input[@class='ShowReport']").click()
                dwnld1 = br.find_element_by_xpath(".//img[@title='Télécharger']").click()
                selection = br.find_element_by_id('MenuCell_DownloadDiv')
                selection2 = selection.find_element_by_link_text('Format Excel de Microsoft (*.xls)').click()
                time.sleep(2)
                br.switch_to_window(br.window_handles[2])
                telecharger2 = br.find_element_by_xpath(".//input[@type='button']").click()
                br.close()
                br.switch_to_window(br.window_handles[1])
                br.close()                                                                      # Closing window
                br.switch_to_window(br.window_handles[0])                                       # switching back to original window
                print('\tFrench demand data downloaded.')
        except:
                print('\n\tWebsite Error: French demand data not downloaded.\n')

# 3.14 - function14: French GDF
def function14(): 
        try:
                print('14)\tDownloading French GDF storage data...')
                tab11 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('https://sam.storengy.com/ContrastFront/pubdgi/suiviStocksFR.action?redirige=true')
                time.sleep(35)
                br.execute_script("document.getElementById('calendar_dateDebut').readOnly = false") # Executing some javascript here to remove the 'read-only' property from the date field. That way I can simply 'send' my date to the field.
                gdf = br.find_element_by_id("calendar_dateDebut")
                gdf.clear()                                                                     # Clearing the date field
                gdf.send_keys(t3)                                                               # Inputting my start date
                gdfSearch = br.find_element_by_xpath(".//input[@onclick='refreshTabOnDatesChange();']").click()
                time.sleep(35)
                gdfExport = br.find_element_by_xpath(".//input[@onclick='exporter();']").click()
                print('\tFrench GDF storage data downloaded.')
        except:
                print('\n\tWebsite Error: French GDF storage data not downloaded.\n')

# 3.15 - function15: French TIGF
def function15(): 
        try:
                print('15)\tDownloading French TIGF storage data...')
                tab12 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('https://tetra.tigf.fr/SBT/public/StockageFiltree.do')
                tigf1 = br.find_element_by_xpath(".//input[@name='validiteDebut']")
                tigf1.clear()                                                                   # Clearing the date field
                tigf1.send_keys(t3)                                                             # Inputting my start date
                tigf2 = br.find_element_by_xpath(".//select[@name='unite' and @id='uniteListe']/option[@value='MWH0DC']").click()
                tigfExport = br.find_element_by_id("boutonExport").click()
                print('\tFrench TIGF data downloaded.')
        except:
                print('\n\tWebsite Error: TIGF data not downloaded.\n')

# 3.16 - function16: German BAFA
def function16(): 
        try:
                print('16)\tDownloading German BAFA data...')
                tab13 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://www.bafa.de/bafa/de/energie/erdgas/index.html')
                bafa1 = br.find_element_by_xpath(".//a[@href='ausgewaehlte_statistiken/egasmon_xls.xls']").click()
                print('\tGerman BAFA data downloaded.')
        except:
                print('\n\tWebsite Error: German BAFA data not downloaded.\n')

# 3.17 - function17: German IEA
def function17():
        try:
                print('17)\tDownloading German IEA data...')
                tab14 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://wds.iea.org/wds/default.aspx')
                try:    # Trying to take into account the potential erratic login behaviour of the site with some error handling
                        las = br.find_element_by_link_text('Purchase Data Points')
                        ieaHome = br.find_element_by_xpath(".//img[@alt='Click here for IEA Data Services Home Page']").click()
                        ieaLogin = br.find_element_by_xpath(".//a[@href='/payment/login.aspx' and @class='ico-login']").click()
                        ieaUser = br.find_element_by_xpath(".//input[@id='ctl00_ctl00_cph1_cph1_ctrlCustomerLogin_LoginForm_UserName']")
                        ieaUser.clear()
                        ieaPass = br.find_element_by_xpath(".//input[@id='ctl00_ctl00_cph1_cph1_ctrlCustomerLogin_LoginForm_Password']")
                        ieaPass.clear()
                        ieaUser.send_keys('WOODMAC')
                        ieaPass.send_keys('DW9T34')
                        ieaLogin2 = br.find_element_by_xpath(".//input[@id='ctl00_ctl00_cph1_cph1_ctrlCustomerLogin_LoginForm_LoginButton']").click()
                        go = br.find_element_by_link_text("GO DIRECTLY TO IEA DATA SERVICES").click()
                        br.switch_to_window(br.window_handles[1])
                        ieaUser2 = br.find_element_by_xpath(".//input[@id='Login']")
                        ieaUser2.clear()
                        ieaUser2.send_keys("WOODMAC")
                        ieaPass2 = br.find_element_by_xpath(".//input[@id='Pwd']")
                        ieaPass2.clear()
                        ieaPass2.send_keys("DW9T34")
                        ieaSign = br.find_element_by_xpath(".//input[@type='submit' and @value='Sign in']").click()
                except NoSuchElementException:                          # If the site is good, then just proceed as usual...
                        ieaUser = br.find_element_by_xpath(".//input[@id='Login']")
                        ieaUser.clear()
                        ieaPass = br.find_element_by_xpath(".//input[@id='Pwd']")
                        ieaPass.clear()
                        ieaUser.send_keys('WOODMAC')  
                        ieaPass.send_keys('DW9T34')
                        ieaSign = br.find_element_by_xpath(".//input[@type='submit' and @value='Sign in']").click()
                GermanIEA1 = br.find_element_by_xpath(".//a[@title='Natural Gas Monthly']").click()
                GermanIEA2 = br.find_element_by_xpath(".//a[@title='Natural Gas Balance']").click()
                GermanDimension = br.find_element_by_xpath(".//button[@id='ActBtn']").click()
                GermanDimension2 = br.find_element_by_xpath(".//div[@id='item6']").click() # Configuring dimension order and stuff...
                germanBal = br.find_element_by_xpath(".//option[@value='2']").click()
                germanMove10 = br.find_element_by_id("Move10").click()
                germanTime = br.find_element_by_xpath(".//option[@value='3']").click()
                germanMove01 = br.find_element_by_id("Move01").click()
                germanNatG = br.find_element_by_xpath(".//option[@value='0']").click()
                germanMove12 = br.find_element_by_id("Move12").click()
                germanCountry = br.find_element_by_xpath(".//option[@value='1']").click()
                germanMove12 = br.find_element_by_id("Move12").click()
                germanDimensionApply = br.find_element_by_xpath(".//input[@value='Apply']").click()
                time.sleep(1)
                germanTJ = br.find_element_by_xpath(".//input[@value='1' and @type='checkbox']").click()
                germanTime = br.find_element_by_link_text("TIME").click()
                time.sleep(1)                                                                  # Ticking checkboxes I want
                baseMonth = 129                                                                # Refering to October 2015
                jan2005 = br.find_element_by_xpath(".//input[@value='0' and @name='WD_Item']") # Using January 2005 as an anchor
                jan2005.send_keys(Keys.END)                                                    # Scrolling down to the bottom of the page
                time.sleep(2)                                               # Waiting 2 seconds for the browser to catch up with the code
                Checkboxes = br.find_elements_by_xpath(".//*[@onclick]")
                for elem in Checkboxes:
                        if elem.get_attribute('name') == 'WD_Item' and int(elem.get_attribute('value')) > baseMonth: # Simple logical element selection
                                elem.click()
                getCountry = br.find_element_by_link_text("COUNTRY").click()                                    # Selecting Germany as country...
                getGermany = br.find_element_by_xpath(".//input[@value='14' and @name='WD_Item']").click()
                gerBal = br.find_element_by_link_text("BALANCE").click()                                        # Selecting the balance...
                gerInd = br.find_element_by_xpath(".//input[@value='0' and @name='WD_Item']").click()           # Indigenous production
                gerImp = br.find_element_by_xpath(".//input[@value='1' and @name='WD_Item']").click()           # Total imports (Entries)
                gerExp = br.find_element_by_xpath(".//input[@value='3' and @name='WD_Item']").click()           # Total Exports (Exits)
                gerStocks = br.find_element_by_xpath(".//input[@value='5' and @name='WD_Item']").click()        # Stock Change
                gerStats = br.find_element_by_xpath(".//input[@value='7' and @name='WD_Item']").click()         # Statistical Differnce
                gerGross = br.find_element_by_xpath(".//input[@value='8' and @name='WD_Item']").click()         # Gross Inland Deliveries (observed)
                gerGrossAdj = br.find_element_by_xpath(".//input[@value='9' and @name='WD_Item']").click()      # Adjusted Gross Inland Deliveries
                getGermany2 = br.find_element_by_xpath(".//img[@alt='View as table' and @title='View as table']").click() # Viewing data as table
                getGermany3 = br.find_element_by_xpath(".//button[@type='button' and @id='ActBtn']").click()    # Downloading
                getGermany4 = br.find_element_by_xpath(".//div[@id='item30']").click()
                getGermany5 = br.find_element_by_xpath(".//div[@id='item6' and @class='menuItem']").click()
                print('\tGerman IEA data downloaded.')
        except:
                print('\n\tWebsite Error: German IEA data not downloaded.\n')

# 3.18 - function18: German DESTATIS        
def function18(): 
        try:
                print('18)\tDownloading DESTATIS data...')
                tab15 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                destatis = requests.get('https://www.destatis.de/EN/FactsFigures/EconomicSectors/'
                        'Energy/Production/Tables/GasSupplier.html#Footnote2', proxies = proxyDict)
                destatis.raise_for_status()
                soup = bs4.BeautifulSoup(destatis.text,"html.parser")    # Creating my soup
                dPeriod = soup.find("th", {"colspan":"1","rowspan":"1"}) # The current month
                cc = soup.find_all("td", {"colspan":"1", "rowspan":"1"}) # The actual figures I want
                someList = []                                             # Creating some list to store my values
                for i in cc[:]:
                        someList.append(i.text)                               # Putting the elements of my soup into my list
                xl = win32com.client.Dispatch("Excel.Application")    # Creating some workbook to put them in...
                xl.DisplayAlerts = False
                wb = xl.Workbooks.Add()
                wb.SaveAs(DestatisExcel + str(dPeriod.text) + ".xlsx") # Set DestatisExcel at beginning of code to your Downloads folder.
                xl.visible = 0                                         # You will need to change the path above. This is mine.
                sht = wb.Worksheets('Sheet1')
                sht.Range('A1').value = 'Specification'                # Creating sheet headers
                sht.Range('B1').value = str(dPeriod.text)+' - MWh'
                sht.Range('C1').value = 'Change on previous month in %'
                sht.Range("A2:A11").Value = [[i] for i in someList[0::3]] # Copying Specification categories from my list into Excel
                sht.Range("B2:B11").Value = [[i] for i in someList[1::3]] # Copying monthly values from my list into Excel
                sht.Range("C2:C11").Value = [[i] for i in someList[2::3]] # Copying values from change in previous from my list into Excel
                wb.Save()                                             # Note on the above: To create an array of rows.. (Also saving changes to workbook)
                xl.DisplayAlerts = True                               # ...for Excel to understand that it's a column...(Also re-enabling alerts)
                xl.Quit()                                             # ...I use [i] rather than simply i...(Oh and finally: quitting Excel)
                print('\tDESTATIS data downloaded.')
        except:
                print('\n\tWebsite Error: DESTATIS data not downloaded.\n\tPlease check Wood Mackenzie Internet Explorer proxy settings.\n')

# 3.19 - function19: All Countries - GIE storage                 
def function19():
        try:
                print('19)\tDownloading all GIE storage data...')
                br.get('http://www.gie.eu/')
                gerGIE = br.find_element_by_link_text('MAPS & DATA').click()
                gerGIE2 = br.find_element_by_link_text('AGSI+ TRANSPARENCY PLATFORM').click()
                br.switch_to_window(br.window_handles[1])
                gerGIE3 = br.find_element_by_xpath(".//img[@id='but_historical']").click()
                gerGIE4 = br.find_element_by_xpath(".//option[@value='09']").click() # German GIE data
                gerGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=09']").click() 
                belGIE = br.find_element_by_xpath(".//option[@value='03']").click()  # Belgium GIE
                belGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=03']").click()
                fraGIE = br.find_element_by_xpath(".//option[@value='08']").click()  # French GIE
                fraGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=08']").click()
                itaGIE = br.find_element_by_xpath(".//option[@value='12']").click()  # Italian GIE
                itaGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=12']").click()
                holGIE = br.find_element_by_xpath(".//option[@value='15']").click()  # Dutch GIE
                holGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=15']").click()
                spaGIE = br.find_element_by_xpath(".//option[@value='21']").click()  # Spanish GIE
                spaGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=21']").click()
                ausGIE = br.find_element_by_xpath(".//option[@value='01']").click()  # Austrian GIE
                ausGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=01']").click()
                bulGIE = br.find_element_by_xpath(".//option[@value='04']").click()  # Bulgarian GIE
                bulGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=04']").click()
                croGIE = br.find_element_by_xpath(".//option[@value='05']").click()  # Croatian GIE
                croGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=05']").click()
                czeGIE = br.find_element_by_xpath(".//option[@value='06']").click() # Czech GIE
                czeGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=06']").click()
                denGIE = br.find_element_by_xpath(".//option[@value='07']").click() # Danish GIE
                denGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=07']").click()
                hunGIE = br.find_element_by_xpath(".//option[@value='10']").click() # Hungarian GIE
                hunGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=10']").click()
                polGIE = br.find_element_by_xpath(".//option[@value='16']").click() # Polish GIE
                polGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=16']").click()
                porGIE = br.find_element_by_xpath(".//option[@value='17']").click() # Portugal GIE
                porGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=17']").click()
                sloGIE = br.find_element_by_xpath(".//option[@value='20']").click() # Slovakian GIE
                sloGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=20']").click()
                ukGIE = br.find_element_by_xpath(".//option[@value='24']").click() # British GIE
                ukGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=24']").click()
                ukrGIE = br.find_element_by_xpath(".//option[@value='25']").click() # Ukranian GIE
                ukGIEdwn = br.find_element_by_xpath(".//a[@href='/history_download.php?code=25']").click()
                print('\tGIE storage data for all countries downloaded')
                br.close()
                br.switch_to_window(br.window_handles[0])
        except:
                print('\n\tWebsite Error: Not all GIE storage data downloaded.\n')

# 3.20 - function20: Dutch Statline 
def function20(): 
        try:
                print('20)\tDownloading Netherlands Statline data...')
                tab16 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://statline.cbs.nl/statweb/')
                search = br.find_element_by_xpath(".//input[@class='input_big SearchBoxWatermark']")
                search.clear()
                search.send_keys('Industrie en energie')
                search.clear()                                  # Repeated because the text isn't really cleared for some reason
                search.send_keys('Industrie en energie')
                zoeken = br.find_element_by_xpath(".//input[@type='image' and @class='buttonzoeken']").click()
                thema = br.find_element_by_xpath(".//input[@class='buttonthema']").click()
                Ind = br.find_element_by_link_text('Industrie en energie').click()
                Ener = br.find_element_by_link_text('Energie').click()
                Aar = br.find_element_by_link_text('Aardgas, aardolie, kolen').click()
                ver = br.find_element_by_link_text('Aardgas; aanbod en verbruik').click()
                ond = br.find_element_by_xpath(".//input[@title='Selecteer alles']").click()
                perioden = br.find_element_by_id('ctl00_ctl00_MainContent_MainContentDataMaster_TopicAndDimTabs_tabtitle_0').click()
                Maanden = br.find_element_by_link_text('Maanden per jaar').click()
                vanaf2010 = br.find_element_by_link_text("Vanaf 2010").click()
                vanaf2016 = br.find_element_by_link_text('2016').click()
                statMonths = br.find_elements_by_xpath(".//input[@type='checkbox']") # Identify all checkboxes for monthly data.
                chosenMth = re.compile(r'^\d+$')                                     # Create my regex to match checkboxes based on digits only.
                for i in statMonths[:]:                                              # Begin my loop through the elements to select checkboxes...
                        if i.get_attribute('value') > str(613):                      # ...whose 'value' attribute is greater than 613...
                                dd = str(i.get_attribute('value'))                   # ...because 614 represents January 2016 (my first month).
                                if chosenMth.search(dd) is not None:                 # Furthermore check that the element begins and ends with digits...
                                        i.click()                                    # ...and if it does, select it - every month available for 2016.                                  
                toonGv = br.find_element_by_xpath(".//input[@alt='Toon de tabelgegevens op basis van de gemaakte selectie.']").click()
                toonGv2 = br.find_element_by_xpath(".//img[@title='Pas de indeling van de tabel aan. Verplaats variabele naar kolommen.']").click()
                statDwn = br.find_element_by_xpath(".//input[@type='image' and @title='Download']").click()
                statDwn2 = br.find_element_by_link_text('Microsoft Excel 2007 en later (gereed voor het maken van een draaitabel)').click()
                time.sleep(2)
                br.switch_to_window(br.window_handles[1])
                br.close()
                br.switch_to_window(br.window_handles[0])
                print('\tNetherlands Statline data downloaded.')
        except:
                print('\n\tWebsite Error: Netherlands Statline data not downloaded.\n')

# 3.21 - function21: Dutch IEA
def function21():
        try:
                print('21)\tDownloading Dutch IEA data...')
                br.get('http://wds.iea.org/wds/default.aspx')
                try:    # As usual, trying to take into account the potential erratic login behaviour of the site with some error handling
                        nlas = br.find_element_by_link_text('Purchase Data Points')
                        nieaHome = br.find_element_by_xpath(".//img[@alt='Click here for IEA Data Services Home Page']").click()
                        nieaLogin = br.find_element_by_xpath(".//a[@href='/payment/login.aspx' and @class='ico-login']").click()
                        nieaUser = br.find_element_by_xpath(".//input[@id='ctl00_ctl00_cph1_cph1_ctrlCustomerLogin_LoginForm_UserName']")
                        nieaUser.clear()
                        nieaPass = br.find_element_by_xpath(".//input[@id='ctl00_ctl00_cph1_cph1_ctrlCustomerLogin_LoginForm_Password']")
                        nieaPass.clear()
                        nieaUser.send_keys('WOODMAC')
                        nieaPass.send_keys('DW9T34')
                        nieaLogin2 = br.find_element_by_xpath(".//input[@id='ctl00_ctl00_cph1_cph1_ctrlCustomerLogin_LoginForm_LoginButton']").click()
                        ngo = br.find_element_by_link_text("GO DIRECTLY TO IEA DATA SERVICES").click()
                        br.switch_to_window(br.window_handles[1])
                        nieaUser2 = br.find_element_by_xpath(".//input[@id='Login']")
                        nieaUser2.clear()
                        nieaUser2.send_keys("WOODMAC")
                        nieaPass2 = br.find_element_by_xpath(".//input[@id='Pwd']")
                        nieaPass2.clear()
                        nieaPass2.send_keys("DW9T34")
                        nieaSign = br.find_element_by_xpath(".//input[@type='submit' and @value='Sign in']").click()
                except NoSuchElementException: # No issues? Alright then. Let's proceed as normal
                        nieaUser = br.find_element_by_xpath(".//input[@id='Login']")
                        nieaUser.clear()
                        nieaPass = br.find_element_by_xpath(".//input[@id='Pwd']")
                        nieaPass.clear()
                        nieaUser.send_keys('WOODMAC')  
                        nieaPass.send_keys('DW9T34')
                        nieaSign = br.find_element_by_xpath(".//input[@type='submit' and @value='Sign in']").click() # Configuring the Dimension Order and stuff like that...
                netherlandsIEA1 = br.find_element_by_xpath(".//a[@title='Natural Gas Monthly']").click()
                netherlandsIEA2 = br.find_element_by_xpath(".//a[@title='Natural Gas Balance']").click()
                netherlandsDimension = br.find_element_by_xpath(".//button[@id='ActBtn']").click()
                netherlandsDimension2 = br.find_element_by_xpath(".//div[@id='item6']").click()
                netherlandsBal = br.find_element_by_xpath(".//option[@value='2']").click()
                netherlandsMove10 = br.find_element_by_id("Move10").click()
                netherlandsTime = br.find_element_by_xpath(".//option[@value='3']").click()
                netherlandsMove01 = br.find_element_by_id("Move01").click()
                netherlandsNatG = br.find_element_by_xpath(".//option[@value='0']").click()
                netherlandsMove12 = br.find_element_by_id("Move12").click()
                netherlandsCountry = br.find_element_by_xpath(".//option[@value='1']").click()
                netherlandsMove12 = br.find_element_by_id("Move12").click()
                netherlandsDimensionApply = br.find_element_by_xpath(".//input[@value='Apply']").click()
                time.sleep(1)
                netherlandsTJ = br.find_element_by_xpath(".//input[@value='1' and @type='checkbox']").click()
                netherlandsTime = br.find_element_by_link_text("TIME").click()
                time.sleep(1)
                netherlandsBaseMonth = 129
                nJan2005 = br.find_element_by_xpath(".//input[@value='0' and @name='WD_Item']")
                nJan2005.send_keys(Keys.END)
                time.sleep(2)
                netherlandsCheckboxes = br.find_elements_by_xpath(".//*[@onclick]")
                for elem in netherlandsCheckboxes:
                        if elem.get_attribute('name') == 'WD_Item' and int(elem.get_attribute('value')) > netherlandsBaseMonth:
                                elem.click()
                netherlandsCountry = br.find_element_by_link_text("COUNTRY").click()                                    # Selecting Germany as country...
                OECD = br.find_element_by_xpath(".//input[@value='0' and @name='WD_Item']").send_keys(Keys.END)
                getNetherlands = br.find_element_by_xpath(".//input[@value='25' and @name='WD_Item']").click()
                netherlandsBal = br.find_element_by_link_text("BALANCE").click()                                        # Selecting the balance...
                netherlandsInd = br.find_element_by_xpath(".//input[@value='0' and @name='WD_Item']").click()           # Indigenous production
                netherlandsImp = br.find_element_by_xpath(".//input[@value='1' and @name='WD_Item']").click()           # Total imports (Entries)
                netherlandsExp = br.find_element_by_xpath(".//input[@value='3' and @name='WD_Item']").click()           # Total Exports (Exits)
                netherlandsStocks = br.find_element_by_xpath(".//input[@value='5' and @name='WD_Item']").click()        # Stock Change
                netherlandsStats = br.find_element_by_xpath(".//input[@value='7' and @name='WD_Item']").click()         # Statistical Differnce
                netherlandsGross = br.find_element_by_xpath(".//input[@value='8' and @name='WD_Item']").click()         # Gross Inland Deliveries (observed)
                netherlandsGrossAdj = br.find_element_by_xpath(".//input[@value='9' and @name='WD_Item']").click()      # Adjusted Gross Inland Deliveries
                netherlandsADj = br.find_element_by_xpath(".//img[@alt='View as table' and @title='View as table']").click() # Viewing data as table
                getNetherlands3 = br.find_element_by_xpath(".//button[@type='button' and @id='ActBtn']").click()        # Downloading
                getNetherlands4 = br.find_element_by_xpath(".//div[@id='item30']").click()
                getNetherlands5 = br.find_element_by_xpath(".//div[@id='item6' and @class='menuItem']").click()
                print('\tDutch IEA data downloaded.')
        except:
                print('\n\tWebsite Error: Dutch IEA data not downloaded.\n')

# 3.22 - function22: Norwegian Production
def function22():
        try:
                print('22)\tDownloading Norwegian production data...')
                tab17 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://www.npd.no/en/')
                NorProd1 = br.find_element_by_link_text('PRODUCTION FIGURES').click()
                NorProd2 = br.find_element_by_partial_link_text('Production figures').click()
                NorProdDownload = WebDriverWait(br, 30).until(
                        EC.presence_of_element_located((By.LINK_TEXT,'Excel'))).click()
                print('\tNorwegian production data downloaded.')
        except:
                print('\n\tWebsite Error: Norwegian production data not downloaded.\n')

# 3.23 - function23: Norwegian Gas Exports
def function23():
        try:
                print('23)\tDownloading Norwegian gas exports...')
                tab18 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('https://www.ssb.no/statistikkbanken/'
                       'selectvarval/Define.asp?subjectcode=&ProductId=&MainTable='
                       'UhMdVareLand&nvl=&PLanguage=1&nyTmpVar=true&CMSSubjectArea='
                       'utenriksokonomi&KortNavnWeb=muh&StatVariant=&checked=true'
                       )
                norGasUntickQuantity1 = br.find_element_by_xpath(".//option[@value='Mengde1']").click()    
                norGasQuantity2 = br.find_element_by_xpath(".//option[@value='Mengde2']").click()  # We want Quantity 2 instead.
                allCountries = br.find_element_by_xpath(".//img[@title='Select all' and @onclick=\"MerkFjernAlle(\'var3\',\'marker\',-1,1)\"]").click()
                exportsNorway = br.find_element_by_xpath(".//select[@multiple='MULTIPLE' and @name='var2']/option[@value='2']").click()
                mineralProducts = WebDriverWait(br, 25).until(
                        EC.presence_of_element_located((By.XPATH, ".//option[@value='GR¿Kap2527']"))).click() # Select Mineral Products (option 25-27)
                naturalGas = br.find_element_by_xpath(".//option[@value='27112100']").click()
                mthList = br.find_elements_by_xpath(".//select[@multiple='MULTIPLE' and @name='var4']") # Looking at HTML container which contains the months
                norwayYears = [2016, 2017]                             # Creating a list of interested years. One will need to add to this going forward.
                norwayMonths = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] # Next I create a list of all months denoted by an integer.
                v = int(datetime.date.today().year)                    # I then create a variable to store the current year. We'll need this soon.
                def GrabNorwayMonth2():
                        for i in mthList:                                                       # Loop through every element in the HTML months container
                                if len(i.get_attribute('value')) == 7:                          # But select only those which are made up of 7 characters
                                        Month1 = str(i.get_attribute('value'))                  # Get the attribute value which is something like 2016M02
                                        Month1split = Month1.rsplit('M',1)[1]                   # Split this up between month and year
                                        g = re.findall(r"[\d{4}\w]+\d+", Month1split)           # Using a regex to retrieve just the month part
                                        g = [int(i) for i in g]                                 # Now convert this month part into an int (List comprehension)
                                        if v in norwayYears:                                    # Meanwhile, let's loop through our norwayYears
                                                vStr = str(v)                                   # Convert that into a string
                                        for itm in g:                                           # Meanwhile lets look through our months list
                                                if itm in norwayMonths and itm < 10:            # If that month satisfies some conditions and is the current month...
                                                        month2 = vStr + 'M0' + str(itm - 1)     # Go back a month (itm - 1) to get the previous month
                                                        selMonth2 = br.find_element_by_xpath(".//option[@value='" + month2 + "']").click()      # Finally select it
                                                else:                                                                                           # Alternate formatting
                                                        month2 = vStr + 'M' + str(itm - 1)
                                                        selMonth2 = br.find_element_by_xpath(".//option[@value='" + month2 + "']").click()
                GrabNorwayMonth2() # I call the function     
                showTable = br.find_element_by_xpath(".//input[@value='Show table >>' and @id='Submit2']").click()
                time.sleep(10)
                exportElems = br.find_elements_by_xpath(".//input[@class='arrowbutton' and @value='OK']") # Finding the list of all export elements.
                exportElems[2].click() # Selecting the third element (Save as Excel)
                print('\tNorwegian gas exports downloaded.')
        except:
                print('\n\tWebsite Error: Norwegian gas exports not downloaded.\n')

# 3.24 - function24: Belgian Loenhout Inventories Storage
def function24(): 
        try:
                print('24)\tDownloading Belgian inventories storage data...')
                tab19 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('https://gasdata.fluxys.com/sdp/Pages/Reports/Inventories.aspx?report=InventoriesStorage')
                try:
                        flxDate = WebDriverWait(br, 25).until(
                                EC.presence_of_element_located((By.XPATH, ".//input[@data-id='datepicker']")))
                        flxDate.clear()
                        flxDate2 = WebDriverWait(br, 25).until(
                                EC.presence_of_element_located((By.XPATH, ".//input[@data-id='datepicker']")))
                        flxDate2.send_keys(t3)
                        flxM3 = br.find_element_by_xpath(".//input[@value='m3']").click()
                        time.sleep(5)
                        flxLoadData = WebDriverWait(br, 25).until(
                                EC.presence_of_element_located((By.XPATH, ".//input[@value='Load Data']"))).click()
                        time.sleep(10)
                        flxDropDwn = br.find_element_by_xpath(".//img[@alt='Export drop down menu']").click()
                        flxExcel = br.find_element_by_xpath(".//a[@title='Excel']")
                except (ElementNotVisibleException, NoSuchElementException):
                        flxHme = br.find_element_by_xpath(".//img[@title='Home']").click()
                        flxEng = br.find_element_by_xpath(".//a[@href='/lang/en/']").click()
                        flxStorage = br.find_element_by_link_text('Storage').click()
                        flxFlows = br.find_elements_by_xpath(".//div[@class='box-text']")[2].click()
                        InventoriesReport = br.find_elements_by_xpath(".//div[@class='box-text']")[0].click()
                        flxDate = WebDriverWait(br, 25).until(
                                EC.presence_of_element_located((By.XPATH, ".//input[@data-id='datepicker']")))
                        flxDate.clear()
                        flxDate2 = br.find_element_by_xpath(".//input[@data-id='datepicker']")
                        flxDate2.send_keys(t3)
                        flxM3 = br.find_element_by_xpath(".//input[@value='m3']").click()
                        time.sleep(10)
                        flxLoadData = WebDriverWait(br, 25).until(
                                EC.presence_of_element_located((By.XPATH, ".//input[@value='Load Data']"))).click()
                        time.sleep(5)
                        flxDropData = WebDriverWait(br, 25).until(
                                EC.presence_of_element_located((By.XPATH, ".//img[@alt='Export drop down menu']"))).click()
                        time.sleep(10)
                        flxExcel = WebDriverWait(br, 25).until(
                                EC.presence_of_element_located((By.XPATH, ".//a[@title='Excel']"))).click()
                except TimeoutException:
                        flxEng = br.find_element_by_xpath(".//a[@href='/lang/en/']").click()
                        flxStorage = br.find_element_by_link_text('Storage').click()
                        flxFlows = br.find_elements_by_xpath(".//div[@class='box-text']")[2].click()
                        InventoriesReport = br.find_elements_by_xpath(".//div[@class='box-text']")[0].click()
                        time.sleep(10)
                        flxDate = WebDriverWait(br, 25).until(
                                EC.presence_of_element_located((By.XPATH, ".//input[@data-id='datepicker']")))
                        flxDate.clear()
                        flxDate2 = br.find_element_by_xpath(".//input[@data-id='datepicker']")
                        flxDate2.send_keys(t3)
                        flxM3 = br.find_element_by_xpath(".//input[@value='m3']").click()
                        time.sleep(10)
                        flxLoadData = WebDriverWait(br, 25).until(
                                EC.presence_of_element_located((By.XPATH, ".//input[@value='Load Data']"))).click()
                        time.sleep(10)
                        flxDropData = WebDriverWait(br, 25).until(
                                EC.presence_of_element_located((By.XPATH, ".//img[@alt='Export drop down menu']"))).click()
                        time.sleep(10)
                        flxExcel = WebDriverWait(br, 25).until(
                                EC.presence_of_element_located((By.XPATH, ".//a[@title='Excel']"))).click()
                        print('\tBelgian inventories storage data downloaded.')
        except:
                print('\n\tWebsite Error: Belgian inventories storage data not doownloaded.\n')

# 3.25 - function25: Italian Demand Dgerm
def function25(): 
        try:
                print('25)\tDownloading Italian demand...')
                tab20 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://dgsaie.mise.gov.it/dgerm/')
                bilancio = br.find_element_by_link_text('Bilancio').click()
                bilancioList = br.find_elements_by_partial_link_text('Bilancio_GAS_') # Put all the files into a list...
                currentBilancio = bilancioList[-1].click()    
                print('\tItalian Demand downloaded')
        except:
                print('\n\tWebsite Error: Italian Dgerm not downloaded.\n')

# 3.26 - function26: Spanish Demand
def function26():
        try:
                print('26)\tDownloading Spanish demand...')
                tab14 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://www.enagas.com/enagas/es/Gestion_Tecnica_Sistema'
                       '/Seguimiento_del_Sistema_Gasista/Boletin_Estadistico_Gas')
                boletines = br.find_elements_by_xpath(".//a[@title='Ver o descargar el fichero']") # Yes - another list of download options.
                currentBoletine = boletines[0].click()                                             # And just clicking on the one I want.
                print('\tSpanish demand downloaded')
        except:
                print('\n\tWebsite Error: Spanish demand data not downloaded.\n')

# 3.27 - function27: Indigenous Production (Norway, Netherlands, UK)
def function27():
        try:
                print('27)\tDownloading Norwegian, Dutch and British indigenous production...')
                tab15 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://wds.iea.org/wds/default.aspx')
                try:    # Once more, using some error handling to take care of the potential erratic login behaviour of the IEA site.
                        try1 = br.find_element_by_link_text('Purchase Data Points')
                        try2 = br.find_element_by_xpath(".//img[@alt='Click here for IEA Data Services Home Page']").click()
                        try3 = br.find_element_by_xpath(".//a[@href='/payment/login.aspx' and @class='ico-login']").click()
                        try4 = br.find_element_by_xpath(".//input[@id='ctl00_ctl00_cph1_cph1_ctrlCustomerLogin_LoginForm_UserName']")
                        try4.clear()
                        try5 = br.find_element_by_xpath(".//input[@id='ctl00_ctl00_cph1_cph1_ctrlCustomerLogin_LoginForm_Password']")
                        try5.clear()
                        try4.send_keys('WOODMAC')
                        try5.send_keys('DW9T34')
                        tryLogin = br.find_element_by_xpath(".//input[@id='ctl00_ctl00_cph1_cph1_ctrlCustomerLogin_LoginForm_LoginButton']").click()
                        try6 = br.find_element_by_link_text("GO DIRECTLY TO IEA DATA SERVICES").click()
                        br.switch_to_window(br.window_handles[1])
                        try1a = br.find_element_by_xpath(".//input[@id='Login']")
                        try1a.clear()
                        try1a.send_keys("WOODMAC")
                        try2a = br.find_element_by_xpath(".//input[@id='Pwd']")
                        try2a.clear()
                        try2a.send_keys("DW9T34")
                        trySign = br.find_element_by_xpath(".//input[@type='submit' and @value='Sign in']").click()
                except NoSuchElementException: # No issues? Alright then. Let's proceed as normal
                        try1b = br.find_element_by_xpath(".//input[@id='Login']")
                        try1b.clear()
                        try2b = br.find_element_by_xpath(".//input[@id='Pwd']")
                        try2b.clear()
                        try1b.send_keys('WOODMAC')  
                        try2b.send_keys('DW9T34')
                        try2Sign = br.find_element_by_xpath(".//input[@type='submit' and @value='Sign in']").click()
                nor1 = br.find_element_by_xpath(".//a[@title='Natural Gas Monthly']").click() # Configuring the dimension order and stuff
                nor2 = br.find_element_by_xpath(".//a[@title='Natural Gas Balance']").click()
                dimensionB = br.find_element_by_xpath(".//button[@id='ActBtn']").click()
                dimensionB2 = br.find_element_by_xpath(".//div[@id='item6']").click()
                configl = br.find_element_by_xpath(".//option[@value='2']").click()
                configMove10 = br.find_element_by_id("Move10").click()
                configTime = br.find_element_by_xpath(".//option[@value='3']").click()
                configMove01 = br.find_element_by_id("Move01").click()
                config2 = br.find_element_by_xpath(".//option[@value='0']").click()
                configMove12 = br.find_element_by_id("Move12").click()
                configCountry = br.find_element_by_xpath(".//option[@value='1']").click()
                configMove12 = br.find_element_by_id("Move12").click()
                configDimensionApply = br.find_element_by_xpath(".//input[@value='Apply']").click()
                time.sleep(1)
                configTJ = br.find_element_by_xpath(".//input[@value='1' and @type='checkbox']").click()
                configTime = br.find_element_by_link_text("TIME").click()
                time.sleep(1)
                basePeriod = 129 # Getting ready to tick my checkboxes
                baseJan2005 = br.find_element_by_xpath(".//input[@value='0' and @name='WD_Item']")
                baseJan2005.send_keys(Keys.END)
                time.sleep(2)
                allMyCheckboxes = br.find_elements_by_xpath(".//*[@onclick]") # Putting all checkbox elements in this list
                for elem in allMyCheckboxes:                                  # Looping through said checkbox list
                        if elem.get_attribute('name') == 'WD_Item' and int(elem.get_attribute('value')) > basePeriod:# My selections
                                elem.click()
                configCountry = br.find_element_by_link_text("COUNTRY").click() # Selecting countries below...
                oecdMarker = br.find_element_by_xpath(".//input[@value='0' and @name='WD_Item']").send_keys(Keys.END) # Scroll to bottom of page
                getNor = br.find_element_by_xpath(".//input[@value='27' and @name='WD_Item']").click()                # Norway
                getNe = br.find_element_by_xpath(".//input[@value='25' and @name='WD_Item']").click()                 # Netherlands
                getUK = br.find_element_by_xpath(".//input[@value='36' and @name='WD_Item']").click()                 # The UK
                configView = br.find_element_by_xpath(".//img[@alt='View as table' and @title='View as table']").click()
                configDownload = br.find_element_by_xpath(".//button[@type='button' and @id='ActBtn']").click()
                configDownload2 = br.find_element_by_xpath(".//div[@id='item30']").click()
                configDownload3 = br.find_element_by_xpath(".//div[@id='item6' and @class='menuItem']").click()
                print('\tNorwegian, Dutch and British indigenous production data downloaded.')
        except:
                print('\n\tWebsite Error: Indigenous production data not downloaded.\n')

# 3.28 - function28: Belgian Inventories LNG
def function28():
        try:
                print('28)\tDownloading Belgian Inventories LNG report...')
                tab16 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('https://gasdata.fluxys.com/SDP/Pages/Reports/Inventories.aspx?report=inventoriesLNG')
                try:
                        flxLngDate = WebDriverWait(br,25).until(
                                EC.presence_of_element_located((By.XPATH, ".//input[@data-id='datepicker']")))
                        flxLngDate.clear()
                        flxLngDate2 = br.find_element_by_xpath(".//input[@data-id='datepicker']")
                        flxLngDate2.send_keys(t3)
                        flxLngM3 = br.find_element_by_xpath(".//input[@value='m3']").click()
                        time.sleep(10)
                        flxLoadData = WebDriverWait(br,25).until(
                                EC.presence_of_element_located((By.XPATH, ".//input[@value='Load Data']"))).click()
                        time.sleep(10)
                        flxLngLoadData = WebDriverWait(br,25).until(
                                EC.presence_of_element_located((By.XPATH, ".//img[@alt='Export drop down menu']"))).click()
                        time.sleep(10)
                        flxLngExcel = WebDriverWait(br,25).until(
                                EC.presence_of_element_located((By.XPATH, ".//a[@title='Excel']"))).click()
                        print('\tBelgian Inventories LNG report downloaded.')
                except (ElementNotVisibleException, NoSuchElementException, TimeoutException):
                        flxLngEng = br.find_element_by_xpath(".//a[@href='/lang/en/']").click()
                        flxLng = br.find_element_by_link_text('LNG terminalling').click()
                        flxLngFlows = br.find_elements_by_xpath(".//div[@class='box-text']")[3].click()
                        flxLngReport = br.find_elements_by_xpath(".//div[@class='box-text']")[0].click()
                        time.sleep(10)
                        flxLngDate = WebDriverWait(br,25).until(
                                EC.presence_of_element_located((By.XPATH, ".//input[@data-id='datepicker']")))
                        flxLngDate.clear()
                        flxLngDate2 = br.find_element_by_xpath(".//input[@data-id='datepicker']")
                        flxLngDate2.send_keys(t3)
                        flxLngM3 = br.find_element_by_xpath(".//input[@value='m3']").click()
                        time.sleep(10)
                        flxLoadData = WebDriverWait(br,25).until(
                                EC.presence_of_element_located((By.XPATH, ".//input[@value='Load Data']"))).click()
                        time.sleep(10)
                        flxLngLoadData = WebDriverWait(br,25).until(
                                EC.presence_of_element_located((By.XPATH, ".//img[@alt='Export drop down menu']"))).click()
                        time.sleep(10)
                        flxLngExcel = WebDriverWait(br,25).until(
                                EC.presence_of_element_located((By.XPATH, ".//a[@title='Excel']"))).click()
                        print('\tBelgian Inventories LNG report downloaded.')
        except:
                print('\n\tWebsite Error: Belgian Inventories LNG report not downloaded.\n')
                
# 3.29 - function29: French Sendout              
def function29(): 
        try:
                print('29)\tDownloading French sendout data(Smart GRTgaz)...')
                br.get('http://www.smart.grtgaz.com/en/flux_physiques/PITTM')
                br.execute_script("document.getElementById('selector-set-range-date-from').readOnly = false") # Executing some javascript to remove the read-only property
                smartDate = br.find_element_by_xpath(".//input[@name='selector-set-range-date-from' and @id='selector-set-range-date-from']").clear()
                smartDate2 = br.find_element_by_xpath(".//input[@name='selector-set-range-date-from' and @id='selector-set-range-date-from']")
                smartDate2.send_keys(t3)
                smartXLS = br.find_element_by_xpath(".//a[@data-format='xlsLink']").click()
                print('\tFrench sendout data downloaded.')
        except:
                print('\n\tWebsite Error: sendout data not downloaded.\n')

# 3.30 - function30: Montoir recorded flows
def function30():
        try:
                print('30)\tDownloading Montoir recorded flows...')
                tab18 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('https://www.elengy.com/en/')
                montoirContracts = br.find_element_by_link_text('CONTRACTS AND OPERATIONS').click()
                time.sleep(3)
                montoirManagement = br.find_element_by_link_text('OPERATIONAL MANAGEMENT').click()
                montoirHistoricals = br.find_element_by_link_text('Historical and scheduled data').click()
                montoirFlows = br.find_element_by_link_text('Flow rates observed at Montoir-de-Bretagne').click()
                time.sleep(1)
                print('\tMontoir recorded flows downloaded.')
        except:
                print('\n\tWebsite Error: Montoir recorded flows not downloaded.\n')

# 3.31 - function31: Fos Tonkin recorded flows
def function31(): 
        try:
                print('31)\tDownloading Fos Tonkin recorded flows...')
                flowsFos = br.find_element_by_link_text('Flow rates observed at Fos Tonkin').click()
                print('\tFos Tonkin recorded flows downloaded.')
        except:
                print('\n\tWebsite Error: Fos Tonkin flows not downloaded.\n')

# 3.32 - function32: Fos Cavaou recorded flows
def function32(): 
        try:
                print('32)\tDownloading Fos-Cavaou recorded flows...')
                tab19 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://www.fosmax-lng.com/en/commercial-section/telechargements.html')
                cavaouFlowLog = br.find_element_by_link_text('Flow rate log.').click()
                print('\tFos-Cavaou recorded flows downloaded')
        except:
                print('\n\tWebsite Error Fos-Cavaou flows not downloaded.\n')

# 3.33 - function33: Netherlands Services Gate    
def function33(): 
        try:
                print('33)\tDownloading Netherlands Services Gate data...')
                tab20 = br.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
                br.get('http://www.gate.nl/en/commercial/services-gate.html')
                usageList = br.find_elements_by_xpath(".//a[@target='_blank' and @class='download']") # Yep, finding all elements in this list
                usageList[-1].click()                                                                 # And simply picking the latest
                print('\tNetherlands Services Gate data downloaded.\n\tAll data downloaded.')
                br.close()
        except:
                print('\n\tWebsite Error: Services Gate data not downloaded.\n')
                br.close()

# 4.0 - The function_Dictionary Dictionary
# ----------------------------------------
# And now, I put all these functions into some dictionary which I shall call function_Dictionary (because I'm not creative in naming objects, you see?)
function_Dictionary = {1:function1, 2:function2, 3:function3, 4:function4,           # The beauty of using a dictionary is that I can call my functions
                      5:function5, 6:function6, 7:function7, 8:function8,            # in any order that I like.
                      9:function9, 10:function10, 11:function11, 12:function12,      # More importantly, I can skip websites which don't work
                      13:function13, 14:function14, 15:function15, 16:function16,
                      17:function17, 18:function18, 19:function19, 20:function20,
                      21:function21, 22:function22, 23:function23, 24:function24,
                      25:function25, 26:function26, 27:function27, 28:function28,
                      29:function29, 30:function30, 31:function31, 32:function32,
                      33:function33}

# 5.0 - The downloadHandler function
#-----------------------------------
# Finally I simply call each function from my dictionary with this other function which I shall call...
def downloadHandler():
        for function in function_Dictionary: # Let's loop through each function in funcition_Dictionary. (The mappings are really important here.)
                try:
                        function_Dictionary[function]()  # And we simply call each function. Easy, isn't it?
                except (ElementNotVisibleException, NoSuchElementException, ):
                        function_Dictionary[function+1]()                     # I'm sure the logic is easy to follow. If a site doesn't work, skip it.
                                                                              # Who needs it anyway?

# 5.1 - Invoking the downloadHandler
# ----------------------------------
downloadHandler()
br.quit()

# 6.0 - Wrapping Up
# -----------------
# 6.1 - Raw Folder cleanup
os.chdir('C:\\Users\\'+ getpass.getuser() +'\Desktop')
print('\n\tMigrating files to:')
print('\t'+ currentFolder)
globCsv = glob.glob("*.csv")  # Criteria for file deletion
globXls = glob.glob("*.xls")
globXlsx = glob.glob("*xlsx")
globPDF = glob.glob("*.pdf")

# 6.2 - File Migration
# --------------------
# Simple loops
for File in globCsv:
    shutil.move(File, currentFolder)

for File in globXls:
    shutil.move(File, currentFolder)

for File in globXlsx:
    shutil.move(File, currentFolder)

for File in globPDF:
    shutil.move(File, currentFolder)

# 7.0 - End
# ---------
#print('\n\tComplete\n\t(Malachi 3:2: But who can endure the day of His coming?\n\tFor He is like a refiner\'s fire.)')
