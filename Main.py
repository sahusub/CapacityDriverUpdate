import xlrd
import collections
import selenium
from selenium import webdriver, common
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.select import Select
import xlrd
import unittest
from openpyxl import load_workbook


chrome_options = Options()
chrome_options.add_argument("--disable-extensions")
driver = webdriver.Chrome(
    r'C:\Users\sahusub\Downloads\chromedriver_win32 - latest_78\chromedriver.exe')
driver.set_window_size(1024, 600)
driver.maximize_window()
chrome_options.add_argument('disable-extensions')
chrome_options.add_argument('disable-infobars')

# Login to Dev env

def loginDEVQA():
    
    driver.get('https://novartis3-sb.pvcloud.com/testing/')
    dsn=Select(driver.find_element_by_name('DSN'))
    dsn.select_by_visible_text('NVTSB1QA')
    username = driver.find_element_by_id('Username')
    password = driver.find_element_by_id('UserPass')
    username.send_keys('sahusub')
    password.send_keys('horizon')
    password.submit()
    print('Login is done !!')

def login():

    driver.get('https://novartis.pvcloud.com/planview/MyPlanview/MyPlanview.aspx?ptab=HV_DASH&pt=HOMEVIEW&scode=$None')
    
    print('Login is done !!')

# Search the WP and Open the Capacity Driver screen


def search_open(WP1):
    print('Search & Open: ' + WP1)
    searchbox = driver.find_element_by_id('bannerSearchBox')
    time.sleep(4)
    searchbox.clear()
    #for i in WP1:
    #    searchbox.send_keys(i)
    #   time.sleep(0.5)
    searchbox.send_keys(WP1)
    time.sleep(3)
    searchbox.send_keys(Keys.ENTER)
    WPCode_in_SearchResult = WebDriverWait(driver, 40).until(EC.presence_of_element_located(
        (By.XPATH, r'//*[@id="searchResults"]/tbody/tr[2]/td[4]')))
    if WPCode_in_SearchResult.text == WP1:
        
        element = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, r'//*[@id="searchResults"]/tbody/tr[2]/td[3]/a')))
        element.click()
        # driver.find_element_by_xpath('//*[@id="searchResults"]/tbody/tr[2]/td[3]/a').click()
        time.sleep(3)
        
        
        UCD = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, r'//*[@id="ui-id-8"]')))
        UCD.click()
        #driver.find_element_by_xpath('//*[@id="ui-id-8"]').click()
    else:
        print("WP Code Requested: "+WP1+" But WP Opened: "+WPCode_in_SearchResult.text)
        
        searchbox = driver.find_element_by_id('bannerSearchBox')
        time.sleep(2)
        searchbox.clear()
        for i in WP1:
            searchbox.send_keys(i)
            time.sleep(0.5)
        #searchbox.send_keys(WP1)
        time.sleep(5)
        searchbox.send_keys(Keys.ENTER) 
        
        WPCode_in_SearchResult = WebDriverWait(driver, 40).until(EC.presence_of_element_located(
        (By.XPATH, r'//*[@id="searchResults"]/tbody/tr[2]/td[4]')))
        if WPCode_in_SearchResult.text == WP1:
            
            element = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
                (By.XPATH, r'//*[@id="searchResults"]/tbody/tr[2]/td[3]/a')))
            element.click()
            # driver.find_element_by_xpath('//*[@id="searchResults"]/tbody/tr[2]/td[3]/a').click()
            time.sleep(3)
            UCD = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
                (By.XPATH, r'//*[@id="ui-id-8"]')))
            UCD.click()

# Update the Capacity Driver screen


def upadteCD(atr, val):
    print('Update CD with: ' + atr + ":" + val)
    seq = driver.find_elements_by_tag_name('iframe')
    #print(len(seq))
    time.sleep(3)
    if atr == 'Number of DMCs':
        driver.switch_to_frame(2)
        time.sleep(1)
        ele = driver.find_element_by_name('PVCfg_F_num_dmcs').clear()
        driver.find_element_by_name(
            'PVCfg_F_num_dmcs').send_keys(int(float(val)))
        driver.switch_to.default_content()
    elif atr == 'DMC SP Sourcing':
        driver.switch_to_frame(2)
        time.sleep(1)
        s1 = Select(driver.find_element_by_id('PVCfg_S_Wbs322'))
        print(val)
        s1.select_by_visible_text(val)
        driver.switch_to.default_content()
    elif atr == 'SDTM SP Sourcing':
        driver.switch_to_frame(2)
        time.sleep(1)
        s1 = Select(driver.find_element_by_id('PVCfg_S_Wbs323'))
        print(val)
        s1.select_by_visible_text(val)
        driver.switch_to.default_content()
    elif atr == 'Number of Unique CRF pages':
        driver.switch_to_frame(2)
        time.sleep(1)
        ele = driver.find_element_by_name('PVCfg_F_unique_crf_pages').clear()
        driver.find_element_by_name(
            'PVCfg_F_unique_crf_pages').send_keys(int(float(val)))
        driver.switch_to.default_content()
    elif atr == 'Number of Edit checks':
        driver.switch_to_frame(2)
        time.sleep(1)
        ele = driver.find_element_by_name('PVCfg_F_num_edit_checks').clear()
        driver.find_element_by_name(
            'PVCfg_F_num_edit_checks').send_keys(int(float(val)))
        driver.switch_to.default_content()
    elif atr == 'Number of  Data Transfer Specs':
        driver.switch_to_frame(2)
        time.sleep(1)
        ele = driver.find_element_by_name('PVCfg_F_num_dts').clear()
        driver.find_element_by_name(
            'PVCfg_F_num_dts').send_keys(int(float(val)))
        driver.switch_to.default_content()
    elif atr == 'Number of Interim Analyses with CSR':
        driver.switch_to_frame(2)
        time.sleep(1)
        ele = driver.find_element_by_name('PVCfg_F_trial_analyses_csr').clear()
        driver.find_element_by_name(
            'PVCfg_F_trial_analyses_csr').send_keys(int(float(val)))
        driver.switch_to.default_content()
    else:
        print('Element not found')


# Click on the SAVE button in Capacity Driver screen

def saveCD():
    driver.switch_to_frame(2)
    driver.find_element_by_id('Submit2').click()
    time.sleep(2)


loginDEVQA()
workbook = xlrd.open_workbook(
    r'C:\Users\sahusub\Desktop\CapacityDriver_DEV\PythonData.xlsx')
sheet = workbook.sheet_by_name("Sheet2")
rowcount = sheet.nrows  
colcount = sheet.ncols
result_data = []
WP_Code = []
for curr_row in range(1, rowcount, 1):
    row_data = []
    # WP_Code=[]
    for curr_col in range(0, 1, 1):
        data = str(sheet.cell_value(curr_row, curr_col))
        data1 = str(sheet.cell_value(curr_row, curr_col+1))
        data2 = str(sheet.cell_value(curr_row, curr_col+2))
        row_data.append(data)
        row_data.append(data1)
        row_data.append(data2)
        WP_Code.append(data)
    WP1 = WP_Code
    WP = row_data
    print(WP)
    print(WP1)
    if len(WP_Code) > 1:
        if WP_Code[-1] == WP_Code[-2]:
            upadteCD(WP[1], WP[2])
            # saveCD()
        else:
            saveCD()
            search_open(WP[0])
            upadteCD(WP[1], WP[2])

    else:
        search_open(WP1[0])
        upadteCD(WP[1], WP[2])
        # saveCD()
        

saveCD()







