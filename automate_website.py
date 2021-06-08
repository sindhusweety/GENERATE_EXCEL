import os
import time


from RPA.Browser.Selenium import Selenium

from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from selenium.webdriver.support import expected_conditions as EC

from RPA.Excel.Files import Files
from RPA.HTTP import HTTP


CURDIR = os.getcwd() + '/Download/'

browser = Selenium()
#headless is True
browser.open_available_browser(headless=False)

browser.set_download_directory('/home/sindhukumari/PycharmProjects/Upwork/Download')
print(browser)

driver = browser.driver


URL ="website link"

class Login_Report(object):
    try:
        def __init__(self, client_id, username, password):
            self.client_id  = client_id
            self.username = username
            self.password = password
            self.elem = None
            self.wait = None
        def login(self):
            driver.get(URL)
            driver.current_window_handle
            driver.implicitly_wait(5)

            # client id input
            self.elem = driver.find_element_by_xpath('//*[@id="ClientID"]')
            self.elem.send_keys(self.client_id)

            self.elem = driver.find_element_by_xpath('//*[@id="loginButton"]')
            self.elem.click()

            # uid/pw login
            self.wait = WebDriverWait(driver, 20)

            #issue is here? added EC import
            self.elem = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="userNameInput"]')))
            self.elem.send_keys(self.username)

            self.elem = driver.find_element_by_xpath('//*[@id="passwordInput"]')
            self.elem.send_keys(self.password)#removed whitespace at end and changed password $Blu30c3an21 is old

            self.elem = driver.find_element_by_xpath('//*[@id="submitButton"]')
            self.elem.click()

            self.wait = WebDriverWait(driver, 20)

            #Open Status Report

            #self.elem = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="reportMenuSvg"]')))
            #self.elem.click()

            #self.elem = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ReportSearchTerms"]')))
            #self.elem.send_keys('MRI_OPENSTAT')

            #print(driver.window_handles)
            #self.elem = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ReportSearchTableContainer"]/div/div/div[1]/div[1]')))
            #self.elem.click()
######################################################
            #print(driver.window_handles)
            #self.elem = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="MRI_1BTNCANCEL"]')))
            #self.elem.click()

#############################################


            time.sleep(10)

    except:
        pass

class Report(Login_Report):
    def __init__(self, client_id, username, password):
        Login_Report.__init__(self, client_id, username, password)
        Login_Report.login(self)
    #def run_close(self):

class Communication_center(Login_Report):
    def __init__(self, client_id, username, password):
        Login_Report.__init__(self, client_id, username, password)
        Login_Report.login(self)

    def view_box(self):
        #self.elem = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="CommCenterPopupRows"]/div/div[2]/div[2]/a')))
        self.elem = self.wait.until(EC.element_to_be_clickable((By.XPATH, '// *[ @ id = "commCenterFooter"] / div')))
        self.elem.click()
        time.sleep(5)

    ##Here, get first element
    def open_excel(self):
        #self.elem = driver.find_element_by_partial_link_text('Open Tabled Excel')# Clicks on most recent "Open Tabled Excel" button if available
        self.elem = driver.find_element_by_partial_link_text('Open Paged Excel')

        self.elem.click()
        time.sleep(15)

    def remove_rows(self):# Select all button
        self.elem = driver.find_element_by_xpath('//*[@id="CommCenterSelectAll"]')
        self.elem.click()
        # Remove button
        self.elem = driver.find_element_by_xpath('//*[@id="CommCenterRemove"]')
        self.elem.click()
        time.sleep(5)

ComOBJ = Communication_center('****', '****', '****')
ComOBJ.view_box()
ComOBJ.open_excel()
ComOBJ.remove_rows()



