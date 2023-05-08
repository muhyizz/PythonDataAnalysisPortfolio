#!/usr/bin/env python
# coding: utf-8

# # Teambinder AutoClicker, Processing, and Auto Email

# In[1136]:


def login(id,organisation,password):
    """
    The login function is used to get to the prescribed teambinder url and handling the login page. The input to the login
    function would be the userID, Organisation, and Password.
    """
    url = 'https://asia01.teambinder.com/TeamBinder5/Home/'
    driver.get(url)

    #Main Page Teambinder
    time.sleep(5)
    login_bar = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID,"link-wcsmxdrzrvl")))
    login_bar.click()

    #Teambinder Login Page
    driver.switch_to.window(driver.window_handles[1])
    time.sleep(5)
    driver.find_element(by=By.NAME, value = 'txtUserId').send_keys(id)
    driver.find_element(by=By.NAME, value = 'txtCompanyId').send_keys(organisation)
    driver.find_element(by=By.NAME, value = 'txtPassword').send_keys(password)
    driver.find_element(by=By.XPATH, value='//*[@id="lnkLogon"]').click()
    time.sleep(10)



# In[1137]:


def downloadform():
    """
    The download function autoclicks the required button needed to click to download the form
    """
    #Download Form

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "moduleDropDown"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[2]/div/div[1]/div[4]/div/div/div[1]/div/ul/li[6]/a/span'))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "exportDropDownMenu"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="exportDropDownMenu"]/ul/li[1]/a'))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="exportDropDownMenu"]/ul/li[1]/ul/li[1]/a'))).click()

    #Back To homepage

    time.sleep(5)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="homeButton"]'))).click()
    


# In[1138]:


def xslsPanda():
    """
    The xslsPanda function will find the latest downloaded xsls file and extract it as a Data Frame
    """
    os.chdir(download_dir)
    file_extension = '.xlsx'
    file_one = max(glob.iglob('*{}'.format(file_extension)), key=os.path.getctime)

    # Read the downloaded file into a pandas DataFrame
    dataframe = pd.read_excel(file_one)
    os.chdir(original_dir)
    return dataframe



# In[1139]:


def downloadReview():
    #Download To Review
    """
    The download review function autoclicks the required button needed to click to download the awaiting review list
    """
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divContent_Workflow_3"]/div[2]/table/tbody/tr/td[3]'))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="exportButton"]/div/i'))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/review/div[1]/div[2]/div/div[2]/div/div/ul/li/a'))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/review/div[1]/div[2]/div/div[2]/div/div/ul/li/ul/li[1]/a'))).click()
    time.sleep(10)


# In[1140]:


def loginEnd():
    """
    The loginEnd quit the Selenium Driver and close off all the browsers
    """
    #Close Selinium
    driver.quit()


# In[1142]:


def processform():
    """
    The processform function process the form Data Frame into respective type
    """
    df_cr = df_form[df_form['Type'] == 'CR']
    df_rfi = df_form[(df_form['Type'] == 'RFI-COKUL') | (df_form['Type'] == 'RFI')]
    df_ncr = df_form[df_form['Type'] == 'ENC']
    df_mcc = df_form[df_form['Type'] == 'MCC']
    df_mev = df_form[df_form['Type'] == 'MEV']
    return (df_cr,df_rfi,df_ncr,df_mcc,df_mev)


# In[1143]:


def processunderreview():
    """
    The processreview function process the form Data Frame date column to a date we can compared to
    """
    df_underreview['RequiredBy'] = pd.to_datetime(df_underreview['RequiredBy']).dt.normalize()


# In[1144]:


def printTXT(processedform,overdue,today,tomorrow):
    """
    The printTXT function process the all the data developed from the other function and write into a txt file
    """
    os.chdir(download_dir)
    todaystr = str(today)
    name = todaystr.replace(":", "-") + " Project Name.txt"
    f = open(name, "w")

    #Header
    print("Project Name", file =f)
    print(f"Report Generated on {todaystr} \n", file =f)

    #Document Revie STATUS
    print("DOCUMENT REVIEW STATUS: \n", file =f)
    #Overdue
    print("Document that are overdue: \n", file =f)
    print(tabulate(df_underreview.query('RequiredBy == "%s"' % overdue), headers='keys'), file = f)
    print("", file =f)

    #To be completed today
    print("Document review that are to be completed today: \n", file =f)
    print(tabulate(df_underreview.query('RequiredBy == "%s"' % today), headers='keys'), file = f)
    print("", file =f)

    print("Document to review by tomorrow: \n", file =f)
    print(tabulate(df_underreview.query('RequiredBy == "%s"' % tomorrow), headers='keys'), file = f)
    print("", file =f)

    #FORM STATUS
    #Change Request STATUS
    print("FORM STATUS: \n", file =f)
    print("Change Request:", file =f)
    df_cr = processedform[0]
    print((df_cr['Status'].value_counts()), file=f)
    print("", file =f)
    
    #RFI STATUS
    
    print("RFI Status:", file =f)
    df_rfi = processedform[1]
    print((df_rfi['Status'].value_counts()), file=f)
    print("", file =f)
    
    #NCR STATUS
   
    print("NCR Status:", file =f)
    df_ncr = processedform[2]
    print((df_ncr['Status'].value_counts()), file=f)
    print("", file =f)
    
    #MCC STATUS
    
    print("MCC Status:", file =f)
    df_mcc = processedform[3]
    print((df_mcc['Status'].value_counts()), file=f)
    print("", file =f)
    
    #MEV STATUS
    
    print("MEV Status:", file =f)
    df_mev = processedform[4]
    print((df_mev['Status'].value_counts()), file=f)
    print("", file =f)

    

    f.close()


# In[1145]:


def toEmail():
    """
    The toEmail function uses the win32 interface to allow automatic email of the txt file
    """
    ol = win32com.client.Dispatch('Outlook.Application')
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'Project Name Automated Email - Document and Form Status'
    newmail.To = ''
    newmail.Body = f'This is an automated email breaking down the details of form status and document awaiting review as of {todaystr}'

    os.chdir(download_dir)
    file_extension = '.txt'
    attachment = max(glob.iglob('*{}'.format(file_extension)), key=os.path.getctime)
    attach = os.path.abspath(attachment)
    os.chdir(original_dir)
    newmail.Attachments.Add(attach)
    newmail.Send()



# In[1146]:


from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import time
import os
import glob
import pandas as pd
import win32com.client
from datetime import date
from datetime import timedelta
from tabulate import tabulate
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

#Fixed variables used in multiple function

driver = webdriver.Chrome(ChromeDriverManager().install())
original_dir = r'C:\Users'
download_dir = r'C:\Users\pc\Downloads'
overdue = pd.Timestamp(date.today() - timedelta(days=1))
today = pd.Timestamp(date.today())
tomorrow = pd.Timestamp(date.today() + timedelta(days=1))


if __name__ == '__main__':
    datetime()
    id = input('User ID: ')
    organisation = input('User Organisation: ')
    password = input('Password: ')
    login(id,organisation,password)
    downloadform()
    df_form = xslsPanda()
    downloadReview()
    df_underreview = xslsPanda()
    loginEnd()
    processedform = processform()
    processunderreview()
    printTXT(processedform,overdue,today,tomorrow)

