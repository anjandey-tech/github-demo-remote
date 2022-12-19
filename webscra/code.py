from lib2to3.pgen2 import driver
from selenium import webdriver
import selenium
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
import requests
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

df = pd.DataFrame({})

driver = webdriver.Edge(executable_path='D:\edgedriver_win32\msedgedriver.exe')
driver.get('https://persistentsystems.sharepoint.com/sites/Pi/Search/SitePages/People.aspx#k=#s=1121#l=1033')

element_waited = WebDriverWait(driver, 50).until(
EC.presence_of_element_located((By.XPATH, '//*[@id="divSearchTable"]/table/tbody/tr[1]/td[2]/div[1]/a'))
)
num = 1
time.sleep(5)

while(num < 22421):
    element_waited = WebDriverWait(driver, 50).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="divSearchTable"]/table/tbody'))
    )
    element = driver.find_element(By.XPATH, '//*[@id="divSearchTable"]/table/tbody')
    children = element.find_elements(By.XPATH, './child::*')
    for i in range(len(children)):
        element_waited = WebDriverWait(driver, 50).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="divSearchTable"]/table/tbody'))
        )
        element_empid = element.find_element(By.XPATH, f'./tr[{i+1}]/td[2]/div[2]')
        element_name = element.find_element(By.XPATH, f'./tr[{i+1}]/td[2]/div[1]/a')
        element_email = element.find_element(By.XPATH, f'./tr[{i+1}]/td[4]/div[1]/a')
        element_address = element.find_element(By.XPATH,f'./tr[{i+1}]/td[2]/div[4]')
    
        temp_num = element.find_element(By.XPATH, f'./tr[{i+1}]/td[4]')
        if(len(temp_num.find_elements(By.XPATH, './child::*')) >= 3):
            element_number = element.find_element(By.XPATH, f'./tr[{i+1}]/td[4]/div[3]')
        else:
            element_number = ''
    
        element_designation = element.find_element(By.XPATH, f'./tr[{i+1}]/td[2]/div[3]')
        element_ou = element.find_element(By.XPATH, f'./tr[{i+1}]/td[6]/div[1]')
        element_bu = element.find_element(By.XPATH, f'./tr[{i+1}]/td[6]/div[2]')
        element_du = element.find_element(By.XPATH, f'./tr[{i+1}]/td[6]/div[3]')
        element_pr = element.find_element(By.XPATH, f'./tr[{i+1}]/td[6]/div[4]')

        if(type(element_number) == str):
            temp = {'Emp_id': element_empid.text, 
                    'Name': element_name.text,
                    'Email': element_email.text,
                    'Phone number': element_number,
                    'Address': element_address.text, 
                    'Designation': element_designation.text,
                    'OU': element_ou.text,
                    'BU': element_bu.text,
                    'DU': element_du.text,
                    'PR': element_pr.text}

        else:
            temp = {'Emp_id': element_empid.text, 
                    'Name': element_name.text,
                    'Email': element_email.text,
                    'Phone number': element_number.text,
                    'Address': element_address.text, 
                    'Designation': element_designation.text,
                    'OU': element_ou.text,
                    'BU': element_bu.text,
                    'DU': element_du.text,
                    'PR': element_pr.text}

        #print(element_empid.text)

        el = driver.find_element(By.XPATH, f'//*[@id="divSearchTable"]/table/tbody/tr[{i+1}]/td[2]/div[1]/a')
        el.click()
        driver.switch_to.window(driver.window_handles[1])
        element_waited = WebDriverWait(driver, 50).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/form/div[4]/div/div[2]/div[1]/div/div[1]/div[2]/div[1]/div/div/div/div/div/div[1]/div/div/div/div/div[3]/div/div[2]/div[1]/div/div/div[1]/div/div[2]/p'))
        )
        time.sleep(5)
        element_jobfam = driver.find_element(By.XPATH, '/html/body/form/div[4]/div/div[2]/div[1]/div/div[1]/div[2]/div[1]/div/div/div/div/div/div[1]/div/div/div/div/div[3]/div/div[2]/div[1]/div/div/div[1]/div/div[2]/p')

        temp['Job_fammily'] = element_jobfam.text

        df = df.append(temp, ignore_index=True)

        driver.close()
        driver.switch_to.window(driver.window_handles[0])

        #time.sleep(5)
        element_waited = WebDriverWait(driver, 50).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="divSearchTable"]/table/tbody/tr[1]/td[2]/div[1]/a'))
        )
        driver.execute_script("scroll(0, 1600)")

    nextpage = driver.find_element(By.XPATH, '//*[@id="PageLinkNext"]')
    driver.execute_script("arguments[0].click();", nextpage)
    time.sleep(10)
    driver.execute_script('window.scrollTo(0, 0);')
    df.to_csv('finaldata9.csv')

print(df)
get_url = driver.current_url
print("The current url is:"+str(get_url))
driver.quit()
