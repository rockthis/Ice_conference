from openpyxl import load_workbook
from selenium import webdriver
from time import sleep
import xlrd
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys

workbook = xlrd.open_workbook('Ice_members.xlsx')
worksheet = workbook.sheet_by_index(0)
rows = worksheet.nrows
company_links = []
for i in range (rows-1):
     company_links.append(worksheet.cell_value(1+i,2))

driver = webdriver.Chrome('C:\\Users\\Admin\\Downloads\\chromedriver')

for i in range(len(company_links)):
    driver.get('https://www.similarweb.com')
    sleep(3)
    driver.find_element_by_xpath('//*[@id="js-swSearch-input"]').send_keys(company_links[i])
    sleep(2)
    # driver.find_element_by_xpath('//*[@id="js-swSearch-input"]').send_keys(Keys.RETURN)
