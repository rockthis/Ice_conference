from openpyxl import load_workbook
from selenium import webdriver
from time import sleep
import xlrd
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import requests
import requests
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
from xlsxwriter import Workbook
import xlrd
import re
from selenium.webdriver.chrome.options import Options
from decimal import *
from selenium.webdriver.common.action_chains import ActionChains

from openpyxl import load_workbook

workbook = xlrd.open_workbook('Ice_members.xlsx')
worksheet = workbook.sheet_by_index(0)
rows = worksheet.nrows
company_links = []
for i in range (rows-1):
     company_links.append(worksheet.cell_value(1+i,2))
# driver = webdriver.Chrome('/Users/kostyafrolov/Downloads/chromedriver')
# ua = UserAgent()
# header = {'user-agent': ua.chrome}
# for i in range (len(company_links)):
#     ua = UserAgent()
#     header = {'user-agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36'}
#     print(company_links[i])
#     page = requests.get(
#         'https://www.similarweb.com/website/'+company_links[i],
#         headers=header)
#     soup = BeautifulSoup(page.text, 'lxml')
#     print(soup.prettify())
#     print(soup.find('span',{'class':'rankingItem-value js-countable'}).getText())

chrome_options = Options()
chrome_options.add_extension('/Users/kostyafrolov/Desktop/4.3.0_0.crx');
driver = webdriver.Chrome(executable_path='/Users/kostyafrolov/Downloads/chromedriver', chrome_options=chrome_options)
driver.get('http://www.icetotallygaming.com/exhibitor-list')
sleep(30)
driver.get("chrome-extension://hoklmmgfnpapgjgcpechhaamimifchmp/popup.html")
sleep(5)
soup = BeautifulSoup(driver.page_source, 'lxml')
print(soup.prettify())
# driver.find_element_by_xpath('//*[@id="quickSearch"]').send_keys('google.com')
sleep(15)
driver.quit()


# for i in range(len(company_links)):
#     driver.get('chrome-extension://hoklmmgfnpapgjgcpechhaamimifchmp/popup.html')
#     sleep(3)
#     el = driver.find_element_by_xpath('//*[@id="js-swSearch-input"]').send_keys(company_links[i])
#     builder = ActionChains(driver)
#     builder.move_to_element(el).click(el).perform()
#     sleep(3)
#
#     sleep(2)
#     # driver.find_element_by_xpath('//*[@id="js-swSearch-input"]').send_keys(Keys.RETURN)
