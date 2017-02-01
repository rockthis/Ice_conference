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
from random import randint
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

linkedin_link = []

def check_exists_by_xpath(xpath):
    try:
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def find_linkedin_companies():
    while True:
        for el in driver.find_elements_by_xpath('//*[@class="title main-headline"]'):
            linkedin_link.append(el.get_attribute('href'))
        if check_exists_by_xpath('//*[@class="next"]'):
            driver.find_element_by_xpath('//*[@class="next"]').click()
            sleep(randint(2,5))
        else:
            break
    return linkedin_link
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


# chrome_options = Options()
# chrome_options.add_extension('/Users/kostyafrolov/Desktop/4.3.0_0.crx'); #mac
driver = webdriver.Chrome('C:\\Users\\Admin\\Downloads\\chromedriver') #windows
# driver = webdriver.Chrome('/Users/kostyafrolov/Downloads/chromedriver')  #mac
driver.get("https://www.linkedin.com")
sleep(randint(2,5))
driver.find_element_by_xpath('//*[@id="login-email"]').send_keys('konstantin@payzoff.com')
sleep(randint(2,5))
driver.find_element_by_xpath('//*[@id="login-password"]').send_keys('kostya13')
sleep(randint(2,5))
driver.find_element_by_xpath('//*[@id="login-password"]').send_keys(Keys.ENTER)
sleep(randint(2,5))
# driver.find_element_by_xpath('//*[@id="login-submit"]').click()
driver.find_element_by_xpath('//*[@id="main-search-box"]').send_keys(Keys.ENTER) #страница поиска
sleep(randint(2,5))
driver.find_element_by_xpath('//*[@id="search-types"]/div/ul/li[4]/a').click() #компании
sleep(randint(2,5))
driver.find_element_by_xpath('//*[@id="facet-I"]/fieldset/legend').click() #отрасль
sleep(randint(2,5))
driver.find_element_by_xpath('//*[@id="facet-I"]/fieldset/div/div/div/button').click() # добавить
sleep(randint(2,5))
driver.find_element_by_xpath('//*[@id="facet-I"]/fieldset/div/div/input').send_keys('Азарт') # вводим значение отрасли
sleep(randint(2,5))
driver.find_element_by_xpath('//*[@id="pagekey-voltron_company_search_internal_jsp"]/div[1]/div/div[2]/ul/li[1]/h4').click()
sleep(randint(2,5))
driver.find_element_by_xpath('//*[@id="facet-CS"]/fieldset/legend').click() #размер компании
sleep(randint(2,5))
temp = ['B','C','D','E','F','G','H','I']
for i in temp:
    driver.find_element_by_xpath('//*[@id="'+i+'-CS-ffs"]').click() #размер компании 1- 10
    sleep(randint(3,6))
    find_linkedin_companies()
    driver.find_element_by_xpath('//*[@id="'+i+'-CS-ffs"]').click()
    sleep(randint(2,7))
#
# driver.find_element_by_xpath('//*[@id="C-CS-ffs"]').click() #размер компании 11 - 50
# sleep(2)
# find_linkedin_companies()
# driver.find_element_by_xpath('//*[@id="C-CS-ffs"]').click()
# sleep(2)
# driver.find_element_by_xpath('//*[@id="D-CS-ffs"]').click() #51-200
# sleep(2)
# find_linkedin_companies()
# driver.find_element_by_xpath('//*[@id="D-CS-ffs"]').click()
# sleep(2)
# driver.find_element_by_xpath('//*[@id="E-CS-ffs"]').click()
# sleep(2)
# find_linkedin_companies()
# driver.find_element_by_xpath('//*[@id="E-CS-ffs"]').click()
# sleep(2)
# driver.find_element_by_xpath('//*[@id="F-CS-ffs"]').click()
# sleep(2)
# find_linkedin_companies()
# driver.find_element_by_xpath('//*[@id="F-CS-ffs"]').click()
# sleep(2)
# driver.find_element_by_xpath('//*[@id="G-CS-ffs"]').click()
# sleep(2)
# find_linkedin_companies()
# driver.find_element_by_xpath('//*[@id="G-CS-ffs"]').click()
# sleep(2)
# driver.find_element_by_xpath('//*[@id="H-CS-ffs"]').click()
# sleep(2)
# find_linkedin_companies()
# driver.find_element_by_xpath('//*[@id="H-CS-ffs"]').click()
# sleep(2)
# driver.find_element_by_xpath('//*[@id="I-CS-ffs"]').click()
# sleep(2)
# find_linkedin_companies()
# driver.find_element_by_xpath('//*[@id="I-CS-ffs"]').click()
# sleep(2)
sleep(10)
print(len(linkedin_link))
# driver.find_element_by_xpath('//*[@id="js-swSearch-input"]').send_keys('ebay.com').send_keys(Keys.ENTER)
# print(driver.find_elements_by_tag_name("iframe"))
# print(driver.find_elements_by_tag_name("iframe")[0].get_attribute('id'))

# driver.find_element_by_xpath('//*[@id="quickSearch"]').send_keys('google.com')

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
