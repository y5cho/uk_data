from openpyxl.reader.excel import load_workbook
from selenium import webdriver
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.webdriver import WebDriver 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from collections import defaultdict
import pandas as pd 
from random import randint
from selenium.webdriver.common.keys import Keys

options = webdriver.ChromeOptions()
options.add_argument('--window-size=1920,1080')
options.add_argument('--disable-gpu')
options.add_argument('--headless')
options.add_argument(
    "user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.61 Safari/537.36"
)
browser = webdriver.Chrome("/Users/yeonilcho/Desktop/canvas/chromedriver", chrome_options=options)

def add_column():
    add_button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > div:nth-of-type(1) > a")))
    add_button.click()

def click_mortgage_data():
    button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > div:nth-of-type(1) > div > div:nth-of-type(2) > div > ul > li:nth-of-type(12) > div > span")))
    button.click()

def add_all():
    button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > div:nth-of-type(1) > div > div:nth-of-type(3) > div > ul > li:nth-of-type(1) > span")))
    button.click()     

def add():
    button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(6) > div:nth-of-type(3) > a:nth-of-type(2)")))
    button.click()

def apply():
    button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(3) > form > div > input:nth-of-type(2)")))
    button.click()

def next_page():
    button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > ul > li:nth-of-type(3) > img")))
    button.click()

def remove_column():
    button =  WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > div:nth-of-type(2) > div:nth-of-type(1) > div:nth-of-type(2) > ul > li:nth-of-type(3) > img:nth-of-type(1)")))
    button.click()                                                                   

browser.get("https://fame.bvdinfo.com/version-2021101/fame/Companies/Login/")
browser.find_element_by_id("user").send_keys("")
browser.find_element_by_id("pw").send_keys("")
browser.find_element_by_class_name("ok").click()
time.sleep(5)
mort_button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > div > div:nth-of-type(2) > div:nth-of-type(2) > div:nth-of-type(1) > div > div:nth-of-type(1) > div > ul > li:nth-of-type(9) > span")))
mort_button.click()
time.sleep(1)
mort_mort_button = button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > div > div:nth-of-type(2) > div:nth-of-type(2) > div:nth-of-type(1) > div > div:nth-of-type(2) > div > div > div:nth-of-type(10) > ul > li > span:nth-of-type(2) > a > span")))
mort_mort_button.click()
time.sleep(5)
ok_button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(1) > div:nth-of-type(2) > div > a:nth-of-type(2)")))
ok_button.click()
time.sleep(5)

result_button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(1) > div > div > a")))
result_button.click()
time.sleep(5)

add_column()
for x in range(7):
    remove_column()
    time.sleep(1)
click_mortgage_data()

mortgage_reg = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > div:nth-of-type(1) > div > div:nth-of-type(2) > div > ul > li:nth-of-type(12) > ul > li:nth-of-type(1) > div > span")))
mortgage_reg.click()
time.sleep(3)
add_all()
time.sleep(2)
add()
time.sleep(5)

mortgage_detail = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > div:nth-of-type(1) > div > div:nth-of-type(2) > div > ul > li:nth-of-type(12) > ul > li:nth-of-type(2) > div > span")))
mortgage_detail.click()
time.sleep(1)
mortgage_des = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > div:nth-of-type(1) > div > div:nth-of-type(3) > div > ul > li:nth-of-type(2) > div:nth-of-type(2) > span")))
mortgage_des.click()

persons_enti = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > div:nth-of-type(1) > div > div:nth-of-type(3) > div > ul > li:nth-of-type(6) > div:nth-of-type(2) > span")))
persons_enti.click()
time.sleep(1)

page_button = browser.find_element_by_css_selector("html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > div:nth-of-type(1) > div > div:nth-of-type(1) > div:nth-of-type(2) > div > input")
page_button.send_keys("uk sic")
page_button.send_keys(Keys.RETURN)
time.sleep(3)

sic = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > div:nth-of-type(1) > div > div:nth-of-type(3) > div > ul > li:nth-of-type(2) > div:nth-of-type(2) > span")))
sic.click()
time.sleep(2)

apply()
time.sleep(5)
def get_fame(page_source):
    soup = BeautifulSoup(page_source, 'lxml')
    companies = []
    company_rows = soup.find("table", attrs={"class":"jGrid-table fixed-data"}).find("tbody").find_all("tr")
    nested_dict = defaultdict(lambda: defaultdict(list))
    data_rows = soup.find("table", attrs={"class":"jGrid-table scroll-data"}).find("tbody").find_all("tr", attrs={"data-itemid":True})
    for row in range(len(data_rows)):
        current_row = data_rows[row]
        column = current_row.find_all("td")
        nested_dict[current_row['data-itemid']]['charge_number'].append(column[0].text)
        nested_dict[current_row['data-itemid']]['charge_registered_day'].append(column[1].text)
        nested_dict[current_row['data-itemid']]['charge_created_day'].append(column[2].text)
        nested_dict[current_row['data-itemid']]['charge_status'].append(column[3].text)
        nested_dict[current_row['data-itemid']]['charge_fully_satisfied'].append(column[4].text)
        nested_dict[current_row['data-itemid']]['mortgage_description'].append(column[5].text)
        nested_dict[current_row['data-itemid']]['persons_entitled'].append(column[6].text)
        nested_dict[current_row['data-itemid']]['uk_sic_code'].append(column[7].text)
        


    sample = dict(nested_dict)

    dataid = []
    var1 = []
    var2 = []
    var3 = []
    var4 = []
    var5 = []
    var6 = []
    var7 = []
    var8 = []

    for key, variables in sample.items():
        dataid.append(key)
        var1.append(sample[key]['charge_number'])
        var2.append(sample[key]['charge_registered_day'])
        var3.append(sample[key]['charge_created_day'])
        var4.append(sample[key]['charge_status'])
        var5.append(sample[key]['charge_fully_satisfied'])
        var6.append(sample[key]['mortgage_description'])
        var7.append(sample[key]['persons_entitled'])
        var8.append(sample[key]['uk_sic_code'])

    for company in company_rows:
        z = company.find("td", attrs={"class":"columnAlignLeft"}).find("a")
        try:
            name = z.get_text()
            companies.append(name)
        except AttributeError:
            continue 
    
    return companies, var1, var2, var3, var4, var5, var6, var7, var8

all_companies = []
all_var1 = []
all_var2 = []
all_var3 = []
all_var4 = []
all_var5 = []
all_var6 = []
all_var7 = []
all_var8 = []

page_button = browser.find_element_by_css_selector("html > body > section:nth-of-type(3) > div:nth-of-type(3) > div > div:nth-of-type(2) > ul > li:nth-of-type(2) > input")
page_button.clear()
page_button.send_keys("9100")
page_button.send_keys(Keys.RETURN)
time.sleep(8)

n=2000
for i in range(n):
    try:
        page_source = browser.page_source
        time.sleep(randint(1,3))
        companies, var1, var2, var3, var4, var5, var6, var7, var8 = get_fame(page_source)
        all_companies = all_companies + companies
        all_var1 = all_var1 + var1
        all_var2 = all_var2 + var2
        all_var3 = all_var3 + var3
        all_var4 = all_var4 + var4
        all_var5 = all_var5 + var5
        all_var6 = all_var6 + var6
        all_var7 = all_var7 + var7
        all_var8 = all_var8 + var8
        print(i)
        next_page()
        time.sleep(6)
    except TimeoutError and Exception:
        print(i, "------------------------------------EXCEPTION------------------------------------")


df = pd.DataFrame({'Compnay':all_companies, 'Charge Number':all_var1, 'Charge Registered Date':all_var2, 'Charged Created Date':all_var3, 'Charge Status':all_var4, 'Charge Fully Satisfied Date':all_var5, 'Mortgage Description':all_var6, 'Persons Entitled':all_var7, 'Uk Sic Code':all_var8})
print(df)
df.to_excel('/Users/yeonilcho/Desktop/canvas/ruhm_9100.xlsx', header=True, sheet_name='fame')
