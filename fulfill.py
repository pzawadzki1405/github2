from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
import pickle

def look_for_order(order):
    driver.find_element_by_id('_searchstring').send_keys(order)
    time.sleep(2)
    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//li[@class='uir-global-search-item']")))
    driver.find_element_by_xpath("//li[@class='uir-global-search-item']/a[1]").click()
    ordernr = driver.find_element_by_xpath("//div[@class='uir-record-id']")
    print(ordernr.text)

def fulfill():
    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "process")))
    driver.find_element_by_id("process").click()
    shipping_status = driver.find_element_by_id('inpt_shipstatus5')
    shipping_status.click()
    time.sleep(2)
    driver.find_element_by_id('inpt_shipstatus5').send_keys("Packed")
    driver.find_element_by_id('spn_multibutton_submitter').click()
    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='uir-record-status']")))
    status = driver.find_element_by_xpath("//div[@class='uir-record-status']")
    if(status.text == 'PACKED'):
        time.sleep(2)

orders = ['WS69974']
driver = webdriver.Firefox()
driver.maximize_window()
time.sleep(2)
driver.get("https://6241176-sb1.app.netsuite.com/app/center/card.nl?sc=-29&whence=")#put here the adress of your page
#elem = driver.find_elements_by_xpath("//*[@type='submit']")#put here the content you have put in Notepad, ie the XPath
#button = driver.find_element_by_id('process') #//Or find button by ID.
#driver.get("https://www.python.org")
cookies = pickle.load(open("cookies.pkl", "rb"))
for cookie in cookies:
    driver.add_cookie(cookie)
driver.get("https://6241176-sb1.app.netsuite.com/app/center/card.nl?sc=-29&whence=")#put here the adress of your page

time.sleep(2)
print(driver.title)
if (driver.title == 'NetSuite Login'):
    username = driver.find_element_by_id("email")
    password = driver.find_element_by_id("password")
    username.send_keys("warehouse12@omacshop.com.tr")
    password.send_keys("Warehouse12")
    print('tu jest logowanie')

    driver.find_element_by_id("login-submit").click()
    time.sleep(5)

for order in orders:
    look_for_order(order)
    fulfill()

while(True):
    a = input('1 aby zakonczyc: ')
    if (a=='1'):
        pickle.dump( driver.get_cookies() , open("cookies.pkl","wb"))
        break
driver.close()
