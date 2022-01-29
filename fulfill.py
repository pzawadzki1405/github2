from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

driver = webdriver.Firefox()
driver.get("https://6241176-sb1.app.netsuite.com/app/accounting/transactions/salesord.nl?&id=3212986&scrollid=3212986&whence=&cmid=1643470020809_7264")#put here the adress of your page
#elem = driver.find_elements_by_xpath("//*[@type='submit']")#put here the content you have put in Notepad, ie the XPath
#button = driver.find_element_by_id('process') #//Or find button by ID.
#driver.get("https://www.python.org")
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
while(True):
    a = input('1 aby zakonczyc: ')
    if (a=='1'):
        break
driver.close()
