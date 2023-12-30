from selenium import webdriver
import time
import sys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

driver = webdriver.Firefox()

url="https://www.wien.gv.at/flaechenwidmung/public/search.aspx?__jumpie#magwienscroll"

driver.get(url) 

time.sleep (1)

select_element=driver.find_element(By.ID,"groupSearchOption_searchOption")
select=Select(select_element)
select.select_by_value('Adresse')

time.sleep(1)

addr_input=driver.find_element(By.ID,"adrText")
addr_input.send_keys(str(sys.argv[1]))

addr_suche=driver.find_element(By.ID, "groupSearchParams_adrSuche")

addr_button=addr_suche.find_element(By.XPATH,'./div[2]/input')
addr_button.click()

time.sleep(2)

driver.find_element(By.ID, "GroupOfButtons1_btnSearch_input").click()

time.sleep(3)

list_item=driver.find_element(By.CLASS_NAME, "vieobjlistitem1")
list_item.find_element(By.XPATH, "./td[5]/a").click()

time.sleep(2)

messageBox=driver.find_element(By.CLASS_NAME, "viennaGisMessageBox")
messageBox.find_element(By.XPATH, "./div[4]/input").click()

#driver.close()