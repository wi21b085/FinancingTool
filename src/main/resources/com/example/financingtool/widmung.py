from selenium import webdriver
from PIL import Image
from colorama import init, Fore, Back, Style
import time
import sys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

init(autoreset=True)

driver = webdriver.Edge()
driver.maximize_window()

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

oList=driver.find_element(By.ID, "searchResult_oList")
oList.find_element(By.XPATH, "./li[1]").click()

time.sleep(1)

driver.find_element(By.ID, "GroupOfButtons1_btnSearch_input").click()

time.sleep(4)

list_item=driver.find_element(By.CLASS_NAME, "vieobjlistitem1")
list_item.find_element(By.XPATH, "./td[5]/a").click()

time.sleep(4)

messageBox=driver.find_element(By.CLASS_NAME, "viennaGisMessageBox")
messageBox.find_element(By.XPATH, "./div[4]/input").click()
mapImage=driver.find_element(By.ID, "mapImage")
time.sleep(2)

driver.execute_script("document.querySelector('#smap').style.display = 'none';")
driver.execute_script("document.querySelector('#mapResize').style.display = 'none';")
driver.execute_script("document.querySelector('#headerResize').style.display = 'none';")
driver.execute_script("document.querySelector('.map-buttons').style.display = 'none';")
driver.execute_script("document.querySelector('#kachelButtons').style.display = 'none';")
driver.execute_script("document.querySelector('#gugContainer').style.display = 'none';")
driver.execute_script("document.querySelector('#scale').style.display = 'none';")
print()
print(Back.GREEN + Fore.BLACK + "Ausschnitt auswählen und Enter drücken...")
input()
driver.execute_script("document.querySelector('body').style.pointerEvents = 'none';")
body=driver.find_element(By.TAG_NAME, "body")
driver.execute_script("arguments[0].style.pointerEvents='none'", body)

time.sleep(2)

element=driver.find_element(By.ID, "mapImage")
filename='src\\main\\resources\\com\\example\\financingtool\\images\\adresse.png'
element.screenshot(filename)

bild=Image.open(filename)
w, h = bild.size

time.sleep(1)

def bounding_box_screenshot(bounding_box, filename):
    base_image = Image.open(filename)
    cropped_image = base_image.crop(bounding_box)
    base_image = base_image.resize(cropped_image.size)
    base_image.paste(cropped_image, (0, 0))
    base_image.save(filename)
    return base_image

bounding_box = (1/4*w, 000, 3/4*w, h)
bounding_box_screenshot(bounding_box, filename) # Screenshot the bounding box (1/4th, 000, 3/4th, max-height)

screenshot=Image.open(filename)
screenshot.show()

driver.close()