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

url="https://www.google.com/maps/place/"+sys.argv[1]

driver.get(url)

time.sleep (1)

deny=driver.find_element(By.CLASS_NAME,"lssxud").click()
#deny2=deny.find_element(By.XPATH, "./button")
#deny2.click()

print()
print(Back.GREEN + Fore.BLACK + "Ausschnitt auswählen und Enter drücken...")
input()

time.sleep(2)

close=driver.find_element(By.CLASS_NAME,"gYkzb")

button=close.find_element(By.XPATH, "./button").get_attribute("jsaction")
if "drawer.open" not in button:
    close.click()

time.sleep(1)

driver.execute_script("document.querySelector('#minimap').style.display = 'none';")
driver.execute_script("document.querySelector('.app-bottom-content-anchor').style.display = 'none';")
driver.execute_script("document.querySelector('#vasquette').style.display = 'none';")
driver.execute_script("document.querySelector('.scene-footer-container').style.display = 'none';")
driver.execute_script("document.querySelector('#QA0Szd').style.display = 'none';")
#driver.execute_script("document.querySelector('body').style.pointerEvents = 'none';")
#body=driver.find_element(By.TAG_NAME, "body")
#driver.execute_script("arguments[0].style.pointerEvents='none'", body)

time.sleep(1)

element=driver.find_element(By.ID, "content-container")
filename='src\\main\\resources\\com\\example\\financingtool\\standort.png'
element.screenshot(filename)

time.sleep(1)

def bounding_box_screenshot(bounding_box, filename):
    base_image = Image.open(filename)
    cropped_image = base_image.crop(bounding_box)
    base_image = base_image.resize(cropped_image.size)
    base_image.paste(cropped_image, (0, 0))
    base_image.save(filename)
    return base_image

#bounding_box = (400, 000, 800, 485)
#bounding_box_screenshot(bounding_box, filename) # Screenshot the bounding box (400, 000, 800, 485)

screenshot=Image.open(filename)
screenshot.show()

driver.close()