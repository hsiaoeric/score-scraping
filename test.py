from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import cv2
from matplotlib import pyplot as plt
import pytesseract
driver = webdriver.Firefox()
url = "https://ap.ceec.edu.tw/RegExam/RegInfo/Login?examtype=B"

PID = "R125096132"
born_year = "94"
born_month = "09"
born_date = "04"


driver.get(url)
elem = driver.find_element(By.ID, "PID")
elem.send_keys(PID)
elem = driver.find_element(By.ID, "PBdYear")
elem.send_keys(born_year)
# time.sleep(1)
elem = driver.find_element(By.ID, "PBdMon")
elem.send_keys(born_month)
# time.sleep(1)
elem = driver.find_element(By.ID, "PBdDay")
elem.send_keys(born_date)   
vali_image = driver.find_element(By.ID, "valiCode")
loc = vali_image.location

driver.save_screenshot('screenshot.png')
image = cv2.imread('screenshot.png')
vali = image[1185: 1237, 820:975 ]
plt.imshow(vali)
plt.show()

raw = pytesseract.image_to_string(vali, lang="eng", config='--psm 13 -c tessedit_char_whitelist=0123456789').replace("\n", '')
print(raw)