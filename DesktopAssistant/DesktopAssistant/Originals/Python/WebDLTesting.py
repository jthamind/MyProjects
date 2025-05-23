from selenium import webdriver
from selenium.webdriver.common.by import By

driver = webdriver.Chrome()
driver.get('https://github.com/jthamind/DesktopAssistant')

download_button = driver.find_element(By.CSS_SELECTOR, 'div.ButtonGroup-sc-1gxhls1-0:nth-child(1) > button:nth-child(2) > svg:nth-child(1)')
download_button.click()

driver.quit()