from selenium import webdriver
from selenium.webdriver.common.by import By
import time


driver = webdriver.Firefox(executable_path="C:\\Users\\soufi\\Documents\\repositories\\Destination-Insights-with-Google\\webdrivers\\geckodriver.exe")
driver.get("https://destinationinsights.withgoogle.com/")


assert "Destination Insights with Google" in driver.title


# ACCEPT COOKIES
cookies = driver.find_element(By.XPATH, '/html/body/div[1]/div/span[2]/a[2]')
cookies.click()

# SET FILTER
time.sleep(5)
driver.execute_script('document.getElementById("select_value_label_0").firstChild.innerHTML="France"')
driver.execute_script('document.getElementById("select_value_label_2").firstChild.innerHTML="Spain"')
driver.execute_script('document.getElementById("select_value_label_4").firstChild.innerHTML="International"')
driver.execute_script('document.getElementById("select_value_label_5").firstChild.innerHTML="Air"')

# SUBMIT FILTER
#submit_btn = driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div/div/div[2]/div[2]/button')
#submit_btn.click()


'''
# GET ELEMENTS
time.sleep(5)
origin = driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div/div/div[1]/div[1]/md-content[1]/md-select/md-select-value/span[1]')
destination = driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div/div/div[1]/div[2]/md-content[1]/md-select/md-select-value/span[1]')
trip_type = driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div/div/div[2]/md-content[1]/md-select/md-select-value/span[1]')
demand_category = driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div/div/div[2]/md-content[2]/md-select/md-select-value/span[1]')

driver.execute_script('arguments[0].innerTEXT = "spain";', origin)
driver.execute_script('arguments[0].innerTEXT = "France";', destination)
driver.execute_script('arguments[0].innerTEXT = "International";', trip_type)
driver.execute_script('arguments[0].innerTEXT = "Air";', demand_category)
'''

