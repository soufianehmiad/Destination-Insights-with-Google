from selenium import webdriver
from selenium.webdriver.common.by import By
import time


driver = webdriver.Firefox(executable_path="webrivers\\geckodriver.exe")
driver.get("https://destinationinsights.withgoogle.com/")



assert "Destination Insights with Google" in driver.title



# ACCEPT COOKIES

cookies = driver.find_element(By.XPATH, '/html/body/div[1]/div/span[2]/a[2]')

cookies.click()



# GET ELEMENTS

time.sleep(10)

origin = driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div/div/div[1]/div[1]/md-content[1]/md-select/md-select-value/span[1]')

destination = driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div/div/div[1]/div[2]/md-content[1]/md-select/md-select-value/span[1]')

trip_type = driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div/div/div[2]/md-content[1]/md-select/md-select-value/span[1]')

demand_category = driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div/div/div[2]/md-content[2]/md-select/md-select-value/span[1]')



# SET FILTER

driver.execute_script('arguments[0].innerHTML = "Worldwide";', origin)

driver.execute_script('arguments[0].innerHTML = "Morocco";', destination)

driver.execute_script('arguments[0].innerHTML = "International";', trip_type)

driver.execute_script('arguments[0].innerHTML = "Air";', demand_category)



# SUBMIT FILTER

submit_btn = driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div/div/div[2]/div[2]/button')

submit_btn.click()