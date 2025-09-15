
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
import time

service = ChromeService(executable_path=r"C:\Windows\System32\chromedriver.exe")
opts = webdriver.ChromeOptions()
opts.add_experimental_option("detach", True)  # keep window open
driver = webdriver.Chrome(service=service, options=opts)
driver.get("https://www.google.com")
time.sleep(10)