#! python3clear
from selenium import webdriver
from threading import Thread
import time
import datetime
import timeit
import requests
import urllib.request as ur
from openpyxl import Workbook
from openpyxl.styles import Side, Alignment, Protection, Font
from bs4 import BeautifulSoup

startTime = timeit.default_timer()
yahooStatsUrl = "http://finance.yahoo.com/quote/aapl/key-statistics?p=aapl"
driver = webdriver.Chrome("/Users/Sparky/Downloads/chromedriver")
time1 = timeit.default_timer()
driver.get(yahooStatsUrl)
time2 = timeit.default_timer()
time.sleep(5)
yahooStatsPage = driver.page_source
yahooStatsSoup = BeautifulSoup(yahooStatsPage, "html.parser")
time3 = timeit.default_timer()
driver.quit()

time4 = timeit.default_timer()
val = yahooStatsSoup.find('td' ,{'class' : "Fz(s) Fw(500) Ta(end)"})
#searching for the text, not by the tags..?
numpty50 = yahooStatsSoup.find_all('td', {'class' : "Fz(s) Fw(500) Ta(end)"})
print(numpty50[36]) #This is the correct one. 37 should be the 200day moving average
# numpty200 = yahooStatsSoup.find(string = "200-day Moving Average")
stopTime = timeit.default_timer()

print("Total time is " + str(stopTime - startTime))
print("startup time is " + str(time1 - startTime))
print("getting time is " + str(time2 - time1))
print("Rendering time is " + str(time3 - time2 -5))
print("Find All time is " + str(stopTime - time4))
