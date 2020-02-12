#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb  9 20:57:29 2020

@author: Mishaun_Bhakta
"""

import pandas as pd
import os
import re

#sale parameters
state = "New Mexico"
date = "February 6, 2020"
bidder = '3'


#Navigate to energynet/govt sale and get sale page
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup

url = 'https://www.energynet.com/page/Government_Sales_Results'

filepath = abspath = os.path.dirname(__file__)

#driver will be used based on operating system - windows or mac
try:
    driver = webdriver.Chrome(filepath + "/chromedriver.exe")
except:
    driver = webdriver.Chrome(filepath + "/chromedriver")

driver.implicitly_wait(30)
driver.get(url)

link = driver.find_element_by_xpath('//*[@id="main_page"]/div/div[2]/div/div/div[1]/div[2]/a').click()
salehtml = BeautifulSoup(driver.page_source, "html.parser")

def webscrape_presale(parsepage):
    '''This function will take a page and scrape its data for sale lot information
    '''
    #webscrape sale page
    
    serialnums = parsepage.find_all("span", "lot-name")
    serialnums = [i.text for i in serialnums]
    
    legalinfo = parsepage.find_all("td", "lot-legal")
    
    acres = []
    desc = []
    county = []
    
    for item in legalinfo:
        county.append(item.contents[0].text)
        desc.append(item.contents[1].text)
        #getting acres by splitting at : and blankspace to get string of numerical value - taking out a comma if above 1000 in order to convert to float
        acres.append(float(re.split(":\W",item.contents[2])[1].replace(',','')))

    return acres, desc, county


acres, descriptions, counties = webscrape_presale(salehtml)

#open workbook and store serial numbers
df = pd.read_excel('BLM NM 2-6-20 Sale Notes.xlsm', header = 6, usecols = "B:T")
serials = df["Serial numbers"].iloc[0:-1]


def scrape_lotswon(parsepage, bidderNum):
    
    bidstatus = parsepage.find_all("td", "lot-bid")
    ##need to finish function####
    return bidstatus

bidderinfo = scrape_lotswon(salehtml, bidder)




#use pdf reader to fill in form
# conda install -c conda-forge pdfrw
import pdfrw

fields = {
        "State": "",
        "Date of Sale": "",
        'Check Box for Oil and Gas' : "",
        "Oil and Gas/Parcel No" : "",
        "TOTAL BID FOR Oil and Gas Lease" : "",
        "PAYMENT SUBMITTED WITH BID for Oil and Gas" : "",
        "Print or Type Name of Lessee" : "",
        "Address of Lessee": "",
        "City" : "",
        "State_2": "",
        "Zip Code" : ""
        }