#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb  9 20:57:29 2020

@author: Mishaun_Bhakta
"""

import pandas as pd
import os

#sale parameters
state = "New Mexico"
date = "February 6, 2020"


#Navigate to energynet/govt sale and get sale page
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup


url = 'https://www.energynet.com/page/Government_Sales_Results'

filepath = abspath = os.path.dirname(__file__)

driver = webdriver.Chrome(filepath + "/chromedriver")
driver.implicitly_wait(30)
driver.get(url)

link = driver.find_element_by_xpath('//*[@id="main_page"]/div/div[2]/div/div/div[1]/div[2]/a').click()



#open workbook and store serial numbers

df = pd.read_excel('BLM NM 2-6-20 Sale Notes.xlsm', header = 6, usecols = "B:T")
serials = df["Serial numbers"].iloc[0:-1]

#use pdf reader to fill in form
# conda install -c conda-forge pdfrw

import pdfrw