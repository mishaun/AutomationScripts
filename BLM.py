#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb  9 20:57:29 2020

@author: Mishaun_Bhakta
"""
import os
import re

#sale parameters
state = "New Mexico"
stinitials = "NM"
date = "Feb 6, 2020"
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

    return acres, desc, county, serialnums

acres, descriptions, counties, serials = webscrape_presale(salehtml)

#Open sale template and update insert information from webscrape
import openpyxl


def fillexcel():
    '''
    This function will take scraped (global) values for lots and insert into sale spreadsheet
    '''
    
    wb = openpyxl.load_workbook("BLM Sale Notes Template.xlsm", keep_vba = True)
    sheet = wb.active
    
    sheet["B6"] = "BLM {} {} Sale Notes".format(stinitials, date)
    #inserting values from webscrape into spreadsheet
    for i in range(0,len(serials)):
        sheet.cell(row = 8+i, column = 2, value = serials[i])
        sheet.cell(row = 8+i, column = 6, value = acres[i])
        sheet.cell(row = 8+i, column = 7, value = counties[i])
        sheet.cell(row = 8+i, column = 8, value = descriptions[i])
    
    wb.save("BLM {} {} Sale Notes.xlsm".format(stinitials, date))
    wb.close()

    
bidtags = salehtml.find_all("td", "lot-bid")
ourwinnings = {}

for i in range (0,len(bidtags)):
    textCont = bidtags[i].text
    
    #extracting bidder number from lot bid tags - try statement prevents break if no bids were received
    try:
        winBidder = re.findall('#\d+', textCont)[0].replace("#",'')
    except:
        print("no bids received for parcel: " + serials[i])
        winBidder = type(None)
        pass
    
    #if we won the bid, then capture the winning bid amount
    if bidder == winBidder:
        winAmount = re.findall('\$\d+', textCont)[0]
        winAmount = winAmount.replace('$','')
    
        ourwinnings[i] = winAmount

#### insert our winnings into sale spreadsheet
wb = openpyxl.load_workbook("BLM {} {} Sale Notes.xlsm".format(stinitials, date), keep_vba = True)
sheet = wb.active

for i in range(0,len(ourwinnings)):
    sheet.cell(row = 8 + list(ourwinnings.keys())[i], column = 17, value = ourwinnings[list(ourwinnings.keys())[i]])
    sheet.cell(row = 8 + list(ourwinnings.keys())[i], column = 16, value = 'Y')

wb.save("BLM {} {} Sale Notes.xlsm".format(stinitials, date))
wb.close()

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