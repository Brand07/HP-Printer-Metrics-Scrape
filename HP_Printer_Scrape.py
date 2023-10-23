import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime


#Printers
printer_1 = "STN2"
printer_2 = "LTL1"
printer_3 = "CLB1"
printer_4 = "SPS2"
printer_5 = "STN1"
printer_6 = "SPS1"
    
#URLs for the Usage Metrics
stn2_pcount_url = 'http://10.10.113.150/hp/device/InternalPages/Index?id=UsagePage'
ltl1_pcount_url = 'http://10.10.113.212/hp/device/InternalPages/Index?id=UsagePage'
clb1_pcount_url = 'http://10.10.113.180/hp/device/InternalPages/Index?id=UsagePage'
sps2_pcount_url = 'http://10.10.113.215/hp/device/InternalPages/Index?id=UsagePage'
stn1_pcount_url = 'http://10.10.113.213/hp/device/InternalPages/Index?id=UsagePage'
sps1_pcount_url = 'http://10.10.113.214/hp/device/InternalPages/Index?id=UsagePage'
#URLs for the consumable levels
stn2_device_status = 'http://10.10.113.150/hp/device/DeviceStatus/Index'
ltl1_device_status = 'http://10.10.113.212/hp/device/DeviceStatus/Index'
clb1_device_status = 'http://10.10.113.180/hp/device/DeviceStatus/Index'
sps2_device_status = 'http://10.10.113.215/hp/device/DeviceStatus/Index'
stn1_device_status = 'http://10.10.113.213/hp/device/DeviceStatus/Index'
sps1_device_status = 'http://10.10.113.214/hp/device/DeviceStatus/Index'

page_count_tag = 'UsagePage.ImpressionsByMediaSizeTable.Print.Letter.Total'
toner_tag = 'SupplyGauge0'
maintenance_tag = 'SupplyGauge1'
"""
This function gets allows you to pass the url of the page that 
contains the printer's page count
"""
def get_page_count(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        data = soup.find('td', {'id': page_count_tag}).text
        return data
    except requests.exceptions.RequestException:
        print("Issue gathering Page Count...")
        return None


def get_toner_level(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        data = soup.find('span', {'id': toner_tag}).text.replace(',','')
        return data
    except requests.exceptions.RequestException:
        print("Issue gathering Toner Level...")
        return None

def get_maint_level(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        data = soup.find('span', {'id': maintenance_tag}).text.replace(',','')
        return data
    except requests.exceptions.RequestException:
        print("Issue gathering Maintenance Level...")
        return None


#Page Counts for the printers   
stn2_page_count = get_page_count(stn2_pcount_url)
ltl1_page_count = get_page_count(ltl1_pcount_url)
clb1_page_count = get_page_count(clb1_pcount_url)
sps2_page_count = get_page_count(sps2_pcount_url)
stn1_page_count = get_page_count(stn1_pcount_url)
sps1_page_count = get_page_count(sps1_pcount_url)
#Toner levlels for each printer.
stn2_toner_level = get_toner_level(stn2_device_status)
ltl1_toner_level = get_toner_level(ltl1_device_status)
clb1_toner_level = get_toner_level(clb1_device_status)
sps2_toner_level = get_toner_level(sps2_device_status)
stn1_toner_level = get_toner_level(stn1_device_status)
sps1_toner_level = get_toner_level(sps1_device_status)
#Maintenance Kit Levels for Each Printer
stn2_maint_level = get_maint_level(stn2_device_status)
ltl1_maint_level = get_maint_level(ltl1_device_status)
clb1_maint_level = get_maint_level(clb1_device_status)
sps2_maint_level = get_maint_level(sps2_device_status)
stn1_maint_level = get_maint_level(stn1_device_status)
sps1_maint_level = get_maint_level(sps1_device_status)

#Get the Curent Date
now = datetime.now()
date_string = now.strftime("%m-%d-%Y")

#Create the Pandas Dataframe for the Excel export.
df = pd.DataFrame({
    'Date' : [date_string]*6,
    'Printer': ['STN2', 'LTL1', 'CLB1', 'SPS2', 'STN1', 'SPS1'],
    'Page Count': [stn2_page_count, ltl1_page_count, clb1_page_count, sps2_page_count, stn1_page_count, sps1_page_count],
    'Toner Level': [stn2_toner_level, ltl1_toner_level, clb1_toner_level, sps2_toner_level, stn1_toner_level, sps1_toner_level],
    'Maintenance Level': [stn2_maint_level, ltl1_maint_level, clb1_maint_level, sps2_maint_level, stn1_maint_level, sps1_maint_level]
    
})

"""
Attemps to load the previous data from Excel if it exists.
If not, it will create a new file in the running directory.
"""
try:
    prev_df =pd.read_excel('Printer_Metrics.xlsx')
except FileNotFoundError:
    prev_df =pd.DataFrame()

#Append new data to the previous data
df = pd.concat([prev_df, df], ignore_index=True)

#Write to Excel 
df.to_excel('Printer_Metrics.xlsx', index=False)


