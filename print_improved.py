import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime

stn2_tdate_url = "http://10.10.113.150/hp/device/InternalPages/Index?id=SuppliesStatus"
ltl1_tdate_url = "http://10.10.113.212/hp/device/InternalPages/Index?id=SuppliesStatus"
clb1_tdate_url = "http://10.10.113.180/hp/device/InternalPages/Index?id=SuppliesStatus"
sps2_tdate_url = "http://10.10.113.215/hp/device/InternalPages/Index?id=SuppliesStatus"
stn1_tdate_url = "http://10.10.113.213/hp/device/InternalPages/Index?id=SuppliesStatus"
sps1_tdate_url = "http://10.10.113.214/hp/device/InternalPages/Index?id=SuppliesStatus"

# URLs for the Usage Metrics
stn2_pcount_url = 'http://10.10.113.150/hp/device/InternalPages/Index?id=UsagePage'
ltl1_pcount_url = 'http://10.10.113.212/hp/device/InternalPages/Index?id=UsagePage'
clb1_pcount_url = 'http://10.10.113.180/hp/device/InternalPages/Index?id=UsagePage'
sps2_pcount_url = 'http://10.10.113.215/hp/device/InternalPages/Index?id=UsagePage'
stn1_pcount_url = 'http://10.10.113.213/hp/device/InternalPages/Index?id=UsagePage'
sps1_pcount_url = 'http://10.10.113.214/hp/device/InternalPages/Index?id=UsagePage'
# URLs for the consumable levels
stn2_device_status = 'http://10.10.113.150/hp/device/DeviceStatus/Index'
ltl1_device_status = 'http://10.10.113.212/hp/device/DeviceStatus/Index'
clb1_device_status = 'http://10.10.113.180/hp/device/DeviceStatus/Index'
sps2_device_status = 'http://10.10.113.215/hp/device/DeviceStatus/Index'
stn1_device_status = 'http://10.10.113.213/hp/device/DeviceStatus/Index'
sps1_device_status = 'http://10.10.113.214/hp/device/DeviceStatus/Index'

stn2_pprinted_url = "http://10.10.113.150/hp/device/InternalPages/Index?id=SuppliesStatus"
ltl1_pprinted_url = "http://10.10.113.212/hp/device/InternalPages/Index?id=SuppliesStatus"
clb1_pprinted_url = "http://10.10.113.180/hp/device/InternalPages/Index?id=SuppliesStatus"
sps2_pprinted_url = "http://10.10.113.215/hp/device/InternalPages/Index?id=SuppliesStatus"
stn1_pprinted_url = "http://10.10.113.213/hp/device/InternalPages/Index?id=SuppliesStatus"
sps1_pprinted_url = "http://10.10.113.214/hp/device/InternalPages/Index?id=SuppliesStatus"

# HTML Element Tags
page_count_tag = 'UsagePage.ImpressionsByMediaSizeTable.Print.Letter.Total'
toner_tag = 'SupplyGauge0'
maintenance_tag = 'SupplyGauge1'
toner_installed = "BlackCartridge1-FirstInstallDate"
pages_printed_tag = "BlackCartridge1-PagesPrintedWithSupply"

"""
This function gets allows you to pass the url of the page that 
contains the printer's page count
"""
def get_page_count(url):
    print(f"Getting the page count for {url}")
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    data = soup.find('td', {'id': page_count_tag}).text.replace(',','')
    return data

def get_toner_level(url):
    print(f"Getting the toner level for {url}")
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    data = soup.find('span', {'id': toner_tag}).text.replace(',','')
    return data

def get_maint_level(url):
    print(f"Getting the maintenance kit level for {url}")
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    data = soup.find('span', {'id': maintenance_tag}).text.replace(',','')
    if data == "10%":
        print(f"Consider changing the maintenance kit on the printer: {url}")

    return data

def get_toner_date(date_url):
    print(f"Getting the toner install date for {date_url}..")
    try:
        response = requests.get(date_url)
        soup = BeautifulSoup(response.text, 'html.parser')
        data = soup.find('strong', {'id': toner_installed}).text
        date_object = datetime.strptime(data, "%Y%m%d")
        formatted_date = date_object.strftime("%m-%d-%Y")
        return(formatted_date)
    except Exception as e:
        return("Error", e)

def get_supply_pages(date_url):
    print(f"Getting the pages printed with current installed toner on {date_url}..")
    try:
        response = requests.get(date_url)
        soup = BeautifulSoup(response.text, 'html.parser')
        data = soup.find('strong', {'id': pages_printed_tag}).text
        return(data)
    except Exception as e:
        return("Error", e)


# Page Counts for the printers
stn2_page_count = get_page_count(stn2_pcount_url)
ltl1_page_count = get_page_count(ltl1_pcount_url)
clb1_page_count = get_page_count(clb1_pcount_url)
sps2_page_count = get_page_count(sps2_pcount_url)
stn1_page_count = get_page_count(stn1_pcount_url)
sps1_page_count = get_page_count(sps1_pcount_url)
# Toner levels for each printer.
stn2_toner_level = get_toner_level(stn2_device_status)
ltl1_toner_level = get_toner_level(ltl1_device_status)
clb1_toner_level = get_toner_level(clb1_device_status)
sps2_toner_level = get_toner_level(sps2_device_status)
stn1_toner_level = get_toner_level(stn1_device_status)
sps1_toner_level = get_toner_level(sps1_device_status)
# Maintenance Kit Levels for Each Printer
stn2_maint_level = get_maint_level(stn2_device_status)
ltl1_maint_level = get_maint_level(ltl1_device_status)
clb1_maint_level = get_maint_level(clb1_device_status)
sps2_maint_level = get_maint_level(sps2_device_status)
stn1_maint_level = get_maint_level(stn1_device_status)
sps1_maint_level = get_maint_level(sps1_device_status)

#Toner install date for each printer
stn2_toner_date = get_toner_date(stn2_tdate_url)
ltl1_toner_date = get_toner_date(ltl1_tdate_url)
clb1_toner_date = get_toner_date(clb1_tdate_url)
sps2_toner_date = get_toner_date(sps2_tdate_url)
stn1_toner_date = get_toner_date(stn1_tdate_url)
sps1_toner_date = get_toner_date(sps1_tdate_url)

# Pages printed with current installed toner for each printer
stn2_pocs = get_supply_pages(stn2_pprinted_url)
ltl1_pocs = get_supply_pages(ltl1_pprinted_url)
clb1_pocs = get_supply_pages(clb1_pprinted_url)
sps2_pocs = get_supply_pages(sps2_pprinted_url)
stn1_pocs = get_supply_pages(stn1_pprinted_url)
sps1_pocs = get_supply_pages(sps1_pprinted_url)


# Get the Curent Date
now = datetime.now()
date_string = now.strftime("%m-%d-%Y")

# Create the Pandas Dataframe for the Excel export.
df = pd.DataFrame({
    'Date' : [date_string]*6,
    'Printer': ['STN2', 'LTL1', 'CLB1', 'SPS2', 'STN1', 'SPS1'],
    'Page Count': [stn2_page_count, ltl1_page_count, clb1_page_count, sps2_page_count, stn1_page_count, sps1_page_count],
    'Toner Level': [stn2_toner_level, ltl1_toner_level, clb1_toner_level, sps2_toner_level, stn1_toner_level, sps1_toner_level],
    'Maintenance Level': [stn2_maint_level, ltl1_maint_level, clb1_maint_level, sps2_maint_level, stn1_maint_level, sps1_maint_level],
    'Toner Install Date': [stn2_toner_date,ltl1_toner_date,clb1_toner_date,sps2_toner_date,stn1_toner_date,sps1_toner_date],
    'Current Toner Pages': [stn2_pocs,ltl1_pocs,clb1_pocs,sps2_pocs,stn1_pocs,sps1_pocs]
    
})

"""
Attemps to load the previous data from Excel if it exists.
If not, it will create a new file in the running directory.
"""
try:
    prev_df =pd.read_excel('AllPrinters.xlsx')
except FileNotFoundError:
    prev_df =pd.DataFrame()

#Append new data to the previous data
df = pd.concat([prev_df, df], ignore_index=True)

#Write to Excel 
df.to_excel('AllPrinters.xlsx', index=False)
print("Completed!")
print("File Saved!")
