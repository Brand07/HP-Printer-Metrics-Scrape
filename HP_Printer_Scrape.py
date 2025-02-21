import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime


"""
Printer dictionary (NAME:PRINTER IP)
Change or add the printer names/IP addresses for your environment.
"""
printers = {
    "STN2": "10.10.113.150",
    "LTL1": "10.10.113.212",
    "CLB1": "10.10.113.180",
    "SPS2": "10.10.113.215",
    "STN1": "10.10.113.213",
    "SPS1": "10.10.113.214"
}

# Tags we're looking for
page_count_tag = 'UsagePage.ImpressionsByMediaSizeTable.Print.Letter.Total'
toner_tag = 'SupplyGauge0'
maintenance_tag = 'SupplyGauge1'

def get_data(url, tag, element='td'):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        data = soup.find(element, {'id': tag}).text.replace(',', '')
        return data
    except requests.exceptions.RequestException:
        print(f"Issue getting data from {url}...")
        return None

# Collect and store the data in an empty list
data =[]
for name, ip, in printers.items():
    page_count_url = f'http://{ip}/hp/device/InternalPages/Index?id=UsagePage'
    device_status_url = f'http://{ip}/hp/device/DeviceStatus/Index'

    page_count = get_data(page_count_url, page_count_tag)
    toner_level = get_data(device_status_url, toner_tag, 'span')
    maint_level = get_data(device_status_url, maintenance_tag, 'span')

    #Append the data to the empty list
    data.append({
        'Printer': name,
        'Page Count': page_count,
        'Toner Level': toner_level,
        'Maintenance Level': maint_level
    })

# Get the current date
now = datetime.now()
date_string = now.strftime("%m-%d-%Y")

# Create the pandas df for the excel expor
df = pd.DataFrame(data)
df['Date'] = date_string

# Try to load the previous file to append new data
try:
    prev_df = pd.read_excel('Printer_Metrics.xlsx')
except FileNotFoundError:
    prev_df = pd.DataFrame()

# Append the new data to the existing file
df = pd.concat([prev_df, df], ignore_index=True)

# Write the data to the excel file
df.to_excel('Printer_Metrics.xlsx', index=False)




