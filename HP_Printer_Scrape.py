import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime

# Printer details
printers = {
    "STN2": "10.10.113.150",
    "LTL1": "10.10.113.212",
    "CLB1": "10.10.113.180",
    "SPS2": "10.10.113.215",
    "STN1": "10.10.113.213",
    # "SPS1": "10.10.113.214"
}

# Tags for scraping
page_count_tag = "UsagePage.ImpressionsByMediaSizeTable.Print.Letter.Total"
toner_tag = "SupplyGauge0"
maintenance_tag = "SupplyGauge1"


def get_data(url, tags, element="td"):
    try:
        print(f"Getting data for {url}")
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")
        data = {
            tag: soup.find(element, {"id": tag}).text.replace(",", "") for tag in tags
        }
        return data
    except requests.exceptions.RequestException:
        print(f"Issue gathering data from {url}...")
        return {tag: None for tag in tags}


# Collect data for each printer
data = []
for name, ip in printers.items():
    page_count_url = f"http://{ip}/hp/device/InternalPages/Index?id=UsagePage"
    device_status_url = f"http://{ip}/hp/device/DeviceStatus/Index"

    page_count = get_data(page_count_url, [page_count_tag])[page_count_tag]
    device_status = get_data(device_status_url, [toner_tag, maintenance_tag], "span")

    data.append(
        {
            "Printer": name,
            "Page Count": page_count,
            "Toner Level": device_status[toner_tag],
            "Maintenance Level": device_status[maintenance_tag],
        }
    )

# Get the current date
now = datetime.now()
date_string = now.strftime("%m-%d-%Y")

# Create the Pandas DataFrame for the Excel export
df = pd.DataFrame(data)
df["Date"] = date_string

# Attempt to load the previous data from Excel if it exists
try:
    print("Looking for existing .xlsx file")
    prev_df = pd.read_excel("Printer_Metrics.xlsx")
except FileNotFoundError:
    prev_df = pd.DataFrame()

# Append new data to the previous data
print("Appending the data")
df = pd.concat([prev_df, df], ignore_index=True)

# Write to Excel
df.to_excel("Printer_Metrics.xlsx", index=False)
print("Data appended!")
