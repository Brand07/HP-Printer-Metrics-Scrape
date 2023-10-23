# HP Printer Metrics Scraper


# What is it?
This is a simple Python script that scrapes the internal web pages of HP Printers for data such as page count, toner level, and maintenance kit level. After the script pulls all the data, it exports it to an .xlsx file in the running directory. 
If the .xlsx file already exists, it will append the data. If not, it will create a new file.

**This script has parameters customized for my own environment, but things like the printer names and the quanity of printers can be changed.**


Example Output:

![Screenshot 2023-10-23 114533](https://github.com/Brand07/HP-Printer-Metrics-Scrape/assets/81128304/580b4bf4-c41e-464d-a54d-453c353698ac)


# Can it scrape any other data from the printer?
Sure. You would need to find the specific HMTL tag of the information you want to scrape and pass it off into the script.

# Any other requirements?
See requirements.txt for the libraries you will need for this script.
They are:
+ BeautifulSoup
+ Pandas
+ Datetime (default Python library - no need to install)
+ Requests (default Python library - no need to install)


### Other Requirements
This script requires that you disable *Encrypt All Web Communication (not including IPP)* on the printer. This can be done from the internal webpage under the Networking Tab >> Mgmt. Protocals.
This allows the script to request the webpage, otherwise the script would fail to reach the webpage.

Do this at your own risk as you are disabling a security feature on your printer
![254900750-857f363f-b7dd-4d55-934e-ab152758a916](https://github.com/Brand07/HP-Printer-Metrics-Scrape/assets/81128304/b7b84f02-1c4e-43da-9cf7-1059ed8ba43b)

