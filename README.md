# The-Download-Handler
This script which I wrote during my time at Wood Mackenzie as a Data Analyst, is written for the sole purpose of aiding the analysts in their data analysis. The script automatically handles the time consuming task(s) of going to different web pages, copying, selecting and then saving the data to various folders on the network. This version - which has been adapted for use outwith of the company - largely fulfills the download routine but has no access to the company network drives. 

As such, it is meant purely as a demonstration of how Python employs some specific libraries and modules to interact with websites. These modules are namely: 
- Selenium web driver
- Requests
- BeautifulSoup4
- Scrapy (which I don't actually use in this script version). 

Naturally many other modules are employed in tandem with the aforementioned modules to achieve the desired objective: 
- OS module
- Win32com
- Pythoncom
- Shutil
