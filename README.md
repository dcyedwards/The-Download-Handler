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

Requirements
------------

1. Mozilla Firefox (version 46)
   https://ftp.mozilla.org/pub/firefox/releases/46.0/
   or
   https://ftp.mozilla.org/pub/firefox/releases/46.0.1/
   or
   http://filehippo.com/download_firefox/67599/
   
2. Python 3
   https://www.python.org/
   
3. MimeTypes.rdf file located in repository.
   This file should overwrite the existing file (after backing up the original) in the .default folder in the following Mozilla Firefox      directory or similar:
   
   C:\Users\Computer-Username\AppData\Roaming\Mozilla\Firefox\Profiles\****.default

   
