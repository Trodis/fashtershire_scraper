Fastershire Scraper

This is just a simple Scraper for the online service Fastershire. 
The tool is useful if you provide a .xlsx file with several numbers you want to check the broadband availability.
The script iterates over all phonenumbers and retrieves for each Number the result.

So if you have a list of phonenumbers you want to check, you can take .xlsx example and populate the Phone Number fields
with the numbers u like. The Script will check them for you. 


Python modules needed:

sys
logging
requests
json
from openpyxl -> load_workbook
from urllib2 -> urlopen
from requests.exceptions -> RequestException
from bs4 -> BeautifulSoup
