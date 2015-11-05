import sys
import logging
import requests
import json
from openpyxl import load_workbook
from urllib2 import urlopen
from requests.exceptions import RequestException
from bs4 import BeautifulSoup

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

URL = "https://api.superfastmaps.co.uk/1.0/checker/check.ajax.php?map=fastershire&lang=&input=01531660552&address=&extra="
USERAGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.71 Safari/537.36"
"""
This is the header for the post, dont touch the header otherwise the post will
not get the correct request
"""
POSTHEADER = {
    'Origin':'https://api.superfastmaps.co.uk',
    'User-Agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.71 Safari/537.36',
    'Referer':'https://api.superfastmaps.co.uk/1.0/map/?map=fastershire',
}

def setupfaillogger():
    logger = logging.getLogger("scrapelog")
    logger.setLevel(logging.DEBUG)
    #create file handler and set level to debug
    fh = logging.FileHandler("spam.log")
    fh.setLevel(logging.DEBUG)
    #create stream handler
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    #create formatter
    formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
    #add formatter to ch and fh
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)
    #add ch and fh to logger
    logger.addHandler(fh)
    logger.addHandler(ch)

def requestcheck(number):
    l = logging.getLogger("scrapelog")
    form_data = {'input':number}
    try:
        r = requests.post(URL, params=form_data, headers=POSTHEADER)
    except RequestException as e:
        print bcolors.FAIL + "ERROR: The Request Failed, due to connection problems!"
        print "INFO: Contact developer: sreungbrmzra@gmail.com"
        l.exception("Connection Error")
        sys.exit()

    data = r.json()
    return data 

def main():
    setupfaillogger()
    xls_file = raw_input("Enter the name of .xlsx file: ")
    wb = load_workbook(xls_file)
    ws = wb['Sheet1']
    
    l = logging.getLogger("scrapelog")
    for row in ws.rows[1:]:
        l.info(bcolors.OKBLUE + "Requesting Number: %s" %row[0].value)
        data = requestcheck(str(row[0].value))
        messagetype = str(data['messagetype'])
        cabinet = str(data['cabinet'])
        speed = str(data['speed'])
        exchangename = str(data['exchange_eng'])
        if messagetype == "acceptingorders":
            l.info(bcolors.OKGREEN + 'Request Received for Number %s: %s' %(row[0].value, messagetype))
            row[1].value = 0
            row[2].value = cabinet
            row[3].value = exchangename 
        elif messagetype == "numbernotrecognised":
            l.info(bcolors.OKGREEN + 'Request Received for Number %s: %s' %(row[0].value, messagetype))
            row[1].value = "x"
        else:
            l.info(bcolors.OKGREEN + 'Request Received for Number %s: %s' %(row[0].value, messagetype))
            row[1].value = speed
            row[2].value = cabinet
            row[3].value = exchangename
    wb.save(xls_file)

if __name__ == "__main__":
    main()
