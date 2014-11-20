panorama-rules-to-excel
=======================

Create an Excel Spreadsheet from your firewall rules in Palo Alto Networks Panorama

Usage
=====
    usage: pan_to_excel.py [-h] -k APIKEY -f FIREWALL -p PANORAMA
    
    Convert Palo Alto Network Firewall rules from Panorama to Microsoft Excel.
      
      optional arguments:
        -h, --help            show this help message and exit
        -k APIKEY, --apikey APIKEY
                              PAN API Token Key
        -f FIREWALL, --firewall FIREWALL
                              Firewall Name
        -p PANORAMA, --panorama PANORAMA
                              Panorama Managment URL
      
      i.e. pan_to_excel.py --apikey "23j4kl2j34klj2kl4hf5yf" --firewall "Prod
      firewall 1" --panorama "https://panorama.somewhere.com

Dependencies
============
Python with the following modules:
.*requests
.*ElementTree
.*xlsxwriter
.*argparse
