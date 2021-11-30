from simple_salesforce import Salesforce
import requests
import pandas as pd
from io import StringIO

sf = Salesforce(username='dominic.toth@boweryvaluation.com',password='DOMcathy80!@',security_token='977pcdtBN8m5VTRz1FfSzi4hG')
#login for Mike
#sf = Salesforce(username='mike.leahy@boweryvaluation.com',password='',security_token='')