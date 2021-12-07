from simple_salesforce import Salesforce, SalesforceLogin, SFType
import requests
import json
import pandas as pd
from io import StringIO

sf = Salesforce(username='dominic.toth@boweryvaluation.com',password='DOMcathy80!@',security_token='977pcdtBN8m5VTRz1FfSzi4hG')
#login for Mike
#sf = Salesforce(username='mike.leahy@boweryvaluation.com',password='',security_token='')

sf = Salesforce(instance_url='https://boweryvaluation.lightning.force.com', session_id='00O4X000009MgMOUA0')

#get https://boweryvaluation.my.salesforce.com/services/data/vXX.X/resource/

sf.Contact.create({'LastName':'Smith','Email':'example@example.com'})
