#Salesforce API used to systematically identify new jobs, their area and hood, and create an area and report
#https://ryanwingate.com/salesforce/salesforce-and-python/programmatically-obtaining-all-fields/

from unicodedata import name
from numpy import where
from numpy.lib.function_base import select
from simple_salesforce import Salesforce, SalesforceLogin, SFType
import requests
import json
import pandas as pd
from io import StringIO

sf = Salesforce (
    username = 'dominic.toth@boweryvaluation.com',
    password = 'DOMcathy80!@',
    security_token = '977pcdtBN8m5VTRz1FfSzi4hG',
    client_id = 'boweryvaluation',
    #sandbox = True
)

desc = sf.describe()
print(desc)

desc = sf.describe()
objects = []
obj_labels = [obj['label']  for obj in desc['sobjects']]
obj_names  = [obj['name']   for obj in desc['sobjects']]
obj_custom = [obj['custom'] for obj in desc['sobjects']]
for label, name, custom in zip(obj_labels, obj_names, obj_custom):
    objects.append((label, name, custom))
objects = pd.DataFrame(objects,
                       columns = ['label','name','custom'])
print(str(objects.shape[0]) + ' objects')
objects.head(10)

#this example uses the Job object
desc = sf.Job_c.describe()
fields = []
field_labels = [field['label'] for field in desc['fields']]
field_names =  [field['name']  for field in desc['fields']]
field_types =  [field['type']  for field in desc['fields']]
for label, name, dtype in zip(field_labels, field_names, field_types):
    fields.append((label, name, dtype))
fields = pd.DataFrame(fields,
                      columns = ['label','name','type'])
print(str(fields.shape[0]) + ' fields')
fields.head(10)



#field_names = [field['name'] for field in desc['fields']]
#sf_data = sf.query_all("SELECT Address, Market_Research_Due_Date__c, Assigned_To_c")
#sf_df = pd.DataFrame(sf_data['records']).drop(columns='attributes')

#print(sf_df)

#field_names = [field['name'] for field in desc['fields']]
#soql = "SELECT {} FROM Account".format(','.join(field_names))
#results = sf.query_all(soql) #normal download
#results = sf.bulk.Account.query(soql) #bulk download



#login for Mike
#sf = Salesforce(username='mike.leahy@boweryvaluation.com',password='',security_token='')
#sf = Salesforce(instance_url='https://boweryvaluation.lightning.force.com', session_id='00O4X000009MgMOUA0')
#get https://boweryvaluation.my.salesforce.com/services/data/vXX.X/resource/
#sf.Contact.create({'LastName':'Smith','Email':'example@example.com'})
#select Id, name
#from account
#where name = 'Dominic'
