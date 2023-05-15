import requests
import json
import pandas as pd

url = 'https://www.nseindia.com/api/option-chain-indices?symbol=BANKNIFTY'

headers={'user-agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'}
response = requests.get(url, headers=headers, timeout=10)

response_text=response.text
json_object = json.loads(response_text)

#print(json_object)

with open("OC.json", "w") as outfile:
    outfile.write(response_text)

# Below 'data' os a list 
data = json_object['records']['data']
#print(type(data))

# prepare list of expiry date from json_object. This list will be used to fetah OC date based on expiry date
e_date=json_object['records']['expiryDates']

oc_data = {}
for ed in e_date:
    oc_data[ed]={"CE":[], "PE":[]}
    for di in range(len(data)):
        
        if data[di]['expiryDate'] == ed:
            if 'CE' in data[di].keys() and data[di]['CE']['expiryDate'] == ed:
                
                oc_data[ed]["CE"].append(data[di]['CE'])
            else:
                oc_data[ed]["CE"].append('-')

            if 'PE' in data[di].keys() and data[di]['PE']['expiryDate'] == ed:
                oc_data[ed]["PE"].append(data[di]['PE'])
                
            else:
                oc_data[ed]["PE"].append('-')

# This loop is to delete extra keys from oc_data
for k in oc_data.keys():
    
    for i in range(len(oc_data[k]["CE"])):
        
        if oc_data[k]["CE"][i] != '-':
        
            del oc_data[k]["CE"][i]["expiryDate"]
            del oc_data[k]["CE"][i]["underlying"]
            del oc_data[k]["CE"][i]["identifier"]

        if oc_data[k]["PE"][i] != '-':

            del oc_data[k]["PE"][i]["expiryDate"]
            del oc_data[k]["PE"][i]["underlying"]
            del oc_data[k]["PE"][i]["identifier"]

expiry_date = e_date[0]

keys_to_exclue = ['pchangeinOpenInterest',
 'totalTradedVolume',
 'change',
 'pChange',
 'totalBuyQuantity',
 'totalSellQuantity',
 'bidQty',
 'bidprice',
 'askQty',
 'askPrice']

oc_data_dt=oc_data[expiry_date]

CE=list(oc_data_dt['CE'])
PE=list(oc_data_dt['PE'])

for i in range(len(CE)):
    if CE[i] != '-':
        for key in keys_to_exclue:
            del CE[i][key]
    if PE[i] != '-':
        for key in keys_to_exclue:
            del PE[i][key]

l_OC = []

for i in range(len(CE)):
    l_CE=[]
    l_PE=[]
    
    if CE[i] != '-':
        sp = CE[i]['strikePrice']
        l_CE = [
                    CE[i]['openInterest'],
                    CE[i]['changeinOpenInterest'],
                    CE[i]['impliedVolatility'],
                    CE[i]['lastPrice'],
                    CE[i]['strikePrice']
        ]
    else:
        l_CE= list(['-','-','-','-',sp])

    if PE[i] != '-':
        l_PE = [
                    PE[i]['openInterest'],
                    PE[i]['changeinOpenInterest'],
                    PE[i]['impliedVolatility'],
                    PE[i]['lastPrice'],
                    PE[i]['underlyingValue']
        ]
    else:
        l_PE= list(['-','-','-','-','-'])
        
    l_OC_t = l_CE + l_PE
    l_OC_t[:] = [x if x != 0 else '-' for x in l_OC_t]
    
    l_OC.append(l_OC_t)

OC_col = ['C_OI', 'C_CHNG_IN_OI','C_IV', 'C_LTP', 'Strike','P_OI', 'P_CHNG_IN_OI','P_IV', 'P_LTP', 'SPOT']

pd.set_option('display.max_rows', None)
df = pd.DataFrame(l_OC)
df.columns = OC_col

import gspread
import pandas as pd
import numpy as np
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('E:\secret_key.json', scope)
file = gspread.authorize(creds)
df1 = pd.DataFrame(df)
sh = file.open_by_key('1VbA2AQXXQQAqYq46zeBfnOpNM957VdNdZQjsJXBGVBU')
worksheet = sh.get_worksheet(0)
set_with_dataframe(worksheet, df1)

print("BankNifty Completed")

import requests
import json
import pandas as pd

url = 'https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY'

headers={'user-agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'}
response = requests.get(url, headers=headers, timeout=10)

response_text=response.text
json_object = json.loads(response_text)

#print(json_object)

with open("OC.json", "w") as outfile:
    outfile.write(response_text)

# Below 'data' os a list 
data = json_object['records']['data']
#print(type(data))

# prepare list of expiry date from json_object. This list will be used to fetah OC date based on expiry date
e_date=json_object['records']['expiryDates']

oc_data = {}
for ed in e_date:
    oc_data[ed]={"CE":[], "PE":[]}
    for di in range(len(data)):
        
        if data[di]['expiryDate'] == ed:
            if 'CE' in data[di].keys() and data[di]['CE']['expiryDate'] == ed:
                
                oc_data[ed]["CE"].append(data[di]['CE'])
            else:
                oc_data[ed]["CE"].append('-')

            if 'PE' in data[di].keys() and data[di]['PE']['expiryDate'] == ed:
                oc_data[ed]["PE"].append(data[di]['PE'])
                
            else:
                oc_data[ed]["PE"].append('-')

# This loop is to delete extra keys from oc_data
for k in oc_data.keys():
    
    for i in range(len(oc_data[k]["CE"])):
        
        if oc_data[k]["CE"][i] != '-':
        
            del oc_data[k]["CE"][i]["expiryDate"]
            del oc_data[k]["CE"][i]["underlying"]
            del oc_data[k]["CE"][i]["identifier"]

        if oc_data[k]["PE"][i] != '-':

            del oc_data[k]["PE"][i]["expiryDate"]
            del oc_data[k]["PE"][i]["underlying"]
            del oc_data[k]["PE"][i]["identifier"]

expiry_date = e_date[0]

keys_to_exclue = ['pchangeinOpenInterest',
 'totalTradedVolume',
 'change',
 'pChange',
 'totalBuyQuantity',
 'totalSellQuantity',
 'bidQty',
 'bidprice',
 'askQty',
 'askPrice']

oc_data_dt=oc_data[expiry_date]

CE=list(oc_data_dt['CE'])
PE=list(oc_data_dt['PE'])

for i in range(len(CE)):
    if CE[i] != '-':
        for key in keys_to_exclue:
            del CE[i][key]
    if PE[i] != '-':
        for key in keys_to_exclue:
            del PE[i][key]

l_OC = []

for i in range(len(CE)):
    l_CE=[]
    l_PE=[]
    
    if CE[i] != '-':
        sp = CE[i]['strikePrice']
        l_CE = [
                    CE[i]['openInterest'],
                    CE[i]['changeinOpenInterest'],
                    CE[i]['impliedVolatility'],
                    CE[i]['lastPrice'],
                    CE[i]['strikePrice']
        ]
    else:
        l_CE= list(['-','-','-','-',sp])

    if PE[i] != '-':
        l_PE = [
                    PE[i]['openInterest'],
                    PE[i]['changeinOpenInterest'],
                    PE[i]['impliedVolatility'],
                    PE[i]['lastPrice'],
                    PE[i]['underlyingValue']
        ]
    else:
        l_PE= list(['-','-','-','-','-'])
        
    l_OC_t = l_CE + l_PE
    l_OC_t[:] = [x if x != 0 else '-' for x in l_OC_t]
    
    l_OC.append(l_OC_t)

OC_col = ['C_OI', 'C_CHNG_IN_OI','C_IV', 'C_LTP', 'Strike','P_OI', 'P_CHNG_IN_OI','P_IV', 'P_LTP', 'SPOT']

pd.set_option('display.max_rows', None)
df = pd.DataFrame(l_OC)
df.columns = OC_col

import gspread
import pandas as pd
import numpy as np
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('E:\secret_key.json', scope)
file = gspread.authorize(creds)
df1 = pd.DataFrame(df)
sh = file.open_by_key('1VbA2AQXXQQAqYq46zeBfnOpNM957VdNdZQjsJXBGVBU')
worksheet = sh.get_worksheet(1)
set_with_dataframe(worksheet, df1)

print("Nifty Completed")
