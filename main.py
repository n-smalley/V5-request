import pandas as pd
import win32com.client as win32
import os
import json
from datetime import datetime

d = r'c:\Users\nathansmalley\OneDrive - Cook County Government\General - DBMS Private - Budget\FY2025\V5 Report'

# get json data
while True:
    key = input('Enter Office Number: ')
    with open(r'officeContacts.json','r') as file:
        jsn = json.load(file)

    try:
        data = jsn[key]
    except KeyError:
        print('Unknown Office Number\n')
    else:
        break
for i in data:
    inp = input(f'{i.upper()}: {data[i]} ')
    if inp:
        data[i] = [j for j in inp.replace(', ','').split(',')]
        print(f'{i.upper()}: {data[i]}')
fltr = data['office']

# Get most recent file
files = [os.path.join(d,f) for f in os.listdir(d) if os.path.isfile(os.path.join(d,f))]
most_recent = max(files,key=os.path.getmtime)
date = datetime.strptime(os.path.split(most_recent)[1].split('_')[1],'%Y-%m-%d').strftime('%m.%d.%y')

# Clean dataframe
df = pd.read_excel(most_recent)
df.columns = df.iloc[0]
df = df.drop(0).reset_index(drop=True)
if fltr:
    df = df.loc[df['Office'].str[:4].isin(fltr)].reset_index(drop=True)

# export df to excel
path = os.path.join(os.getcwd(),'outgoing',f'e{key.capitalize()} V5 {datetime.today().strftime('%m.%d.%y')}.xlsx')
df.to_excel(path,index=False)

# create email object and send
outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)

for i in data:
    if data[i] == []:
        data[i] = None
    data[i] = f'{data[i]}'.replace("['",'').replace("']",'').replace("', '",'; ')
    
mail.To = data['to']
mail.CC = data['cc']
mail.Subject = f'{data['name']} V5 report'
mail.Body = f'Please find report as of {date} attached. Let me know if you need anything else.'
mail.Attachments.Add(path)

mail.Send()