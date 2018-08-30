import pandas as pd
import numpy as np
from pandas import DataFrame
import pyodbc
import os
import sys
#reload(sys)
#sys.setdefaultencoding('utf-8')



BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CreditCardStatementPath = os.path.join(BASE_DIR,'CreditCardStatement.xlsx')
OutputExcelPath = os.path.join(BASE_DIR,'Reconciliation.xlsx')


#DateFrom and DateTo paramters for the query
DateFrom = "{d'2018-01-01'}"
DateTo = "{d'2018-07-31'}"


# open ODBC connection to Quickbooks and run sp_report to query UnCleared Credit Card Transaction
cn = pyodbc.connect('DSN=QuickBooks Data;')

sql = """sp_report CustomTxnDetail show Date, Account, ClearedStatus, Debit, Credit
parameters DateFrom ="""+DateFrom+""", DateTo = """+DateTo+""", SummarizeRowsBy = 'TotalOnly', AccountFilterType = 'CreditCard'
where RowType = 'DataRow' and AccountFullName Like '%BBVA Credit Card%' and ClearedStatus <> 'Cleared'
ORDER BY Credit ASC"""

#load data to DataFrame2
df2 = pd.read_sql(sql,cn)


df2['Debit'] = df2['Debit'].replace(np.nan,0)
df2['Credit'] = df2['Credit'].replace(np.nan,0)

df2['Transaction_Amount'] = df2['Credit'] - df2['Debit']
df2.drop(['ClearedStatus','Debit','Credit',], axis=1,inplace=True)

df2 = df2.sort_values(['Date','Transaction_Amount'],ascending=[True,True])

df2['Combine']= df2['Transaction_Amount'].astype(str)+ '|' + \
                df2['Account'].str[8:-4].str.strip().str.upper() + '|' + \
                df2['Account'].str[-4:].str.strip()

list3 = []
counter2 = []

for index, row in df2.iterrows():
    list3.append(row['Combine'])
    counter2.append(list3.count(row['Combine']))

df2['Counter'] = counter2
df2['Combine'] = df2['Combine'] + '|' + df2['Counter'].astype(str)


## open Excel file from bank statement, create dataframe from worksheet

df = pd.read_excel(CreditCardStatementPath, header=0)

#drop unneccessary columns
df = df.drop(df.columns[[0,2,4,5]],axis = 1)

#rename some columns
df.rename(columns ={'FIN.PRIMARY TRANSACTION AMOUNT':'Transaction Amount','ACC.ACCOUNT NAME':'AcctName','ACC.ACCOUNT NUMBER':'AcctNumber','FIN.TRANSACTION DATE':'Date'}, inplace = True)

#check if the transaction day is a businessday or not
df['Date'] = pd.to_datetime(df['Date'])
df['Is_Business_Day']= [np.is_busday(x) for x in df['Date'].astype(str)]


#sort the dataframe by Transaction Amount
df = df.sort_values(['Date','Transaction Amount'],ascending=[True,True])

#Upper case the account name column, remove white space
df['AcctName'] = df['AcctName'].str.upper().str.strip()


df['Combine'] = df['Transaction Amount'].astype(str) + '|' \
                + df['AcctName'] + '|' \
                + df['AcctNumber'].str[-4:].str.strip()

#the real account number is the last four digits of Account Number columns
df['AcctNumber'] = df['AcctNumber'].str[-4:]


list3 = []
counter = []

for index, row in df.iterrows():
    list3.append(row['Combine'])
    counter.append(list3.count(row['Combine']))

df['Counter'] = counter
df['Combine'] = df['Combine'] + '|' + df['Counter'].astype(str)

#match the "Combine" values in 2 dataframe, return true if match, false if not match
df['Matched'] = df['Combine'].isin(df2['Combine'])


df2['Matched'] = df2['Combine'].isin(df['Combine'])


writer = pd.ExcelWriter(OutputExcelPath,engine='xlsxwriter')
df.to_excel(writer,sheet_name='Sheet1',startcol=0,startrow=0,index=False,header=True,engine='xlsxwriter')
df2.to_excel(writer,sheet_name='Sheet1',startcol=15,startrow=0,index=False,header=True,engine='xlsxwriter')

writer.save()

cn.close()
