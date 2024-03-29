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
DateFrom = "{d'2022-01-01'}"
DateTo = "{d'2022-09-30'}"


# open ODBC connection to Quickbooks and run sp_report to query UnCleared Credit Card Transaction
cn = pyodbc.connect('DSN=QuickBooks Data;')

sql = f"""sp_report CustomTxnDetail show Date, RefNumber, Account, ClearedStatus, Debit, Credit
parameters DateFrom ={DateFrom}, DateTo = {DateTo}, SummarizeRowsBy = 'TotalOnly', AccountFilterType = 'CreditCard'
where RowType = 'DataRow' and AccountFullName Like '%PNC Credit Card (BBVA)%' and ClearedStatus <> 'Cleared'
ORDER BY Credit ASC"""

#load data to DataFrame2
df2 = pd.read_sql(sql,cn)

print (sql)

df2['Debit'] = df2['Debit'].replace(np.nan,0)
df2['Credit'] = df2['Credit'].replace(np.nan,0)

df2['Transaction_Amount'] = df2['Credit'] - df2['Debit']
df2.drop(['ClearedStatus','Debit','Credit',], axis=1,inplace=True)

df2 = df2.sort_values(['Transaction_Amount','Date'],ascending=[True,True])

df2['Combine']= df2['Transaction_Amount'].astype(str)+ '|' + \
                df2['Account'].str[8:-4].str.strip().str.upper() + '|' + \
                df2['Account'].str[-4:].str.strip()

df2['Counter'] = df2.groupby(['Combine']).cumcount().add(1)
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
df['Date'] = df['Date'].dt.date


#sort the dataframe by Transaction Amount
df = df.sort_values(['Transaction Amount','Date'],ascending=[True,True])

#Upper case the account name column, remove white space
df['AcctName'] = df['AcctName'].str.upper().str.strip()


df['Combine'] = df['Transaction Amount'].astype(str) + '|' \
                + df['AcctName'] + '|' \
                + df['AcctNumber'].str[-4:].str.strip()

#the real account number is the last four digits of Account Number columns
df['AcctNumber'] = df['AcctNumber'].str[-4:]


df['Counter'] = df.groupby(['Combine']).cumcount().add(1)
df['Combine'] = df['Combine'] + '|' + df['Counter'].astype(str)

#match the "Combine" values in 2 dataframe, return true if match, false if not match
df['Matched'] = df['Combine'].isin(df2['Combine'])


df2['Matched'] = df2['Combine'].isin(df['Combine'])


writer = pd.ExcelWriter(OutputExcelPath,engine='xlsxwriter')
df.to_excel(writer,sheet_name='Sheet1',startcol=0,startrow=0,index=False,header=True,engine='xlsxwriter')
df2.to_excel(writer,sheet_name='Sheet1',startcol=15,startrow=0,index=False,header=True,engine='xlsxwriter')
numberformat = writer.book.add_format({'num_format': '#,##0.00'})
writer.sheets['Sheet1'].set_column('C:C', None, numberformat)
writer.sheets['Sheet1'].set_column('R:R', None, numberformat)
writer.sheets['Sheet1'].autofilter('A1:V20000')

#save two dataframe to excel files
writer.save()

#close connection to Quickbooks
cn.close()

#automatically open the Reconciliation.xls from Excel
os.startfile(OutputExcelPath)
