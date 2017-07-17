import pandas as pd
from pandas import DataFrame

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

CreditCardStatementPath = 'CreditCardStatement.xlsx'
QuickbooksExcelPath = 'Quickbooks.xlsx'
OutputExcelPath = 'Reconciliation2.xlsx'

## slice the string, take number of chars from the right
def right(text, num_chars):
    return text[-num_chars:]

## slice the string, obmit number of chars from the right.
def left(text, num_chars):
    return text[:-num_chars]

## open Excel file from bank statement, create dataframe from worksheet
df = pd.read_excel(CreditCardStatementPath, header=0)

#initiate Account Name list and append 'Account Name' from worksheet to list
Account_Name = []
for i in df['ACC.ACCOUNT NAME']:
    Account_Name.append(str(i))

#initiate Account Name list and append 'Transaction Amount' from worksheet to list
Transaction_Amount = []
for i in df['FIN.TRANSACTION AMOUNT']:
    Transaction_Amount.append(float(i))


#initiate Account Number list and append 'Account Number' from worksheet to list
Account_Number = []
for i in df['ACC.ACCOUNT NUMBER']:
    #take 4 leter from the right
    tempAccountNumber = int(right(i, 4))
    Account_Number.append(tempAccountNumber)
    

#merge Transaction Amount, Account Name and Account Number to a list
list2 = []
for k, v, h in zip(Transaction_Amount, Account_Name, Account_Number):
    list2.append(str(k) + '|' + str(v) + '|' + str(h))

list3 = []
counter = []

for i in list2:
    list3.append(i)
    counter.append(list3.count(i))

##Helper Value is a join of Transaction Amount, Account Number, Account Name, and Counter
HelperValue = []
for k, v in zip(list3, counter):
    HelperValue.append(str(k) + '|' + str(v))


##############################

## open Excel file from Quickbooks, create dataframe from worksheet
df2 = pd.read_excel(QuickbooksExcelPath,header=0)


#initiate Account Name list and append 'Credit' from worksheet to list
Transaction_Amount2 = []
for i in df2['Credit']:
    Transaction_Amount2.append(float(i))


#initiate Account Name list and append 'Account' from worksheet to list
#this include 'CardHolder Name and last 4 digits for Credit Card', will separate later.
Account_Name_Number = []
for i in df2['Account']:
    length_of_accountnumber = len(str(i))
    tempAccountNumber = right(str(i), length_of_accountnumber - 9)
    Account_Name_Number.append(tempAccountNumber)

Account_Name2 = []
for i in Account_Name_Number:
    tempAccountNumber = left(str(i),5)
    Account_Name2.append(tempAccountNumber.upper())

Account_Number2 = []
for i in Account_Name_Number:
    tempAccountNumber = right(str(i),4)
    Account_Number2.append(tempAccountNumber)

list2 = []
for k, v, h in zip(Transaction_Amount2, Account_Name2, Account_Number2):
    list2.append(str(k) + '|' + str(v) + '|' + str(h))

list3 = []
counter2 = []
HelperValue2 = []
for i in list2:
    list3.append(i)
    counter2.append(list3.count(i))

for k, v in zip(list3, counter2):
    HelperValue2.append(str(k) + '|' + str(v))


Match_BankStatement = []
for i in HelperValue:
    if i not in HelperValue2:
        Match_BankStatement.append('Not Match')
    else:
        Match_BankStatement.append('Match')


Match_Quickbooks = []
for i in HelperValue2:
    if i not in HelperValue:
        Match_Quickbooks.append('Not Match')
    else:
        Match_Quickbooks.append('Match')


writer = pd.ExcelWriter(OutputExcelPath,engine='xlsxwriter')

df = DataFrame(list(zip(Account_Name, Transaction_Amount, Account_Number, counter, HelperValue,Match_BankStatement)), columns=['Account_Name', 'Transaction_Amount', 'Account_Number', 'Counter', 'Helper','Match'])
df = df.sort_values(by = ['Transaction_Amount'],ascending = True)
df.to_excel(writer,sheet_name='Sheet1',startcol=0,startrow=0,index=False,header=True,engine='xlsxwriter')


df2 = DataFrame(list(zip(Account_Name2, Transaction_Amount2, Account_Number2, counter2, HelperValue2,Match_Quickbooks)), columns=['Account_Name', 'Transaction_Amount', 'Account_Number', 'Counter', 'Helper','Match'])
df2 = df2.sort_values(by=['Transaction_Amount'],ascending = True)
df2.to_excel(writer,sheet_name='Sheet1',startcol=10,startrow=0,index=False,header=True,engine='xlsxwriter')

writer.save()
