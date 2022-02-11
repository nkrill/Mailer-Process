import pandas as pd
import pyodbc
import math
import configparser
config = configparser.ConfigParser()
config.read('config.ini')
abp= pyodbc.connect(config['abp']['abpLogIn'])
Month='April 2022'
writer = pd.ExcelWriter('BATCH 2 EXAMPLE.xlsx', engine='xlsxwriter')
def addressInfo(IFile):
    accounts_id = list(IFile['Edc Account No'])
    n = math.ceil(len(accounts_id)/999)
    for i in range(n):
        accounts = "'" + "','".join(ele.replace("'", r"\'") for ele in accounts_id[999 * i:999 * (i + 1)]) + "'"

    abp_billing = f"SELECT LDC_ACCOUNT.LDC_ACCT_NO," \
                  f" Subquery.ADDR_1_TX as BILL_ADDR_1_TX,Subquery.ADDR_2_TX as BILL_ADDR_2_TX," \
                  f" Subquery.CITY_TX as BILL_CITY_TX,Subquery.STATE_TX as BILL_STATE_TX,Subquery.POSTAL_CD_TX as BILL_POSTAL_CD_TX" \
                  f" FROM ABPSYSTEM.EDI_TRANSACTION,(ABPSYSTEM.LDC_ACCOUNT LDC_ACCOUNT INNER JOIN ABPSYSTEM.ADDRESS ADDRESS ON (LDC_ACCOUNT.ADDR_ID = ADDRESS.ADDR_ID))" \
                  f" INNER JOIN (SELECT LDC_ACCOUNT_1.LDC_ACCT_ID, LDC_ACCOUNT_1.LDC_ACCT_NO, INVOICE_DIST_INFO.INV_NM, ADDRESS_1.ADDR_1_TX, ADDRESS_1.ADDR_2_TX, ADDRESS_1.CITY_TX," \
                  f" ADDRESS_1.STATE_TX, ADDRESS_1.POSTAL_CD_TX, ADDRESS_1.COUNTY_TX FROM ((ABPSYSTEM.INVOICE_DIST_INFO INVOICE_DIST_INFO" \
                  f" INNER JOIN ABPSYSTEM.ADDRESS ADDRESS_1 ON (INVOICE_DIST_INFO.ADDR_ID = ADDRESS_1.ADDR_ID)) INNER JOIN ABPSYSTEM.ACCOUNT ACCOUNT" \
                  f" ON (ACCOUNT.ACCT_ID = INVOICE_DIST_INFO.RELATE_ID)) INNER JOIN ABPSYSTEM.LDC_ACCOUNT LDC_ACCOUNT_1 ON (LDC_ACCOUNT_1.ACCT_ID = ACCOUNT.ACCT_ID)) Subquery" \
                  f" ON (LDC_ACCOUNT.LDC_ACCT_ID = Subquery.LDC_ACCT_ID) WHERE LDC_ACCOUNT.LDC_ACCT_ID = EDI_TRANSACTION.LDC_ACCT_ID" \
                  f" AND LDC_ACCOUNT.LDC_ACCT_NO in ({accounts})"""
    pand = pd.read_sql(abp_billing, abp)
    pand.drop_duplicates(subset ="LDC_ACCT_NO",inplace=True,keep='last')
    return pand
def addressInfoHelper(IFile):
    df=IFile[0:999]
   # df2 = IFile[1000:1999]
    #df3 = IFile[2000:2999]
    df=addressInfo(df)
   # df2 = addressInfo(df2)
    #df3 = addressInfo(df3)
   # df=pd.concat([df,df2, df3])
    IFile=pd.merge(IFile,df,how='left', left_on='Edc Account No',right_on='LDC_ACCT_NO')
    IFile['BILL_ADDR_1_TX']=IFile['BILL_ADDR_1_TX'].str.title()
    IFile['BILL_ADDR_2_TX']=IFile['BILL_ADDR_2_TX'].str.title()
    IFile['BILL_CITY_TX']=IFile['BILL_CITY_TX'].str.title()
    return IFile

inputFile=pd.read_excel('input/April_AR Data_Ohio2.xlsx',sheet_name='OH  Batch 2 -- Power',dtype=str)
inputFile['Full Name']=inputFile['Account Name 2'].str.title()+' '+inputFile['Account Name 1'].str.title()
inputFile['Renewal Rate']=inputFile['Renewal Rate'].astype(float)
inputFile['Renewal Rate']=inputFile['Renewal Rate'].round(decimals = 2)
print(inputFile['Edc Account No'].count()/1000)

AddDf=addressInfoHelper(inputFile)
FinalDF=pd.DataFrame({'Edc Account No':AddDf['Edc Account No'],
                      'Full Name':AddDf['Full Name'],
                      'Street address':AddDf['BILL_ADDR_1_TX'],
                      'Address 2':AddDf['BILL_ADDR_2_TX'],
                      'City':AddDf['BILL_CITY_TX'],
                      'State':AddDf['BILL_STATE_TX'],
                      'Zip':AddDf['BILL_POSTAL_CD_TX'],
                      'Expire Rate': AddDf['Price Charges'],
                      'Renewal Price':AddDf['Renewal Rate'],
                      'Expiration Date':Month,
                      'opt out date':AddDf['Opt Out Date'],
                      'Renewal Term End Date':AddDf['Renewal Term End Date'],
                      'Contract':AddDf['Contract Id'],'SERVICE State':AddDf['State'],
                      'EDC DB':AddDf['Edc D&B No']})




OHDF=FinalDF[FinalDF['SERVICE State']=='OH']
if not (OHDF.empty):
    OHDF=OHDF.drop("Expire Rate", axis=1)
    OHDF.to_excel(writer,sheet_name='OH',index=False)
MDDF=FinalDF[FinalDF['SERVICE State']=='MD']
if not (MDDF.empty):
    MDDF=MDDF.drop("Expire Rate", axis=1)
    MDDF.to_excel(writer,sheet_name='MD',index=False)
ILDF=FinalDF[FinalDF['SERVICE State']=='IL']
if not (ILDF.empty):
    ILDF.to_excel(writer,sheet_name='IL',index=False)
writer.save()