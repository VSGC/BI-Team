# -*- coding: utf-8 -*-
"""
Created on Thu Nov 10 10:52:38 2022

@author: VAIBHAV.SRIVASTAV01
"""


import pandas as pd
import os
import datetime as date
import calendar
import numpy as np
import pyodbc
from functools import reduce
calendar.month
todays_date= date.date.today()

con_dm = pyodbc.connect('DSN=GHF_BI_CONN_DM;UID=GHF_BI_CONN;PWD=Godrej@123')
con = pyodbc.connect('DSN=GHF_BI_CONN;UID=GHF_BI_CONN;PWD=Godrej@123')
con_dm_nbfc = pyodbc.connect('DSN=GHF_BI_CONN_DM_NBFC;UID=GHF_BI_CONN;PWD=Godrej@123')


ghf_loan_view=pd.read_sql("SELECT * from prod_da_db.serve.loan_view WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm)
ghf_loan_view['NBFC_FLAG'] = 'N'
ghf_base_view = pd.read_sql("SELECT * from prod_da_db.serve.BASE_VIEW WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm)
ghf_base_view['NBFC_FLAG'] = 'N'
ghf_disb=pd.read_sql("select * from  prod_da_db.serve.f_finance_disbursement_details WHERE DH_RECORD_ACTIVE_FLAG = 'Y';", con_dm)
ghf_disb['NBFC_FLAG'] = 'N'




gfl_loan_view=pd.read_sql("SELECT * from prod_gfl_da_db.serve.loan_view WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm_nbfc)
gfl_loan_view['NBFC_FLAG'] = 'Y'
gfl_base_view = pd.read_sql("SELECT * from prod_gfl_da_db.serve.BASE_VIEW WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm_nbfc)
gfl_base_view['NBFC_FLAG'] = 'Y'
gfl_disb=pd.read_sql("select * from  prod_gfl_da_db.serve.f_finance_disbursement_details WHERE DH_RECORD_ACTIVE_FLAG = 'Y';", con_dm_nbfc)
gfl_disb['NBFC_FLAG'] = 'Y'

loan_view=pd.concat([ghf_loan_view,gfl_loan_view]) #REFERENCE
base_view=pd.concat([ghf_base_view,gfl_base_view]) # LAN_ID
# base_view.to_excel(r"C:\Users\VAIBHAV.SRIVASTAV01\Desktop\base_view.xlsx")
loan_view.rename(columns={'REFERENCE':'LAN_ID'},inplace=True)
loan_view=loan_view[['LAN_ID','LOAN_PURPOSE','ROI','PRINCIPAL_OUTSTANDING','GPL_FLAG']]
base_view=base_view.merge(loan_view, on='LAN_ID',how='left')
disb=pd.concat([ghf_disb,gfl_disb])
disb=disb[['FINANCE_REFERENCE','DISBURSEMENT_DATE', 'DISBURSEMENT_AMOUNT','DISBURSEMENT_SEQUENCE']]
disb.rename(columns={'FINANCE_REFERENCE':'LAN_ID'},inplace=True)
base_view['FINTYPE']=np.where(base_view['LOAN_PURPOSE'].isin(['LAP Balance Transfer plus Top-up','Loan against Property', 'Industrial LAP Balance Transfer','Industrial LAP Balance Transfer plus Top-up','LAP Balance Transfer','Loan against industrial property','LAP Top Up']),'LP',base_view['FINANCE_TYPE'])
# branch=base_view.copy()
# branch=branch[['LAN_ID','FINTYPE','REPORTING_BRANCH']]
# disb=disb.merge(branch,on='LAN_ID',how='left')
## disb=disb[['FINTYPE','REPORTING_BRANCH','DISBURSEMENT_DATE', 'DISBURSEMENT_AMOUNT']]
base_view['GPLFLAG_SANCTIONS']=np.where((base_view['FINTYPE'].isin(['LP','NP'])==False) & (base_view['GPL_FLAG']=='YES') & (base_view['NBFC_FLAG']=='N') , 'GPL',np.where((base_view['FINTYPE'].isin(['LP','NP'])==False) &(base_view['GPL_FLAG']=='NO') & (base_view['NBFC_FLAG']=='N'),'NON GPL','NIL'))


def error_lans():
    a=['GHF1001FL0000328',
 'GHF1001FL0002120',
 'GHF1002FL0002436',
 'GHF1001FL0002119',
 'GHF1002FL0000299',
 'GHF1002FL0000640',
 'GHF1002HL0002943',
 'GHF1002HL0000836',
 'GHF1002FT0001610',
 'GHF1002FL0001993',
 'GHF1002HL0002773',
 'GHF1001FT0000552',
 'GHF1101LP0003918',
 'GHF1401HL0000123',
 'GHF1002HL0003082',
 'GHF1002FL0000605',
 'GHF1002HL0002977',
 'GHF1001FL0000327',
 'GHF1002HL0003004',
 'GHF1002FL0005334',
 'GHF1001FL0000912',
 'GHF1002FL0000257',
 'GHF1002HL0003367',
 'GHF1002FL0000744',
 'GHF1002FT0001009',
 'GHF1001FL0000329',
 'GHF1401HL0000122',
 'GHF1001HL0000342',
 'GHF1002HL0001014',
 'GHF1002HL0003484',
 'GHF1401HL0001902',
 'GHF1001HL0000900',
 'GHF1002FL0000063',
 'GHF1001FL0002118',
 'GHF1401HL0000124',
 'GHF1002FL0002252',
 'GHF1401HL0000125',
 'GHF1002FL0003178',
 'GHF1002FL0000659',
 'GHF1001HL0000967',
 'GHF1001HL0000966']
    return a
def remove_error_lans(df):
    try:
        df=df[df['REFERENCE'].isin(error_lans())==False]
    except:
        pass
    try:
        df=df[df['LAN_ID'].isin(error_lans())==False]
    except:
        pass
    try:
        df=df[df['FINREFERENCE'].isin(error_lans())==False]
    except:
        pass
    try:
        df=df[df['FINANCE_REFERENCE'].isin(error_lans())==False]
    except:
        pass
    return df
base_view=remove_error_lans(base_view)

eomfin=base_view.copy()
eomfin['EOMLOGN'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)
eomfin['EOMSNCTN'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)
eomfin['EOMRJCT'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)
eomfin['EOMCLTRL'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)
eomfin['BOOKING_DATE'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)



eomfin['REPORTING_BRANCH']=np.where(eomfin['REPORTING_BRANCH']=='Aundh BO (Pune)','Pune',eomfin['REPORTING_BRANCH'])
eomfin['REPORTING_BRANCH']=np.where(eomfin['REPORTING_BRANCH']=='Bangalore (MG Road)','Bangalore',eomfin['REPORTING_BRANCH'])
eomfin['REPORTING_BRANCH']=np.where(eomfin['REPORTING_BRANCH']=='Banglore (Sahakar Nagar)','Bangalore',eomfin['REPORTING_BRANCH'])
eomfin['REPORTING_BRANCH']=np.where(eomfin['REPORTING_BRANCH']=='Gurgaon','Delhi',eomfin['REPORTING_BRANCH'])
eomfin['REPORTING_BRANCH']=np.where(eomfin['REPORTING_BRANCH']=='Thane','Mumbai',eomfin['REPORTING_BRANCH'])

bussdate=pd.read_sql("select * from prod_bi.aabi.business_date;",con_dm)
bussdate1=bussdate.copy()
today=date.datetime.now()

bussdate1.rename(columns={'BUSINESS_DATE':'EOMLOGN'},inplace=True)
eomfin.rename(columns={'EOMLOGN':'CAL_DATE'},inplace=True)
eomfin['CAL_DATE']=pd.to_datetime(eomfin['CAL_DATE']).dt.date
eomfin=eomfin.merge(bussdate1,on='CAL_DATE')
# eomfin['EOMLOGN']=np.where(pd.to_datetime(eomfin['CAL_DATE']).dt.date>pd.to_datetime(eomfin['BOOKING_DATE']).dt.date,pd.to_datetime(eomfin['BOOKING_DATE']).dt.date,(pd.to_datetime(eomfin['EOMLOGN']).dt.date))
# eomfin['EOMSNCTN']=np.where(pd.to_datetime(eomfin['EOMSNCTN']).dt.date>pd.to_datetime(eomfin['BOOKING_DATE']).dt.date,pd.to_datetime(eomfin['BOOKING_DATE']).dt.date,(pd.to_datetime(eomfin['EOMSNCTN']).dt.date))

bussdate2=bussdate.copy()
bussdate2.rename(columns={'BUSINESS_DATE':'EOMSNCTN'},inplace=True)
eomfin.rename(columns={'EOMSNCTN':'PSV1'},inplace=True)
bussdate2.rename(columns={'CAL_DATE':'PSV1'},inplace=True)
eomfin['PSV1']=pd.to_datetime(eomfin['PSV1']).dt.date
eomfin['PSV1'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)
eomfin=eomfin.merge(bussdate2,on='PSV1',how='left')
eomfin['EOMSNCTN'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)


bussdate3=bussdate.copy()
bussdate3.rename(columns={'BUSINESS_DATE':'BOOKING_DATE'},inplace=True)
eomfin.rename(columns={'BOOKING_DATE':'booked_date'},inplace=True)
bussdate3.rename(columns={'CAL_DATE':'booked_date'},inplace=True)
eomfin['booked_date']=pd.to_datetime(eomfin['booked_date']).dt.date
eomfin['booked_date'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)
eomfin=eomfin.merge(bussdate3,on='booked_date',how='left')
eomfin['BOOKING_DATE'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)

bussdate4=bussdate.copy()
bussdate4.rename(columns={'BUSINESS_DATE':'EOMRJCT'},inplace=True)
eomfin.rename(columns={'EOMRJCT':'REJECT_DATE'},inplace=True)
bussdate4.rename(columns={'CAL_DATE':'REJECT_DATE'},inplace=True)
eomfin['REJECT_DATE']=pd.to_datetime(eomfin['REJECT_DATE']).dt.date
eomfin['REJECT_DATE'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)
eomfin=eomfin.merge(bussdate4,on='REJECT_DATE',how='left')
eomfin['EOMRJCT'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)

bussdate5=bussdate.copy()
bussdate5.rename(columns={'BUSINESS_DATE':'EOMCLTRL'},inplace=True)
eomfin.rename(columns={'EOMCLTRL':'IS_DATE'},inplace=True)
bussdate5.rename(columns={'CAL_DATE':'IS_DATE'},inplace=True)
eomfin['IS_DATE']=pd.to_datetime(eomfin['IS_DATE']).dt.date
eomfin['IS_DATE'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)
eomfin=eomfin.merge(bussdate5,on='IS_DATE',how='left')
eomfin['EOMCLTRL'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)


eomfin['LOGIN_YEAR']=pd.to_datetime(eomfin['EOMLOGN']).dt.year
eomfin['LOGIN_YEAR']=eomfin['LOGIN_YEAR'].astype(int)
eomfin['LOGIN_MONTH']=pd.to_datetime(eomfin['EOMLOGN']).dt.month
eomfin['LOGIN_MONTH']=eomfin['LOGIN_MONTH'].astype(int)
eomfin['FS_YEAR']=pd.to_datetime(eomfin['EOMSNCTN']).dt.year
eomfin['FS_YEAR']=eomfin['FS_YEAR'].astype(int)
eomfin['FS_MONTH']=pd.to_datetime(eomfin['EOMSNCTN']).dt.month
eomfin['FS_MONTH']=eomfin['FS_MONTH'].astype(int)
eomfin['BOOKING_YEAR']=pd.to_datetime(eomfin['BOOKING_DATE']).dt.year
eomfin['BOOKING_YEAR']=eomfin['BOOKING_YEAR'].astype(int)
eomfin['BOOKING_MONTH']=pd.to_datetime(eomfin['BOOKING_DATE']).dt.month
eomfin['BOOKING_MONTH']=eomfin['BOOKING_MONTH'].astype(int)
eomfin['REJECT_YEAR']=pd.to_datetime(eomfin['EOMRJCT']).dt.year
eomfin['REJECT_YEAR']=eomfin['REJECT_YEAR'].astype(int)
eomfin['REJECT_MONTH']=pd.to_datetime(eomfin['EOMRJCT']).dt.month
eomfin['REJECT_MONTH']=eomfin['REJECT_MONTH'].astype(int)
eomfin['IS_YEAR']=pd.to_datetime(eomfin['EOMCLTRL']).dt.year
eomfin['IS_YEAR']=eomfin['IS_YEAR'].astype(int)
eomfin['IS_MONTH']=pd.to_datetime(eomfin['EOMCLTRL']).dt.month
eomfin['IS_MONTH']=eomfin['IS_MONTH'].astype(int)
disb['DISBURSE_MONTH']=pd.to_datetime(disb['DISBURSEMENT_DATE']).dt.month
disb['DISBURSE_YEAR']=pd.to_datetime(disb['DISBURSEMENT_DATE']).dt.year
eomfin['LOGIN_YEAR_MONTH']=eomfin['LOGIN_YEAR'].astype(str)+'-'+ eomfin['LOGIN_MONTH'].astype(str)
eomfin['FS_YEAR_MONTH']=eomfin['FS_YEAR'].astype(str)+'-'+ eomfin['FS_MONTH'].astype(str)
eomfin['BOOK_YEAR_MONTH']=eomfin['BOOKING_YEAR'].astype(str)+'-'+ eomfin['BOOKING_MONTH'].astype(str)
disb['DISB_YEAR_MONTH']=disb['DISBURSE_YEAR'].astype(str)+'-'+ disb['DISBURSE_MONTH'].astype(str)
eomfin['REJECT_YEAR_MONTH']=eomfin['REJECT_YEAR'].astype(str)+'-'+ eomfin['REJECT_MONTH'].astype(str)
eomfin['IS_YEAR_MONTH']=eomfin['IS_YEAR'].astype(str)+'-'+ eomfin['IS_MONTH'].astype(str)
eomfin['BOOKING_DATE']=np.where(eomfin['LAN_ID'].isin(['GHF1002HL0010862' ,'GHF1002FL0010865' ,'GHF1002HL0010527' ,'GHF1002HL0010885' ,'GHF1002FL0010994' ,'GHF1003HL0010367' ,'GHF1003FL0010946' ,'GHF1002HL0010399' ,'GHF1002HT0010434' ,'GHF1003HL0009576' ,'GHF1002HL0010722' ,'GHF1002HL0010827']),date.date(2022,4,1),pd.to_datetime(eomfin['BOOKING_DATE']).dt.date)
eomfin['BOOK_YEAR_MONTH']=np.where(eomfin['LAN_ID'].isin(['GHF1002HL0010862' ,'GHF1002FL0010865' ,'GHF1002HL0010527' ,'GHF1002HL0010885' ,'GHF1002FL0010994' ,'GHF1003HL0010367' ,'GHF1003FL0010946' ,'GHF1002HL0010399' ,'GHF1002HT0010434' ,'GHF1003HL0009576' ,'GHF1002HL0010722' ,'GHF1002HL0010827']),'2022-4',eomfin['BOOK_YEAR_MONTH'])
eomfin=eomfin[['REPORTING_BRANCH','LAN_ID',  'FINTYPE', 'LOAN_PURPOSE', 'LOAN_STATUS',
       'BOOKING_AMOUNT', 'BOOKING_DATE', 'ROI', 'NET_PREMIUM',
       'FINANCE_SOURCE_ID', 'EOMLOGN', 'EOMSNCTN', 'REQUESTED_AMOUNT',
       'SANCTION_AMOUNT', 'NBFC_FLAG',
       'DETAILED_STATUS', 'STATUS', 'QUEUE', 'LOGIN_STATUS', 'STATUS_SEG',
       'LOGIN_MONTH', 'LOGIN_YEAR_MONTH', 'FS_MONTH', 'FS_YEAR_MONTH',
       'BOOKING_MONTH', 'BOOK_YEAR_MONTH','REJECT_YEAR_MONTH','IS_YEAR_MONTH','EOMRJCT','EOMCLTRL','GPLFLAG_SANCTIONS']]
eomdisb_1=disb.copy()
################################ SUMMARY ################################
def end_of(MTD,eomfin,eomdisb_1):

    eomfin_book=eomfin.copy()
    
    if type(MTD)==datetime.datetime:
        eomfin_book['bookdate']=np.where(True,pd.to_datetime(eomfin_book['BOOKING_DATE']).dt.date ,0)
        eomfin_book=eomfin_book[eomfin_book['bookdate']==(MTD.date())]
        # eomfin_book['BOOK_VOL']= np.where(eomfin_book['bookdate']==(MTD.date()),1,0)
    else:
        eomfin_book=eomfin_book[eomfin_book['BOOK_YEAR_MONTH'].isin(MTD)]
        # eomfin_book['BOOK_VOL']= np.where((eomfin_book['BOOK_YEAR_MONTH'].isin(MTD)),1,0)
    eomfin_book=eomfin_book[eomfin_book['STATUS']=='Booked']
    eomfin_book['BOOK_VOL']= np.where(eomfin_book['STATUS']=='Booked',1,0)
    
    eomfin_book['BOOKING_AMOUNT']=eomfin_book['BOOKING_AMOUNT'].round(2)
    eomfin_book['WROI']=eomfin_book['ROI']*eomfin_book['BOOKING_AMOUNT']
    eomfin_book=eomfin_book[[ 'LAN_ID','REPORTING_BRANCH', 'BOOKING_AMOUNT','BOOK_VOL','FINTYPE','WROI','NET_PREMIUM','GPLFLAG_SANCTIONS']]
    eomfin_bookview=eomfin_book.groupby('REPORTING_BRANCH').sum(['BOOKING_AMOUNT','BOOK_VOL','WROI'])
    eomfin_bookview['ROI']=eomfin_bookview['WROI']/eomfin_bookview['BOOKING_AMOUNT']
    eomfin_bookview['Premium']=eomfin_bookview['NET_PREMIUM']/100000
    eomfin_bookview['Penetration']=eomfin_bookview['NET_PREMIUM']*100/eomfin_bookview['BOOKING_AMOUNT']
    eomfin_bookview['BOOKING_AMOUNT']=eomfin_bookview['BOOKING_AMOUNT']/10000000
    eomfin_bookview=eomfin_bookview[[  'BOOKING_AMOUNT','BOOK_VOL','ROI','Premium','Penetration']]
    eomfin_book['Prod']=np.where(eomfin_book['FINTYPE'].isin(['LP','NP']),'LAP',np.where(eomfin_book['GPLFLAG_SANCTIONS']=='GPL','GPL','NON GPL')) 
    eomfin_pbview=eomfin_book.groupby('Prod').sum(['BOOKING_AMOUNT','BOOK_VOL','WROI'])
    eomfin_pbview['ROI']=eomfin_pbview['WROI']/eomfin_bookview['BOOKING_AMOUNT']
    eomfin_pbview['Premium']=eomfin_pbview['NET_PREMIUM']/100000
    eomfin_pbview['Penetration']=eomfin_pbview['NET_PREMIUM']*100/eomfin_bookview['BOOKING_AMOUNT']
    eomfin_pbview['BOOKING_AMOUNT']=eomfin_pbview['BOOKING_AMOUNT']/10000000
    eomfin_pbview=eomfin_pbview[[  'BOOKING_AMOUNT','BOOK_VOL','ROI','Premium','Penetration']]
    
    eomfin_log=eomfin.copy()
    
    if type(MTD)==datetime.datetime:
        eomfin_log['logdate']=np.where(True,pd.to_datetime(eomfin_log['EOMLOGN']).dt.date ,0)
        eomfin_log=eomfin_log[eomfin_log['logdate']==(MTD.date())]
        # eomfin_log['LOG_VOL']= np.where(eomfin_log['logdate']==(MTD.date()),1,0)
    else:
        eomfin_log=eomfin_log[eomfin_log['LOGIN_YEAR_MONTH'].isin(MTD)]
        # eomfin_log['LOG_VOL']= np.where((eomfin_log['LOGIN_YEAR_MONTH'].isin(MTD)),1,0)
    eomfin_log['LOG_VOL']= np.where(eomfin_log['LOGIN_STATUS']=='A) Login',1,0)
    eomfin_log=eomfin_log[eomfin_log['LOGIN_STATUS']=='A) Login']
    eomfin_log['REQUESTED_AMOUNT']=eomfin_log['REQUESTED_AMOUNT']/10000000
    eomfin_log['REQUESTED_AMOUNT']=eomfin_log['REQUESTED_AMOUNT'].round(2)
    eomfin_log=eomfin_log[[ 'LAN_ID', 'REPORTING_BRANCH','REQUESTED_AMOUNT','LOG_VOL','FINTYPE','GPLFLAG_SANCTIONS']]
    eomfin_logview=eomfin_log.groupby('REPORTING_BRANCH').sum(['REQUESTED_AMOUNT','LOG_VOL'])
    eomfin_log['Prod']=np.where(eomfin_log['FINTYPE'].isin(['LP','NP']),'LAP',np.where(eomfin_log['GPLFLAG_SANCTIONS']=='GPL','GPL','NON GPL'))
    eomfin_plview=eomfin_log.groupby('Prod').sum(['REQUESTED_AMOUNT','LOG_VOL'])
    
    
    eomfin_sanc=eomfin.copy()
    
    if type(MTD)==datetime.datetime:
        eomfin_sanc['sancdate']=np.where(True,pd.to_datetime(eomfin_sanc['EOMSNCTN']).dt.date ,0)
        eomfin_sanc=eomfin_sanc[eomfin_sanc['sancdate']==(MTD.date())]
        # eomfin_sanc['SANCTION_VOL']= np.where((eomfin_sanc['sancdate']==(MTD.date())),1,0)
    else:
        eomfin_sanc=eomfin_sanc[eomfin_sanc['FS_YEAR_MONTH'].isin(MTD)]
        # eomfin_sanc['SANCTION_VOL']= np.where((eomfin_sanc['FS_YEAR_MONTH'].isin(MTD)),1,0)
    eomfin_sanc['SANCTION_VOL']= np.where(eomfin_sanc['STATUS_SEG']=='A) Final Sanction',1,0)
    eomfin_sanc=eomfin_sanc[eomfin_sanc['STATUS_SEG']=='A) Final Sanction']
    eomfin_sanc['SANCTION_AMOUNT']=eomfin_sanc['SANCTION_AMOUNT']/10000000
    eomfin_sanc['SANCTION_AMOUNT']=eomfin_sanc['SANCTION_AMOUNT'].round(2)
    eomfin_sanc=eomfin_sanc[[ 'LAN_ID','REPORTING_BRANCH', 'SANCTION_AMOUNT','SANCTION_VOL','FINTYPE','GPLFLAG_SANCTIONS']]
    eomfin_sancview=eomfin_sanc.groupby('REPORTING_BRANCH').sum(['SANCTION_AMOUNT','SANCTION_VOL'])

    eomfin_sanc['Prod']=np.where(eomfin_sanc['FINTYPE'].isin(['LP','NP']),'LAP',np.where(eomfin_sanc['GPLFLAG_SANCTIONS']=='GPL','GPL','NON GPL'))
    
    eomfin_psview=eomfin_sanc.groupby('Prod').sum(['SANCTION_AMOUNT','SANCTION_VOL'])
    
    eomfin_insanc=eomfin.copy()
    
    if type(MTD)==datetime.datetime:
        eomfin_insanc['inc_sancdate']=np.where(True,pd.to_datetime(eomfin_insanc['EOMCLTRL']).dt.date ,0)
        eomfin_insanc=eomfin_insanc[eomfin_insanc['inc_sancdate']==(MTD.date())]
        # eomfin_insanc['SANCTION_VOL']= np.where((eomfin_insanc['sancdate']==(MTD.date())),1,0)
    else:
        eomfin_insanc=eomfin_insanc[eomfin_insanc['IS_YEAR_MONTH'].isin(MTD)]
        # eomfin_insanc['SANCTION_VOL']= np.where((eomfin_insanc['FS_YEAR_MONTH'].isin(MTD)),1,0)
    eomfin_insanc['IN_SANCTION_VOL']= np.where(eomfin_insanc['STATUS_SEG']=='C) Income Sanction',1,0)
    eomfin_insanc=eomfin_insanc[eomfin_insanc['STATUS_SEG']=='C) Income Sanction']
    eomfin_insanc['IN_SANCTION_AMOUNT']=eomfin_insanc['SANCTION_AMOUNT']/10000000
    eomfin_insanc['IN_SANCTION_AMOUNT']=eomfin_insanc['SANCTION_AMOUNT'].round(2)
    eomfin_insanc=eomfin_insanc[[ 'LAN_ID','REPORTING_BRANCH', 'IN_SANCTION_AMOUNT','IN_SANCTION_VOL','FINTYPE','GPLFLAG_SANCTIONS']]
    eomfin_insancview=eomfin_insanc.groupby('REPORTING_BRANCH').sum(['IN_SANCTION_AMOUNT','IN_SANCTION_VOL'])
    eomfin_insanc['Prod']=np.where(eomfin_insanc['FINTYPE'].isin(['LP','NP']),'LAP',np.where(eomfin_insanc['GPLFLAG_SANCTIONS']=='GPL','GPL','NON GPL'))
    eomfin_inpsview=eomfin_insanc.groupby('Prod').sum(['IN_SANCTION_AMOUNT','IN_SANCTION_VOL'])
    
    eomfin_reject=eomfin.copy()
    
    if type(MTD)==datetime.datetime:
        eomfin_reject['rejectdate']=np.where(True,pd.to_datetime(eomfin_reject['EOMRJCT']).dt.date ,0)
        eomfin_reject=eomfin_reject[eomfin_reject['rejectdate']==(MTD.date())]
        # eomfin_reject['SANCTION_VOL']= np.where((eomfin_reject['sancdate']==(MTD.date())),1,0)
    else:
        eomfin_reject=eomfin_reject[eomfin_reject['REJECT_YEAR_MONTH'].isin(MTD)]
        # eomfin_reject['SANCTION_VOL']= np.where((eomfin_reject['FS_YEAR_MONTH'].isin(MTD)),1,0)
    eomfin_reject['REJECT_VOL']= np.where(eomfin_reject['STATUS_SEG']=='B) Rejected',1,0)
    eomfin_reject=eomfin_reject[eomfin_reject['STATUS_SEG']=='B) Rejected']
    eomfin_reject['REJECT_AMOUNT']=eomfin_reject['REQUESTED_AMOUNT']/10000000
    eomfin_reject['REJECT_AMOUNT']=eomfin_reject['REQUESTED_AMOUNT'].round(2)
    eomfin_reject=eomfin_reject[[ 'LAN_ID','REPORTING_BRANCH', 'REJECT_AMOUNT','REJECT_VOL','FINTYPE','GPLFLAG_SANCTIONS']]
    eomfin_rejectview=eomfin_reject.groupby('REPORTING_BRANCH').sum(['REJECT_AMOUNT','REJECT_VOL'])
    eomfin_reject['Prod']=np.where(eomfin_reject['FINTYPE'].isin(['LP','NP']),'LAP',np.where(eomfin_reject['GPLFLAG_SANCTIONS']=='GPL','GPL','NON GPL'))
    
    eomfin_rejectpsview=eomfin_reject.groupby('Prod').sum(['REJECT_AMOUNT','REJECT_VOL'])
    
    
    
    m1=pd.merge(eomfin_bookview,eomfin_logview,left_on=('REPORTING_BRANCH'),right_on=('REPORTING_BRANCH'), how='outer')
    m2=pd.merge(m1,eomfin_sancview,left_on=('REPORTING_BRANCH'),right_on=('REPORTING_BRANCH'), how='outer')
    m3=pd.merge(m2,eomfin_insancview,left_on=('REPORTING_BRANCH'),right_on=('REPORTING_BRANCH'), how='outer')
    eomfinview=pd.merge(m3,eomfin_rejectview,left_on=('REPORTING_BRANCH'),right_on=('REPORTING_BRANCH'), how='outer')
    
    
    t1=pd.merge(eomfin_pbview,eomfin_plview,left_on=('Prod'),right_on=('Prod'), how='outer')
    t2=pd.merge(t1,eomfin_psview,left_on=('Prod'),right_on=('Prod'), how='outer')
    t3=pd.merge(t2,eomfin_inpsview,left_on=('Prod'),right_on=('Prod'), how='outer')
    eompfinview=pd.merge(t3,eomfin_rejectpsview,left_on=('Prod'),right_on=('Prod'), how='outer')
    
    eomdisb=pd.merge(eomdisb_1,eomfin,on=('LAN_ID'),how='left')
    # eomdisb=eomdisb_1.copy()
    if type(MTD)==datetime.datetime:
        eomdisb['DISBDATE']=np.where(True,pd.to_datetime(eomdisb['DISBURSEMENT_DATE']).dt.date ,0)
        eomdisb=eomdisb[eomdisb['DISBDATE']==(MTD.date())]
        # eomdisb['DISB_VOL']=np.where((eomdisb['DISBDATE']==(MTD.date())),1,0)
    else:
        eomdisb=eomdisb[eomdisb['DISB_YEAR_MONTH'].isin(MTD)]
        # eomdisb['DISB_VOL']=np.where((eomdisb['DISB_YEAR_MONTH'].isin(MTD)),1,0)
    eomdisb['DISB_VOL']=np.where(True,1,0)
    eomdisb['DISBURSEMENT_AMOUNT']=eomdisb['DISBURSEMENT_AMOUNT']/1000000000
    eomdisb['DISBURSEMENT_AMOUNT']=eomdisb['DISBURSEMENT_AMOUNT'].round(2)
    eomdisb=eomdisb[[ 'REPORTING_BRANCH', 'DISBURSEMENT_AMOUNT','FINTYPE','GPLFLAG_SANCTIONS']]
    eomdisbview=eomdisb.groupby('REPORTING_BRANCH').sum(['DISBURSEMENT_AMOUNT'])
    eomdisb['Prod']=np.where(eomdisb['FINTYPE'].isin(['LP','NP']),'LAP',np.where(eomdisb['GPLFLAG_SANCTIONS']=='GPL','GPL','NON GPL'))
    eomfin_pdview=eomdisb.groupby('Prod').sum(['DISBURSEMENT_AMOUNT'])
    
    eomp=pd.merge(eompfinview,eomfin_pdview,left_on=('Prod'),right_on=('Prod'), how='outer')
    eom=pd.merge(eomfinview,eomdisbview,left_on=('REPORTING_BRANCH'),right_on=('REPORTING_BRANCH'), how='outer')
    eom=eom.sort_values(by='REPORTING_BRANCH')
    eomp=eomp.sort_values(by='Prod')
    eom=eom[['LOG_VOL','REQUESTED_AMOUNT','SANCTION_VOL','SANCTION_AMOUNT','BOOK_VOL','BOOKING_AMOUNT','DISBURSEMENT_AMOUNT','REJECT_AMOUNT','REJECT_VOL','IN_SANCTION_AMOUNT','IN_SANCTION_VOL','ROI','Premium','Penetration']]
    eomp=eomp[['LOG_VOL','REQUESTED_AMOUNT','SANCTION_VOL','SANCTION_AMOUNT','BOOK_VOL','BOOKING_AMOUNT','DISBURSEMENT_AMOUNT','REJECT_AMOUNT','REJECT_VOL','IN_SANCTION_AMOUNT','IN_SANCTION_VOL','ROI','Premium','Penetration']]
    
    monbr=list(set(list(base_view.REPORTING_BRANCH.unique()))-set(eom.index.to_list()))
    missprod=list(set(['GPL','NON GPL','LAP'])-set(eomp.index.to_list()))
    for i in monbr:
        eom.loc[i,:]=0
    for i in missprod:
        eomp.loc[i,:]=0
    eom=eom.sort_values(by='REPORTING_BRANCH')
    eomp=eomp.sort_values(by='Prod')
    eom.replace(to_replace=[np.nan,'',' '],value=0,inplace=True)
    eomp.replace(to_replace=[np.nan,'',' '],value=0,inplace=True)
    a=eom['ROI']*eom['BOOKING_AMOUNT']
    a=pd.DataFrame(a)
    a['BOOKING_AMOUNT']=eom['BOOKING_AMOUNT']
    a['NET_PREMIUM']=eom['Premium']*eom['BOOKING_AMOUNT']
    a.loc['Total']= a.sum()
    eom.loc['Total']= eom.sum()
    eom.loc['Total','ROI']=a.loc['Total',0]/a.loc['Total','BOOKING_AMOUNT']
    eom.loc['Total','Penetration']=a.loc['Total','NET_PREMIUM']/a.loc['Total','BOOKING_AMOUNT']
    eomp.loc['Total']= eomp.sum()
    eomp.loc['Total','ROI']=a.loc['Total',0]/a.loc['Total','BOOKING_AMOUNT']
    eomp.loc['Total','Penetration']=a.loc['Total','NET_PREMIUM']/a.loc['Total','BOOKING_AMOUNT']
    eom['Disbursement_Volume']=eom['BOOK_VOL']
    eomp['Disbursement_Volume']=eomp['BOOK_VOL']
    
    eom=eom[['LOG_VOL','REQUESTED_AMOUNT','SANCTION_VOL','SANCTION_AMOUNT','BOOK_VOL','BOOKING_AMOUNT','Disbursement_Volume','DISBURSEMENT_AMOUNT','REJECT_AMOUNT','REJECT_VOL','IN_SANCTION_AMOUNT','IN_SANCTION_VOL','ROI','Premium','Penetration']]
    eomp=eomp[['LOG_VOL','REQUESTED_AMOUNT','SANCTION_VOL','SANCTION_AMOUNT','BOOK_VOL','BOOKING_AMOUNT','Disbursement_Volume','DISBURSEMENT_AMOUNT','REJECT_AMOUNT','REJECT_VOL','IN_SANCTION_AMOUNT','IN_SANCTION_VOL','ROI','Premium','Penetration']]
    eom.replace(to_replace=[np.nan,'',' '],value=0,inplace=True)
    eomp.replace(to_replace=[np.nan,'',' '],value=0,inplace=True)
    
    return eom,eomp



##############################################################################
import datetime as datetime
from datetime import timedelta
YTD=[]
cur_month=todays_date.month
if cur_month in [1,2,3]:
    for i in [1,2,3]:
        YTD.append(i)
        
while cur_month>=4:
    YTD.append(cur_month)
    cur_month=cur_month-1
ytd=[]
if 1 in YTD:
    for i in YTD:
        if i in [1,2,3]:
            ytd.append(str(todays_date.year)+"-"+str(i))
        else:
            ytd.append(str(todays_date.year-1)+"-"+str(i))
else:
    for i in YTD:
        ytd.append(str(todays_date.year)+"-"+str(i))
MTD=[(str(todays_date.year)+"-"+str(todays_date.month))]    
cal_day=datetime.datetime.now()-timedelta(1)
bussdate.set_index('CAL_DATE',inplace=True)
day=bussdate.loc[cal_day.date(),'BUSINESS_DATE']

day=datetime.datetime(day.year,day.month,day.day,0,0,0,0)
eom,eomp=end_of(MTD,eomfin,eomdisb_1)
eoy,eoyp=end_of(ytd,eomfin,eomdisb_1)
eod,eodp=end_of(day,eomfin,eomdisb_1)
mtd_smry=eom[['LOG_VOL', 'REQUESTED_AMOUNT', 'SANCTION_VOL', 'SANCTION_AMOUNT','BOOK_VOL', 'BOOKING_AMOUNT', 'Disbursement_Volume','DISBURSEMENT_AMOUNT', 'Premium','Penetration','ROI']]
mtdp_smry=eomp[['LOG_VOL', 'REQUESTED_AMOUNT', 'SANCTION_VOL', 'SANCTION_AMOUNT','BOOK_VOL', 'BOOKING_AMOUNT', 'Disbursement_Volume','DISBURSEMENT_AMOUNT','Premium','Penetration','ROI']]
ytd_smry=eoy[['LOG_VOL', 'REQUESTED_AMOUNT', 'SANCTION_VOL', 'SANCTION_AMOUNT','BOOK_VOL', 'BOOKING_AMOUNT', 'Disbursement_Volume','DISBURSEMENT_AMOUNT', 'Premium','Penetration','ROI']]
ytdp_smry=eoyp[['LOG_VOL', 'REQUESTED_AMOUNT', 'SANCTION_VOL', 'SANCTION_AMOUNT','BOOK_VOL', 'BOOKING_AMOUNT', 'Disbursement_Volume','DISBURSEMENT_AMOUNT','Premium','Penetration','ROI']]
ftd_smry=eod[['LOG_VOL', 'REQUESTED_AMOUNT', 'SANCTION_VOL', 'SANCTION_AMOUNT','BOOK_VOL', 'BOOKING_AMOUNT', 'Disbursement_Volume','DISBURSEMENT_AMOUNT', 'Premium','Penetration','ROI']]
ftdp_smry=eodp[['LOG_VOL', 'REQUESTED_AMOUNT', 'SANCTION_VOL', 'SANCTION_AMOUNT','BOOK_VOL', 'BOOKING_AMOUNT', 'Disbursement_Volume','DISBURSEMENT_AMOUNT','Premium','Penetration','ROI']]



#######################AUM#####################
def  get_aum(base_view):
    aum=base_view[['REPORTING_BRANCH','FINTYPE','PRINCIPAL_OUTSTANDING','GPLFLAG_SANCTIONS']]
    aum_lap=aum[aum['FINTYPE'].isin(['LP','NP'])]
    aum_gpl=aum[aum['GPLFLAG_SANCTIONS']=='GPL']
    aum_ngpl=aum[aum['GPLFLAG_SANCTIONS']=='NON GPL']
    aum_lap_b=aum_lap.groupby('REPORTING_BRANCH').sum('PRINCIPAL_OUTSTANDING')
    aum_gpl_b=aum_gpl.groupby('REPORTING_BRANCH').sum('PRINCIPAL_OUTSTANDING')
    aum_ngpl_b=aum_ngpl.groupby('REPORTING_BRANCH').sum('PRINCIPAL_OUTSTANDING')
    tot_aum=aum.groupby('REPORTING_BRANCH').sum('PRINCIPAL_OUTSTANDING')
    
    
    monbr=list(set(list(base_view.REPORTING_BRANCH.unique()))-set(aum_gpl_b.index.to_list()))
    for i in monbr:
        aum_gpl_b.loc[i,:]=0
    monbr=list(set(list(base_view.REPORTING_BRANCH.unique()))-set(aum_ngpl_b.index.to_list()))
    for i in monbr:
        aum_ngpl_b.loc[i,:]=0

    monbr=list(set(list(base_view.REPORTING_BRANCH.unique()))-set(aum_lap_b.index.to_list()))
    for i in monbr:
        aum_lap_b.loc[i,:]=0
    monbr=list(set(list(base_view.REPORTING_BRANCH.unique()))-set(tot_aum.index.to_list()))
    for i in monbr:
        tot_aum.loc[i,:]=0
        
        
    aum_gpl_b=aum_gpl_b.sort_values(by='REPORTING_BRANCH')
    aum_gpl_b.loc['Total']= aum_gpl_b.sum()
    aum_ngpl_b=aum_ngpl_b.sort_values(by='REPORTING_BRANCH')
    aum_ngpl_b.loc['Total']= aum_ngpl_b.sum()
    aum_lap_b=aum_lap_b.sort_values(by='REPORTING_BRANCH')
    aum_lap_b.loc['Total']= aum_lap_b.sum()
    tot_aum=tot_aum.sort_values(by='REPORTING_BRANCH')
    tot_aum.loc['Total']= tot_aum.sum()
    tot_aum['PRINCIPAL_OUTSTANDING']=tot_aum['PRINCIPAL_OUTSTANDING']/10000000
    aum_lap_b['PRINCIPAL_OUTSTANDING']=aum_lap_b['PRINCIPAL_OUTSTANDING']/10000000
    aum_gpl_b['PRINCIPAL_OUTSTANDING']=aum_gpl_b['PRINCIPAL_OUTSTANDING']/10000000
    aum_ngpl_b['PRINCIPAL_OUTSTANDING']=aum_ngpl_b['PRINCIPAL_OUTSTANDING']/10000000
    tot_aum.rename(columns={'PRINCIPAL_OUTSTANDING':'TOTAL_AUM'},inplace=True)
    aum_lap_b.rename(columns={'PRINCIPAL_OUTSTANDING':'LAP_AUM'},inplace=True)
    aum_gpl_b.rename(columns={'PRINCIPAL_OUTSTANDING':'GPL_AUM'},inplace=True)
    aum_ngpl_b.rename(columns={'PRINCIPAL_OUTSTANDING':'NON GPL_AUM'},inplace=True)
    df=pd.merge(aum_lap_b,aum_gpl_b, on='REPORTING_BRANCH',how='inner')
    df=df.merge(aum_ngpl_b, on='REPORTING_BRANCH',how='inner')
    df=df.merge(tot_aum, on='REPORTING_BRANCH',how='inner')
    df=df.round(1)
    
    return df
aum=get_aum(base_view.copy())

################################## SUMMARY SHEET COMPLETE ###############################

################################## BRANCH DASHBOARD ###########################


mtd_vol=eom[['LOG_VOL','SANCTION_VOL','BOOK_VOL','Disbursement_Volume','Premium']]
ytd_vol=eoy[['LOG_VOL','SANCTION_VOL','BOOK_VOL','Disbursement_Volume','Premium']]
ftd_vol=eod[['LOG_VOL','SANCTION_VOL','BOOK_VOL','Disbursement_Volume','Premium']]


mtd_val=eom[['REQUESTED_AMOUNT','SANCTION_AMOUNT','BOOKING_AMOUNT','DISBURSEMENT_AMOUNT', 'Premium', 'ROI']]
ytd_val=eoy[['REQUESTED_AMOUNT','SANCTION_AMOUNT','BOOKING_AMOUNT','DISBURSEMENT_AMOUNT', 'Premium', 'ROI']]
ftd_val=eod[['REQUESTED_AMOUNT','SANCTION_AMOUNT','BOOKING_AMOUNT','DISBURSEMENT_AMOUNT', 'Premium', 'ROI']]


branch_aum=aum[['TOTAL_AUM']]
############################### BRANCH DASHBOARD COMPLETE #################### 
################################## PRODUCT DASHBOARD ###########################


mtd_vol_p=eomp[['LOG_VOL','SANCTION_VOL','BOOK_VOL','Disbursement_Volume','Premium']]
ytd_volp=eoyp[['LOG_VOL','SANCTION_VOL','BOOK_VOL','Disbursement_Volume','Premium']]
ftd_volp=eodp[['LOG_VOL','SANCTION_VOL','BOOK_VOL','Disbursement_Volume','Premium']]


mtd_valp=eomp[['REQUESTED_AMOUNT','SANCTION_AMOUNT','BOOKING_AMOUNT','DISBURSEMENT_AMOUNT', 'Premium', 'ROI']]
ytd_valp=eoyp[['REQUESTED_AMOUNT','SANCTION_AMOUNT','BOOKING_AMOUNT','DISBURSEMENT_AMOUNT', 'Premium', 'ROI']]
ftd_valp=eodp[['REQUESTED_AMOUNT','SANCTION_AMOUNT','BOOKING_AMOUNT','DISBURSEMENT_AMOUNT', 'Premium', 'ROI']]

############################### PRODUCT DASHBOARD COMPLETE ####################

############################### MTD FUNNEL ###################################


loan_type={'HOMELOAN':['HL'], 'UNSECURED':['FL', 'FT'], 'BALANCE TRANSFER': ['HT','LT'], 'LAP':['LP'],'NRP': ['NP']}
df_type={}
for key,value in loan_type.items(): 
    fmtd=eomfin.copy()
    fmtd=fmtd[(fmtd['LOGIN_YEAR_MONTH'].isin(MTD)) & (fmtd['FINTYPE'].isin(value))]
    b,p=end_of(MTD, fmtd, eomdisb_1)
    df_type[key]=b

############################### MTD FUNNEL COMPLETED #########################

############################### INCOME SANCTION FLOW #########################
sanc=eomfin.copy()
isflow=sanc[[ 'REPORTING_BRANCH', 'SANCTION_AMOUNT','FINTYPE','IS_YEAR_MONTH','STATUS_SEG']]
isflow['IN_SANCTION_VOL']= np.where(isflow['STATUS_SEG']=='C) Income Sanction',1,0)
isflow=isflow[isflow['IS_YEAR_MONTH'].isin(['1900-1','2020-11','2021-1', '2021-2','2021-3','2020-12'])==False]

def sanctionflow(flows,MTD):
    isflow=flows.copy()
    isflow=isflow[isflow['FINTYPE']=='HL']
    flow={}
    for i in isflow['IS_YEAR_MONTH'].unique():
        flowis=isflow.copy()
        flowis=flowis[flowis['IS_YEAR_MONTH']==i]
        flowis=flowis[['REPORTING_BRANCH','SANCTION_AMOUNT','IN_SANCTION_VOL']]
        flowis=flowis.groupby('REPORTING_BRANCH').sum('SANCTION_AMOUNT','IN_SANCTION_VOL')
        monbr=list(set(list(base_view.REPORTING_BRANCH.unique()))-set(flowis.index.to_list()))
        for j in monbr:
            flowis.loc[j,:]=0
        flowis=flowis.sort_values(by='REPORTING_BRANCH')
        flowis.loc['Total']=flowis.sum()
        flow[i]=flowis
    a=['REPORTING_BRANCH','2021-4','2021-5','2021-6','2021-7','2021-8','2021-9','2021-10','2021-11','2021-12','2022-1','2022-2','2022-3','2022-4','2022-5','2022-6','2022-7','2022-8','2022-9','2022-10','2022-11']
    a=a+ list(set(list(flow.keys()))-set(a))
    sancflow_val=pd.DataFrame(columns=a)
    sancflow_vol=pd.DataFrame(columns=a)
    for x in list(base_view.REPORTING_BRANCH.unique()):
        print(x)
        sancflow_val.loc[x,'REPORTING_BRANCH']=x
        sancflow_vol.loc[x,'REPORTING_BRANCH']=x
    sancflow_val.set_index('REPORTING_BRANCH',inplace=True)
    sancflow_vol.set_index('REPORTING_BRANCH',inplace=True)
    for i in sancflow_val.columns:
        for j in base_view.REPORTING_BRANCH.unique():
            sancflow_val.loc[j,i]=flow[i].loc[j,'SANCTION_AMOUNT']
            sancflow_vol.loc[j,i]=flow[i].loc[j,'IN_SANCTION_VOL']
            
    return sancflow_val,sancflow_vol

sanc_val,sanc_vol=sanctionflow(isflow,MTD)

########################### SANCTION FLOW COMPLETED ##########################
########################### AUM PRODUCT TREND ###############################

def  get_aum_prod(base_view):
    aum_tdf=base_view[['REPORTING_BRANCH','FINTYPE','PRINCIPAL_OUTSTANDING','LOAN_STATUS']]
    aum_tdf=aum_tdf[aum_tdf['LOAN_STATUS']=='Active']
    aum_tdf['PRINCIPAL_OUTSTANDING']=aum_tdf['PRINCIPAL_OUTSTANDING']/10000000
    aum_prod={}
    aum_prod['Total']=aum_tdf.groupby('REPORTING_BRANCH').sum('PRINCIPAL_OUTSTANDING')
    aum_prod['Total'].loc['Total']= aum_prod['Total'].sum()
    for i in aum_tdf.FINTYPE.unique():
        aum=aum_tdf.copy()
        aum=aum[aum['FINTYPE'].isin([i])]
        aum_df=aum.groupby('REPORTING_BRANCH').sum('PRINCIPAL_OUTSTANDING')
        monbr=list(set(list(base_view.REPORTING_BRANCH.unique()))-set(aum_df.index.to_list()))
        for z in monbr:
            aum_df.loc[z,:]=0
        aum_df=aum_df.sort_values(by='REPORTING_BRANCH')
        aum_df.loc['Total']= aum_df.sum()
        aum_df.rename(columns={'PRINCIPAL_OUTSTANDING':i},inplace=True)
        aum_prod[i]=aum_df
        
    dfs=aum_prod.values()
    merged_df = reduce(lambda l, r: pd.merge(l, r, on='REPORTING_BRANCH', how='inner'), dfs)
    merged_df=merged_df.round()
    return merged_df
aum_product_trend=get_aum_prod(base_view.copy())

############################ AUM PRODUCT TREND COMPLETED#######################

############################ TRANCH DISBURSEMENT #############################
def disb_tranch(disb,eomfin,t):
    if t==1:
        disb=disb[disb['DISBURSEMENT_SEQUENCE']==1]
    elif t==0:
        disb=disb[disb['DISBURSEMENT_SEQUENCE']!=1]
    else:
        disb=disb
    
    disbursemnt={}
    
    # timedelta() gets successive dates with
    # appropriate difference
    dates_2020 = [ datetime.date(2020, 4, 1) + datetime.timedelta(days=idx) for idx in range(365)]
    dates_2021= [ datetime.date(2021, 4, 1) + datetime.timedelta(days=idx) for idx in range(365)]
    
    year=[]
    for i in ytd:
        year.append([i])
    for j in year:
        eomdisb_1=disb.copy()
        eomdisb=pd.merge(eomdisb_1,eomfin,on=('LAN_ID'),how='left')
        eomdisb=eomdisb[eomdisb['DISB_YEAR_MONTH'].isin(j)]
        # eomdisb['DISB_VOL']=np.where((eomdisb['DISB_YEAR_MONTH'].isin(MTD)),1,0)
        eomdisb['DISB_VOL']=np.where(True,1,0)
        eomdisb['DISBURSEMENT_AMOUNT']=eomdisb['DISBURSEMENT_AMOUNT']/1000000000
        eomdisb['DISBURSEMENT_AMOUNT']=eomdisb['DISBURSEMENT_AMOUNT'].round(2)
        eomdisb=eomdisb[[ 'REPORTING_BRANCH', 'DISBURSEMENT_AMOUNT','FINTYPE']]
        eomdisbview=eomdisb.groupby('REPORTING_BRANCH').sum(['DISBURSEMENT_AMOUNT'])
        monbr=list(set(list(base_view.REPORTING_BRANCH.unique()))-set(eomdisbview.index.to_list()))
        for z in monbr:
            eomdisbview.loc[z,:]=0
        eomdisbview=eomdisbview.sort_values(by='REPORTING_BRANCH')
        eomdisbview.loc['Total']= eomdisbview.sum()
        eomdisbview.rename(columns={'DISBURSEMENT_AMOUNT':j[0]},inplace=True)
        disbursemnt[j[0]]=eomdisbview
    eomdisb_1=disb.copy()
    eomdisb=pd.merge(eomdisb_1,eomfin,on=('LAN_ID'),how='left')
    eomdisb['DISBDATE']=np.where(True,pd.to_datetime(eomdisb['DISBURSEMENT_DATE']).dt.date ,0)
    eomdisb=eomdisb[eomdisb['DISBDATE'].isin(dates_2020)]
    # eomdisb['DISB_VOL']=np.where((eomdisb['DISB_YEAR_MONTH'].isin(MTD)),1,0)
    eomdisb['DISB_VOL']=np.where(True,1,0)
    eomdisb['DISBURSEMENT_AMOUNT']=eomdisb['DISBURSEMENT_AMOUNT']/1000000000
    eomdisb['DISBURSEMENT_AMOUNT']=eomdisb['DISBURSEMENT_AMOUNT'].round(2)
    eomdisb=eomdisb[[ 'REPORTING_BRANCH', 'DISBURSEMENT_AMOUNT','FINTYPE']]
    eomdisbview=eomdisb.groupby('REPORTING_BRANCH').sum(['DISBURSEMENT_AMOUNT'])
    monbr=list(set(list(base_view.REPORTING_BRANCH.unique()))-set(eomdisbview.index.to_list()))
    for z in monbr:
        eomdisbview.loc[z,:]=0
    eomdisbview=eomdisbview.sort_values(by='REPORTING_BRANCH')
    eomdisbview.loc['Total']= eomdisbview.sum()
    eomdisbview.rename(columns={'DISBURSEMENT_AMOUNT':2020},inplace=True)
    disbursemnt['2020']=eomdisbview
    
    eomdisb_1=disb.copy()
    eomdisb=pd.merge(eomdisb_1,eomfin,on=('LAN_ID'),how='left')
    eomdisb['DISBDATE']=np.where(True,pd.to_datetime(eomdisb['DISBURSEMENT_DATE']).dt.date ,0)
    eomdisb=eomdisb[eomdisb['DISBDATE'].isin(dates_2021)]
    # eomdisb['DISB_VOL']=np.where((eomdisb['DISB_YEAR_MONTH'].isin(MTD)),1,0)
    eomdisb['DISB_VOL']=np.where(True,1,0)
    eomdisb['DISBURSEMENT_AMOUNT']=eomdisb['DISBURSEMENT_AMOUNT']/1000000000
    eomdisb['DISBURSEMENT_AMOUNT']=eomdisb['DISBURSEMENT_AMOUNT'].round(2)
    eomdisb=eomdisb[[ 'REPORTING_BRANCH', 'DISBURSEMENT_AMOUNT','FINTYPE']]
    eomdisbview=eomdisb.groupby('REPORTING_BRANCH').sum(['DISBURSEMENT_AMOUNT'])
    monbr=list(set(list(base_view.REPORTING_BRANCH.unique()))-set(eomdisbview.index.to_list()))
    for z in monbr:
        eomdisbview.loc[z,:]=0
    eomdisbview=eomdisbview.sort_values(by='REPORTING_BRANCH')
    eomdisbview.loc['Total']= eomdisbview.sum()
    eomdisbview.rename(columns={'DISBURSEMENT_AMOUNT':2021},inplace=True)
    disbursemnt['2021']=eomdisbview
    
    dfs=disbursemnt.values()
    merged_df = reduce(lambda l, r: pd.merge(l, r, on='REPORTING_BRANCH', how='inner'), dfs)
    total_disb=merged_df.round()
    return total_disb
tranch1=disb_tranch(disb,eomfin,1)
tranch0=disb_tranch(disb,eomfin,0)
tranchall=disb_tranch(disb,eomfin,2)
############################ TRANCH DISBURSEMENT COMPLETE #############################