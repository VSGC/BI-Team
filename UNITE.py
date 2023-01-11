# -*- coding: utf-8 -*-
"""
Created on Wed Nov  2 14:32:17 2022

@author: VAIBHAV.SRIVASTAV01
"""

import pandas as pd
import calendar
import os
import datetime  
from datetime import timedelta
import numpy as np
from datetime import date
import pyodbc
pyodbc.drivers()
from pptx.util import Inches
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx import Presentation
config = {}
with open(r"C:\Users\VAIBHAV.SRIVASTAV01\Documents\config.txt","r") as file:
    for line in file:
        key, value = line.strip().split(",")
        config[key] = value

del key,line,value,file

conx = pyodbc.connect('DRIVER={SnowflakeDSIIDriver};SERVER='+str(config['server'])+';Warehouse = COMPUTE_WH;DATABASE='+str(config['db1'])+';UID='+ str(config['user1']) +';PWD='+ str(config['password']),autocommit=True)

cur = conx.cursor()


cur.execute("USE DATABASE SAMPLE_DB")

query = "SELECT * FROM DTCRON_MASTER_04122022;"
mydata = pd.read_sql(query, conx)
# insurance=pd.read_sql("SELECT * FROM MIS_INSURANCE_V",conx)
# insurance=insurance[['LANID','INSURANCE_TYPE']]
# mydata=pd.merge(dtcron,insurance,left_on=['FINREFERENCE'],right_on=['LANID'], how='outer')
# mydata.drop(columns='LANID',inplace=True)
# mydata['INSURANCE_TYPE'].replace(to_replace=[np.nan,'',' '],value="0",inplace=True)
ins='select distinct LANID,VASREFERECE,NET_PREMIUM,INSURANCE_CODE,INSURANCE_TYPE,LOAN_ACTIVE_STATUS from MIS_INSURANCE_V union select distinct LANID,VASREFERECE,NET_PREMIUM,INSURANCE_CODE,INSURANCE_TYPE,LOAN_ACTIVE_STATUS from GFL_PLF.PUBLIC.MIS_INSURANCE_V;'
insurance=pd.read_sql(ins,conx)
insurance=insurance[['LANID','NET_PREMIUM','INSURANCE_TYPE']]
insurance.rename(columns={'LANID':'FINREFERENCE'},inplace=True)
# mydata=mydata.merge(insurance,on='FINREFERENCE',how='left')
mydata
# # prs = Presentation()
# # prs.slide_width = 11887200
# # prs.slide_height = 6686550
# # def create_slide(prs,lap_metrics,title):
todays_date = date.today()  
iname=[0,
  'ABH Group Active Secure (PA)',
  'ABH Group Active Secure - co borrower 1 (PA)',
  'ABH Group Protect Secure (Cancer & Heart)',
  'ABH Group Protect Secure-CoBorrower1 Cancer,Heart',
  'ABH Heart secure (Group Active Secure)',
  'ABH Personal Accident Active Secure',
  'ABSLI - Group Asset Assure Plan / ABSLI - GAAP',
  'ABSLI GSS Level Borrower 1',
  'ABSLI GSS Level Borrower 2',
  'ABSLI GSS Reducing Borrower 1',
  'ABSLI GSS Reducing Borrower 2',
  'Aditya Birla health Group Heart Secure',
  'BAGIC Bharat Grah Raksha Borrower 2',
  'BAGIC Bharat Grah Raksha Borrower1',
  'BAGIC Bharat Grah Raksha Borrower2',
  'BAGIC Credit Linked Health Plan Borrower 1',
  'BAGIC Credit Linked Health Plan Borrower 2',
  'New Plan of GCSPlus without Critical Illness',
  'TAGIC - Group Credit Secure Plus / TAGIC - GCS+',
  'TAGIC - Group MediCare',
  'TAGIC - Property Insurance',
  'HDFC Life GCPP Non STP - Borrower 1',
  'HDFC Life GCPP STP - Borrower 1']
ipercent=[0.0,
0.65,
0.65,
0.65,
0.65,
0.65,
0.65,
0.6000000000000001,
0.6000000000000001,
0.55,
0.6000000000000001,
0.55,
0.65,
0.0,
0.35,
0.35,
0.5,
0.5,
0.6,
0.6000000000000001,
0.6000000000000001,
0.4,
0.58,
0.58]
insur_dc=dict(zip(iname,ipercent))
insurance['InsPercent']=0
for i in insur_dc.keys():
    insurance['InsPercent']=np.where(insurance['INSURANCE_TYPE']==i,insur_dc[i],insurance['InsPercent'])
insurance['Net insurance income']=insurance['NET_PREMIUM']*insurance['InsPercent']*0.9/1.18
insurance=insurance.groupby('FINREFERENCE',as_index=False).sum(['Net insurance income'])
insurance=insurance[['FINREFERENCE','Net insurance income']]
mydata=mydata.merge(insurance,on='FINREFERENCE',how='left')

# rework=rework.groupby('FINREFERENCE').max('Rework_Type')

ftr_log=pd.read_excel(r"C:\Users\VAIBHAV.SRIVASTAV01\Desktop\OPS\FTR_LTD_MIS.xlsx",sheet_name='LOGIN_DATA')
# ftr_log=ftr_log[['LOGIN_YEAR_MONTH']=='2022-9']
ftr_log=ftr_log[['FINREFERENCE','BOOKED_AMOUNT','LOGIN_YEAR_MONTH','FTR']]
ftr_log.rename(columns={'LOGIN_YEAR_MONTH':'FTR_LOGIN_YEAR_MONTH'},inplace=True)
ftr_log.rename(columns={'BOOKED_AMOUNT':'FTR_LOG'},inplace=True)
ftr_log.rename(columns={'FTR':'FTR_FLAG'},inplace=True)
ftr_disb=pd.read_excel(r"C:\Users\VAIBHAV.SRIVASTAV01\Desktop\OPS\FTR_LTD_MIS.xlsx",sheet_name='DISBURSEMENT_DATA')
ftr_disb=ftr_disb[['FINREFERENCE','BOOKED_AMOUNT','BOOK_YEAR_MONTH','DOCKET_FTR']]
ftr_disb.rename(columns={'BOOK_YEAR_MONTH':'FTR_BOOK_YEAR_MONTH'},inplace=True)
ftr_disb.rename(columns={'BOOKED_AMOUNT':'FTR_disb'},inplace=True)
mydata=mydata.merge(ftr_log,on='FINREFERENCE',how='left')
mydata=mydata.merge(ftr_disb,on='FINREFERENCE',how='left')



df=mydata.copy()

'''
 ------ LAP ROI ############################
LAP_IND_ROI <- subset(mydata ,mydata$FINTYPE %in% c('LP','NP') & 
                      mydata$STATUS == 'Booked' & mydata$BOOK_YEAR_MONTH == MTD)

 Regular = LAP_Regular_ROI <- subset(mydata, mydata$FINTYPE %in% c('LP','NP') &
                                     !mydata$LOAN_PURPOSE %in% c('LAP Balance Transfer plus Top-up ','LAP Top Up') 
                          & mydata$BOOK_YEAR_MONTH == MTD & mydata$STATUS == 'Booked'
                          & !mydata$SUBPRODUCT == 'BOOSTER')
	Industrial = subset(mydata, mydata$FINTYPE %in% c('LP','NP') & mydata$STATUS == 'Booked' & mydata$BOOK_YEAR_MONTH == MTD)
  
	Booster =  LAP_Booster_ROI <- subset(mydata, mydata$FINTYPE %in% c('LP','NP') & 
                                      mydata$SUBPRODUCT == 'BOOSTER' 
                          & mydata$BOOK_YEAR_MONTH == MTD & mydata$STATUS == 'Booked' )

	Topup = subset(mydata, mydata$FINTYPE %in% c('LP','NP') & mydata$STATUS == 'Booked' 
                        & mydata$BOOK_YEAR_MONTH == MTD
                        & mydata$INDUSTRIAL_PROPERTY_FLAG == 0
                        & mydata$LOAN_PURPOSE %in% c('LAP Balance Transfer plus Top-up ','LAP Top Up'))

	Total = LAP_Total_ROI <- subset(mydata, mydata$FINTYPE %in% c('LP','NP') &
                                 mydata$STATUS == 'Booked'
                        & mydata$BOOK_YEAR_MONTH == MTD )
'''



lap_df = df[['FTR_FLAG','DOCKET_FTR','FTR_disb','FTR_LOG','FTR_BOOK_YEAR_MONTH','FTR_LOGIN_YEAR_MONTH', 'FINREFERENCE','BOOK_YEAR_MONTH','FINBRANCH','FINTYPE','STATUS', 'WROI', 'CRE','STEPFINANCE', 
             'GPLFLAG_SANCTIONS','BOOKED_AMOUNT' ,'SUBPRODUCT','LOAN_PURPOSE','INDUSTRIAL_PROPERTY_FLAG','PROCESSING_FEE','NET_PREMIUM','PRINCIPAL_OUTSTANDING','Net insurance income','LOGINSTATUS','LOGIN_YEAR_MONTH','REJECT_YEAR_MONTH','REQLOANAMT']]
lap_df_roi = lap_df.copy()
lap_df_roi = lap_df_roi[lap_df_roi.GPLFLAG_SANCTIONS == 'NIL']
lap_df_roi = lap_df_roi[(lap_df_roi['FINTYPE'] == 'LP') | (lap_df_roi['FINTYPE'] == 'NP') ]
# lap_df_roi = lap_df_roi[lap_df_roi['BOOK_YEAR_MONTH'].notna()]

#


lap_df_roi['Lap_neo']=np.where( (lap_df_roi['SUBPRODUCT'] == 'NEO LAP')|(lap_df_roi['SUBPRODUCT'] == 'NEO BOOSTER LAP')|(lap_df_roi['SUBPRODUCT'] == 'NEO NRP'), 1,0)
lap_df_roi['Lap_booster'] = np.where( (lap_df_roi['SUBPRODUCT'] == 'BOOSTER'), 1,0)
lap_df_roi['Lap_lrd'] = np.where( (lap_df_roi['SUBPRODUCT'] == 'LRD') , 1,0)
lap_df_roi['Lap_topup'] = np.where((lap_df_roi['Lap_neo']==0)&(lap_df_roi['Lap_booster']==0 )&(lap_df_roi['Lap_lrd'] ==0 )&(lap_df_roi['INDUSTRIAL_PROPERTY_FLAG'] == 0) & ((lap_df_roi['LOAN_PURPOSE']=='LAP Balance Transfer plus Top-up ')|(lap_df_roi['LOAN_PURPOSE']=='LAP Top Up')) , 1,0 )
lap_df_roi['Lap_industrial'] = np.where((lap_df_roi['INDUSTRIAL_PROPERTY_FLAG'] == 1)&(lap_df_roi['Lap_booster']==0 )&(lap_df_roi['Lap_topup'] ==0 ) & (lap_df_roi['Lap_neo']==0) & (lap_df_roi['Lap_lrd'] ==0 ), 1,0 )
lap_df_roi['Lap_regular'] = np.where( (lap_df_roi['Lap_neo']==0)& (lap_df_roi['Lap_industrial'] ==0 ) &  (lap_df_roi['Lap_booster']==0 )& 
                                  (lap_df_roi['Lap_topup'] ==0 )&(lap_df_roi['Lap_lrd'] ==0 ) , 1,0 )


lap_df_roi['Lap_total'] = np.where( (lap_df_roi['FINTYPE'] == 'LP') |  (lap_df_roi['FINTYPE'] =='NP') , 1,0 )
lap_df_roi.to_excel(r"C:\Users\VAIBHAV.SRIVASTAV01\Desktop\lap_productwise.xlsx")
# MTD=['2022-11','2022-10','2022-9','2022-8','2022-7','2022-6','2022-5','2022-4']
MTD=['2022-11']
def bookings(prod_name,MTD):
    df_roi=lap_df_roi.copy()
    if type(MTD)==str:
        prod=df_roi[(df_roi[prod_name]==1)&(df_roi['BOOK_YEAR_MONTH']==MTD)&(df_roi['STATUS']=='Booked')]
    elif type(MTD)==list :
        prod=df_roi[(df_roi[prod_name]==1)&(df_roi['BOOK_YEAR_MONTH'].isin(MTD))&(df_roi['STATUS']=='Booked')]   
    prod['BOOKINGS']=1
    prod.drop(columns = ['FTR_disb','FTR_FLAG','DOCKET_FTR',
     'FTR_LOG',
     'FTR_BOOK_YEAR_MONTH',
     'FTR_LOGIN_YEAR_MONTH',
     'STATUS',
     'BOOK_YEAR_MONTH',
     'Lap_regular',
     'Lap_industrial',
     'Lap_neo',
     'Lap_booster',
     'Lap_topup',
     'Lap_lrd',
     'Lap_total',
     'LOGINSTATUS',
     'LOGIN_YEAR_MONTH',
     'REJECT_YEAR_MONTH',
     'REQLOANAMT'],inplace=True)
    prod_branch1=prod.groupby('FINBRANCH',as_index=False).sum(['WROI','PROCESSING_FEE','NET_PREMIUM', 'BOOKED_AMOUNT','PRINCIPAL_OUTSTANDING','Net insurance income','BOOKINGS'])
    prod_branch1.drop(columns = ['PRINCIPAL_OUTSTANDING'],inplace=True)
    branches=lap_df_roi['FINBRANCH'].unique()
    branches=list(branches)
    branches=list(set(branches)-set(prod_branch1['FINBRANCH']))
    a=pd.DataFrame(branches)
    a.rename(columns={0:'FINBRANCH'},inplace=True)
    prod_branch1=prod_branch1.append(a)
    prod_branch1.replace(to_replace=[np.nan,'',' '],value=0,inplace=True)
      
    # prod_branch1.loc['Total']=prod_branch1.sum()
    return(prod_branch1)

def pos(prod_name,MTD):
    df_roi=lap_df_roi.copy()
    prod2= df_roi[(df_roi[prod_name]==1)]
    prod2.drop(columns = ['FTR_disb','FTR_FLAG','DOCKET_FTR',
 'FTR_LOG',
 'FTR_BOOK_YEAR_MONTH',
 'FTR_LOGIN_YEAR_MONTH',
 'STATUS',
 'BOOK_YEAR_MONTH',
 'WROI',
 'PROCESSING_FEE',
 'NET_PREMIUM',
 'BOOKED_AMOUNT',
 'Lap_regular',
 'Lap_neo',
 'Lap_industrial',
 'Lap_booster',
 'Lap_topup',
 'Lap_lrd',
 'Lap_total',
 'Net insurance income',
 'LOGINSTATUS',
 'LOGIN_YEAR_MONTH',
 'REJECT_YEAR_MONTH',
 'REQLOANAMT',
 'STEPFINANCE',
 'INDUSTRIAL_PROPERTY_FLAG'],inplace=True)
    prod_pos=prod2.groupby('FINBRANCH',as_index=False).sum(['PRINCIPAL_OUTSTANDING'])
    branches=lap_df_roi['FINBRANCH'].unique()
    branches=list(branches)
    branches=list(set(branches)-set(prod_pos['FINBRANCH']))
    a=pd.DataFrame(branches)
    a.rename(columns={0:'FINBRANCH'},inplace=True)
    prod_pos=prod_pos.append(a)
    prod_pos.replace(to_replace=[np.nan,'',' '],value=0,inplace=True)
    return(prod_pos)
    
def login(prod_name,MTD):

    df_roi=lap_df_roi.copy()
    if type(MTD)==str:
        prod3=df_roi[(df_roi[prod_name]==1)&(df_roi['LOGIN_YEAR_MONTH']==MTD)&(df_roi['LOGINSTATUS']=="A) Login")]
    elif type(MTD)==list :
        prod3=df_roi[(df_roi[prod_name]==1)&(df_roi['LOGIN_YEAR_MONTH'].isin(MTD))&(df_roi['LOGINSTATUS']=="A) Login")]
    prod3['LOGINS'] =1   
    prod3.drop(columns = ['FTR_disb','FTR_FLAG','DOCKET_FTR',
 'FTR_LOG',
 'FTR_BOOK_YEAR_MONTH',
 'FTR_LOGIN_YEAR_MONTH',
 'STATUS',
 'BOOK_YEAR_MONTH',
 'WROI',
 'PROCESSING_FEE',
 'NET_PREMIUM',
 'BOOKED_AMOUNT',
 'Lap_regular',
 'Lap_neo',
 'Lap_industrial',
 'Lap_booster',
 'Lap_topup',
 'Lap_lrd',
 'Lap_total',
 'PRINCIPAL_OUTSTANDING',
 'Net insurance income',
 'LOGINSTATUS',
 'LOGIN_YEAR_MONTH',
 'REJECT_YEAR_MONTH','STEPFINANCE',
 'INDUSTRIAL_PROPERTY_FLAG'],inplace=True)
    prod3.rename(columns = {'REQLOANAMT':'LOGIN_VALUE'}, inplace = True)
    prod_log=prod3.groupby('FINBRANCH',as_index=False).sum(['LOGIN_VALUE','LOGINS'])
    branches=lap_df_roi['FINBRANCH'].unique()
    branches=list(branches)
    branches=list(set(branches)-set(prod_log['FINBRANCH']))
    a=pd.DataFrame(branches)
    a.rename(columns={0:'FINBRANCH'},inplace=True)
    prod_log=prod_log.append(a)
    prod_log.replace(to_replace=[np.nan,'',' '],value=0,inplace=True)
    return(prod_log)

def rejects(prod_name,MTD):
    df_roi=lap_df_roi.copy()
    if type(MTD)==str:
        prod4=df_roi[(df_roi[prod_name]==1)&(df_roi['REJECT_YEAR_MONTH']==MTD)&(df_roi['STATUS']=="Rejected")]
    elif type(MTD)==list :
        prod4=df_roi[(df_roi[prod_name]==1)&(df_roi['REJECT_YEAR_MONTH'].isin(MTD))&(df_roi['STATUS']=="Rejected")]
    prod4['REJECTS']=1 
    prod4.rename(columns = {'REQLOANAMT':'REJECT_VALUE'}, inplace = True)
    prod4.drop(columns = ['FTR_FLAG','DOCKET_FTR','PRINCIPAL_OUTSTANDING','WROI','PROCESSING_FEE','NET_PREMIUM', 'Lap_regular','Lap_neo','Lap_industrial','Lap_booster','Lap_topup','Lap_lrd','Lap_total','BOOKED_AMOUNT','Net insurance income','LOGINSTATUS','LOGIN_YEAR_MONTH','REJECT_YEAR_MONTH','FTR_disb','FTR_LOG','FTR_BOOK_YEAR_MONTH','FTR_LOGIN_YEAR_MONTH','STEPFINANCE',
 'INDUSTRIAL_PROPERTY_FLAG'],inplace=True)
    prod_reject=prod4.groupby('FINBRANCH',as_index=False).sum(['REJECT_VALUE','REJECTS'])
    branches=lap_df_roi['FINBRANCH'].unique()
    branches=list(branches)
    branches=list(set(branches)-set(prod_reject['FINBRANCH']))
    a=pd.DataFrame(branches)
    a.rename(columns={0:'FINBRANCH'},inplace=True)
    prod_reject=prod_reject.append(a)
    prod_reject.replace(to_replace=[np.nan,'',' '],value=0,inplace=True)

    return(prod_reject)


def rework(prod_name,MTD):
    df_roi=lap_df_roi.copy()
    prod5=df_roi.copy()
    rework=pd.read_csv(r"C:\Users\VAIBHAV.SRIVASTAV01\Desktop\OPS\REWORK.csv")
    rework=rework[['FINREFERENCE','Rework_Type','take1']]
    rework_first=pd.read_csv(r"C:\Users\VAIBHAV.SRIVASTAV01\Desktop\OPS\LP First Time.csv")
    rework_first=rework_first[['FINREFERENCE','take']]
    prod5=prod5.merge(rework,on='FINREFERENCE',how='left')
    prod5=prod5.merge(rework_first,on='FINREFERENCE',how='left')  
    prod5=prod5[prod5[prod_name]==1]
    prod5.drop(columns = ['FTR_FLAG','DOCKET_FTR','PRINCIPAL_OUTSTANDING','WROI','PROCESSING_FEE','NET_PREMIUM', 'Lap_regular','Lap_neo','Lap_industrial','Lap_booster','Lap_topup','Lap_lrd','Lap_total','BOOKED_AMOUNT','Net insurance income','LOGINSTATUS','LOGIN_YEAR_MONTH','REJECT_YEAR_MONTH','REQLOANAMT','FTR_disb','FTR_LOG','FTR_BOOK_YEAR_MONTH','FTR_LOGIN_YEAR_MONTH','STEPFINANCE',
 'INDUSTRIAL_PROPERTY_FLAG'],inplace=True)
    prod5['Change Terms Rework']=np.where(prod5['Rework_Type']=='Change Terms Rework',1,0)
    prod5['FS Rework']=np.where(prod5['Rework_Type']=='FS Rework',1,0)
    prod5['Reject Rework']=np.where(prod5['Rework_Type']=='Reject Rework',1,0)
    prod_rework=prod5.groupby('FINBRANCH',as_index=False).sum(['take','take1','Reject Rework','FS Rework','Change Terms Rework'])
    branches=lap_df_roi['FINBRANCH'].unique()
    branches=list(branches)
    branches=list(set(branches)-set(prod_rework['FINBRANCH']))
    a=pd.DataFrame(branches)
    a.rename(columns={0:'FINBRANCH'},inplace=True)
    prod_rework=prod_rework.append(a)
    prod_rework.replace(to_replace=[np.nan,'',' '],value=0,inplace=True)
    return(prod_rework)


def ftrlog(prod_name,MTD):
    prod6=lap_df_roi.copy()
    if type(MTD)==str:
        prod6=prod6[prod6['FTR_LOGIN_YEAR_MONTH']==MTD]
    elif type(MTD)==list :
        prod6=prod6[prod6['FTR_LOGIN_YEAR_MONTH'].isin(MTD)]
    prod6=prod6[(prod6[prod_name]==1) & (prod6['FTR_FLAG']=='YES')]
    prod6.drop(columns = ['FTR_FLAG','DOCKET_FTR','PRINCIPAL_OUTSTANDING','WROI','PROCESSING_FEE','NET_PREMIUM', 'Lap_regular','Lap_neo','Lap_industrial','Lap_booster','Lap_topup','Lap_lrd','Lap_total','BOOKED_AMOUNT','Net insurance income','LOGINSTATUS','LOGIN_YEAR_MONTH','REJECT_YEAR_MONTH','REQLOANAMT','FTR_disb','FTR_BOOK_YEAR_MONTH','FTR_LOGIN_YEAR_MONTH','STEPFINANCE',
 'INDUSTRIAL_PROPERTY_FLAG'],inplace=True)
    prod6['FTR_LOG_COUNT']=1
    prod_ftrlog=prod6.groupby('FINBRANCH',as_index=False).sum(['FTR_LOG','FTR_LOG_COUNT'])
    branches=lap_df_roi['FINBRANCH'].unique()
    branches=list(branches)
    branches=list(set(branches)-set(prod_ftrlog['FINBRANCH']))
    a=pd.DataFrame(branches)
    a.rename(columns={0:'FINBRANCH'},inplace=True)
    prod_ftrlog=prod_ftrlog.append(a)
    prod_ftrlog.replace(to_replace=[np.nan,'',' '],value=0,inplace=True)
    return(prod_ftrlog)


def ftrdisb(prod_name,MTD):
    df_roi=lap_df_roi.copy()
    prod7=df_roi.copy()
    if type(MTD)==str:
        prod7=prod7[prod7['FTR_BOOK_YEAR_MONTH']==MTD]
    elif type(MTD)==list :
        prod7=prod7[prod7['FTR_BOOK_YEAR_MONTH'].isin(MTD)]
    prod7=prod7[(prod7[prod_name]==1) & (prod7['DOCKET_FTR']=='YES')]
    prod7.drop(columns = ['FTR_FLAG','DOCKET_FTR','PRINCIPAL_OUTSTANDING','WROI','PROCESSING_FEE','NET_PREMIUM', 'Lap_regular','Lap_neo','Lap_industrial','Lap_booster','Lap_topup','Lap_lrd','Lap_total','BOOKED_AMOUNT','Net insurance income','LOGINSTATUS','LOGIN_YEAR_MONTH','REJECT_YEAR_MONTH','REQLOANAMT','FTR_LOG','FTR_BOOK_YEAR_MONTH','FTR_LOGIN_YEAR_MONTH','STEPFINANCE',
 'INDUSTRIAL_PROPERTY_FLAG'],inplace=True)
    prod7['FTR_BOOK_COUNT']=1
    prod_ftrdisb=prod7.groupby('FINBRANCH',as_index=False).sum(['FTR_disb','FTR_BOOK_COUNT'])
    branches=lap_df_roi['FINBRANCH'].unique()
    branches=list(branches)
    branches=list(set(branches)-set(prod_ftrdisb['FINBRANCH']))
    a=pd.DataFrame(branches)
    a.rename(columns={0:'FINBRANCH'},inplace=True)
    prod_ftrdisb=prod_ftrdisb.append(a)
    prod_ftrdisb.replace(to_replace=[np.nan,'',' '],value=0,inplace=True)
    return(prod_ftrdisb)

def prodwise_merge(prod_name,MTD):
    ftrd=ftrdisb(prod_name,MTD)
    ftrl=ftrlog(prod_name,MTD)
    rework1=rework(prod_name,MTD)
    rej=rejects(prod_name,MTD)
    log=login(prod_name,MTD)
    pos1=pos(prod_name,MTD)
    book=bookings(prod_name,MTD)
    prod=ftrd.merge(ftrl,on='FINBRANCH',how='left')
    prod=prod.merge(rework1,on='FINBRANCH',how='left')
    prod=prod.merge(rej,on='FINBRANCH',how='left')
    prod=prod.merge(log,on='FINBRANCH',how='left')
    prod=prod.merge(pos1,on='FINBRANCH',how='left')
    prod=prod.merge(book,on='FINBRANCH',how='left')
    prod.loc['Country']=prod.sum()
    prod.loc['Country','FINBRANCH']='Country'
    prod.reset_index(drop=True,inplace=True)
    return prod
prod={}

for i in['Lap_regular','Lap_industrial','Lap_booster','Lap_topup','Lap_lrd','Lap_neo','Lap_total']:
    prod[i]=prodwise_merge(i,MTD)

    
def cal_c(df):
    df['ROI']=df.WROI/df.BOOKED_AMOUNT
    df['PF']=100*df.PROCESSING_FEE/df.BOOKED_AMOUNT
    df['GROSS']=100*df.NET_PREMIUM/df.BOOKED_AMOUNT
    df['NET INSURANCE INCOME']=100*(df['Net insurance income']/df.BOOKED_AMOUNT)
for i in['Lap_regular','Lap_industrial','Lap_booster','Lap_topup','Lap_lrd','Lap_neo','Lap_total']:
    cal_c(prod[i])    
def drop_col_lap(df):
    try:
        df.drop(columns = ['WROI','PROCESSING_FEE','NET_PREMIUM', 'Lap_regular','Lap_industrial','Lap_neo','Lap_booster','Lap_topup','Lap_lrd','Lap_total','Net insurance income'],inplace=True)
    except:
        pass
    df.set_index('FINBRANCH',inplace=True)
    return(df)
for i in ['Lap_regular','Lap_industrial','Lap_booster','Lap_topup','Lap_lrd','Lap_neo','Lap_total']:
    prod[i]=drop_col_lap(prod[i])
    
def transpose(df):
    df=df[['ROI','NET INSURANCE INCOME','PF','GROSS', 'LOGINS','LOGIN_VALUE', 'BOOKINGS', 'BOOKED_AMOUNT',
       'PRINCIPAL_OUTSTANDING',  'REJECTS','REJECT_VALUE', 'take','take1',
       'Change Terms Rework', 'FS Rework', 'Reject Rework', 'FTR_LOG',
       'FTR_LOG_COUNT', 'FTR_disb', 'FTR_BOOK_COUNT']]
    df.rename(columns={'take1':'TotalRework'},inplace=True)
    df['ReworkRatio']=df['TotalRework']/df['take']
    df.drop(columns=['take'],inplace=True)
    df['Change Terms Rework']=100*df['Change Terms Rework']/df['TotalRework']
    df['FS Rework']=100*df['FS Rework']/df['TotalRework']
    df['Reject Rework']=100*df['Reject Rework']/df['TotalRework']

    df=df.transpose()
    return df

for i in ['Lap_regular','Lap_industrial','Lap_booster','Lap_topup','Lap_lrd','Lap_neo','Lap_total']:
    prod[i]=transpose(prod[i])
    
lap_bdf={"MTD":['ROI',
'NET INSURANCE INCOME',
'PF',
'GROSS',
'LOGINS',
'LOGIN_VALUE',
'BOOKINGS',
'BOOKED_AMOUNT',
'PRINCIPAL_OUTSTANDING',
'REJECTS',
'REJECT_VALUE',
'ReworkRatio',
'TotalRework',
'Change Terms Rework',
'FS Rework',
'Reject Rework',
'FTR_LOG',
'FTR_LOG_COUNT',
'FTR_disb',
'FTR_BOOK_COUNT',
],
         'LAP REGULAR':['','','','','','','','','','','','','','','','','','','',''],
         'LAP INDUSTRIAL':['','','','','','','','','','','','','','','','','','','',''],
         'LAP BOOSTER':['','','','','','','','','','','','','','','','','','','',''],
         'LAP TOPUP':['','','','','','','','','','','','','','','','','','','',''],
         'LAP LRD':['','','','','','','','','','','','','','','','','','','',''],
         'LAP NEO':['','','','','','','','','','','','','','','','','','','',''],
         'LAP TOTAL':['','','','','','','','','','','','','','','','','','','','']} 
lap_b=pd.DataFrame(data=lap_bdf) 

mdf_br=[prod['Lap_regular'],prod['Lap_industrial'],prod['Lap_booster'],prod['Lap_topup'],prod['Lap_lrd'],prod['Lap_neo'],prod['Lap_total']]
branches=df['FINBRANCH'].unique()
branches=np.append(branches,['Country'])
brnc_dc={}
for i in branches:
    brnc_dc[i]=lap_b.copy()
    brnc_dc[i].set_index('MTD',inplace=True)
    for j in brnc_dc[i].index:
        for z in range(len(brnc_dc[i].columns)):
            if i in (mdf_br[z].columns):
                brnc_dc[i].loc[j,(brnc_dc[i].columns[z])]=(mdf_br[z]).loc[j,i]
            else:
                brnc_dc[i].loc[j,brnc_dc[i].columns[z]]=0
    brnc_dc[i].replace(to_replace=[np.nan,'',' '],value=0,inplace=True)
    
writer = pd.ExcelWriter(r"C:\Users\VAIBHAV.SRIVASTAV01\Desktop\NOV_OPS_REVIEW\SUBPRODUCT_MTD.xlsx",engine='xlsxwriter')
workbook=writer.book
for x in branches:
    worksheet=workbook.add_worksheet(x)
    writer.sheets[x] = worksheet
    worksheet.write_string(0, 0, "MTD")
    
    brnc_dc[x].to_excel(writer,sheet_name=x,startrow=1 , startcol=0)
    # worksheet.write_string(brnc_dc[x].shape[0] + 4, 0, 'YTD')
    # ytd_brnc_dc[x].to_excel(writer,sheet_name=x,startrow=brnc_dc[x].shape[0] + 5, startcol=0)
writer.save()