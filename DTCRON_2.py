# -*- coding: utf-8 -*-
"""
Created on Tue Nov 15 16:27:54 2022

@author: VAIBHAV.SRIVASTAV01
"""

import pandas as pd 
import pyodbc
import sys
import os
import collections

##
con_dm = pyodbc.connect('DSN=GHF_BI_CONN_DM;UID=GHF_BI_CONN;PWD=Godrej@123')
con = pyodbc.connect('DSN=GHF_BI_CONN;UID=GHF_BI_CONN;PWD=Godrej@123')
con_dm_nbfc = pyodbc.connect('DSN=GHF_BI_CONN_DM_NBFC;UID=GHF_BI_CONN;PWD=Godrej@123')

# dtcron=pd.read_sql("SELECT * from dtcron;", con)
ghf_loan_view=pd.read_sql("SELECT * from prod_da_db.serve.loan_view WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm)
ghf_loan_view['NBFC_FLAG'] = 'N'
ghf_base_view = pd.read_sql("SELECT * from prod_da_db.serve.BASE_VIEW WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm)
ghf_base_view['NBFC_FLAG'] = 'N'
ghf_customer_base = pd.read_sql("SELECT * from prod_da_db.serve.customer_base WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm)
ghf_customer_base['NBFC_FLAG'] = 'N'
ghf_insurance_view = pd.read_sql("SELECT * from prod_da_db.serve.INSURANCE_VIEW WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm)
ghf_insurance_view ['NBFC_FLAG'] = 'N'
ghf_lanvas=pd.read_sql("SELECT * from prod_da_db.serve.x_ref_lan_to_vas WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm)
ghf_lanvas['NBFC_FLAG'] = 'N'
ghf_lancif=pd.read_sql("SELECT * from prod_da_db.serve.x_ref_lan_to_cif WHERE DH_RECORD_ACTIVE_FLAG = 'Y' and APPLICANT_TYPE='APPLICANT'", con_dm)
ghf_lancif['NBFC_FLAG'] = 'N'
ghf_cifvas=pd.read_sql("SELECT * from prod_da_db.serve.x_ref_cif_to_vas WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm)
ghf_cifvas['NBFC_FLAG'] = 'N'
ghf_lancollat=pd.read_sql("SELECT * from prod_da_db.serve.x_ref_lan_to_collat WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm)
ghf_lancollat['NBFC_FLAG'] = 'N'

gfl_loan_view=pd.read_sql("SELECT * from prod_gfl_da_db.serve.loan_view WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm_nbfc)
gfl_loan_view['NBFC_FLAG'] = 'Y'
gfl_base_view = pd.read_sql("SELECT * from prod_gfl_da_db.serve.BASE_VIEW WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm_nbfc)
gfl_base_view['NBFC_FLAG'] = 'Y'
gfl_customer_base = pd.read_sql("SELECT * from prod_gfl_da_db.serve.customer_base WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm_nbfc)
gfl_customer_base['NBFC_FLAG'] = 'Y'
gfl_insurance_view = pd.read_sql("SELECT * from prod_gfl_da_db.serve.INSURANCE_VIEW WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm_nbfc)
gfl_insurance_view['NBFC_FLAG'] = 'Y'
gfl_lanvas=pd.read_sql("SELECT * from prod_gfl_da_db.serve.x_ref_lan_to_vas WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm_nbfc)
gfl_lanvas['NBFC_FLAG'] = 'Y'
gfl_lancif=pd.read_sql("SELECT * from prod_gfl_da_db.serve.x_ref_lan_to_cif WHERE DH_RECORD_ACTIVE_FLAG = 'Y' and APPLICANT_TYPE='APPLICANT'", con_dm_nbfc)
gfl_lancif['NBFC_FLAG'] = 'Y'
gfl_cifvas=pd.read_sql("SELECT * from prod_gfl_da_db.serve.x_ref_cif_to_vas WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm_nbfc)
gfl_cifvas['NBFC_FLAG'] = 'Y'
gfl_lancollat=pd.read_sql("SELECT * from prod_gfl_da_db.serve.x_ref_lan_to_collat WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm_nbfc)
gfl_lancollat['NBFC_FLAG'] = 'Y'

loan_view=pd.concat([ghf_loan_view,gfl_loan_view]) #REFERENCE
base_view=pd.concat([ghf_base_view,gfl_base_view]) # LAN_ID
ghf_lancif=ghf_lancif[[ 'CUSTOMER_CIF','LAN_ID','APPLICANT_TYPE']]
ghf_customer_base=ghf_lancif.merge(ghf_customer_base,on = 'CUSTOMER_CIF',how='left')
gfl_lancif=gfl_lancif[[ 'CUSTOMER_CIF','LAN_ID','APPLICANT_TYPE']]
gfl_customer_base=gfl_lancif.merge(gfl_customer_base,on = 'CUSTOMER_CIF',how='left')

customer_view=pd.concat([ghf_customer_base,gfl_customer_base]) # CIF
insurance_view=pd.concat([ghf_insurance_view,gfl_insurance_view]) # LAN_ID
lanvas=pd.concat([ghf_lanvas,gfl_lanvas])
lancif=pd.concat([ghf_lancif,gfl_lancif])
cifvas=pd.concat([ghf_cifvas,gfl_cifvas])
lancollat=pd.concat([ghf_lancollat,gfl_lancollat])
loan_view.rename(columns={'REFERENCE':'LAN_ID'},inplace=True)

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
def find_columns(base_view,dtcron):
    
    base_view=remove_error_lans(base_view)
    dtcron_dict={}
    dt=dtcron.columns.to_list()
    dt.sort()
    for i in dt:
        dtcron_dict[i]=list(dtcron[i].unique())
        
    bv_dict={}
    bv=base_view.columns.to_list()
    bv.sort()
    for i in bv:
        bv_dict[i]=list(base_view[i].unique())
    a=[]
    for j in bv: 
        for i in dt:        
            if len(bv_dict[j])>2 and (collections.Counter(bv_dict[j]) == collections.Counter(dtcron_dict[i]) or set(bv_dict[j]).issubset(set(dtcron_dict[i])) or j==i) :
                a.append(j)
                print(j,i)
                break
    return a         
          
# bv_cols=find_columns(base_view,dtcron)  
# b=['BOOKING_AMOUNT',
#  'EOMCLTRL',
#  'EOMLOGN',
#  'EOMRJCT',
#  'EOMSNCTN',
#  'REQUESTED_AMOUNT',
#  'LOGIN_STATUS']      
# c=list(set(bv_cols)-set(b))
bv_cols=['FIRST_REJECT_DATE',
 'DETAILED_STATUS',
 'SANCTION_AMOUNT',
 'REPORTING_BRANCH',
 'NET_PREMIUM',
 'LAN_ID',
 'QUEUE',
 'LOAN_STATUS',
 'STATUS',
 'STATUS_SEG',
 'FIRST_ENTRY_TO_COLLATERAL',
 'FINANCE_TYPE',
 'LOAN_TYPE',
 'BOOKING_DATE',
 'LOGIN_DATE',
 'BOOKING_AMOUNT',
 'EOMCLTRL',
 'EOMLOGN',
 'EOMRJCT',
 'EOMSNCTN',
 'REQUESTED_AMOUNT','LOGIN_STATUS']

base_view=base_view[bv_cols]      
# lv_cols=find_columns(loan_view,dtcron) 
lv_cols=['BORROWER_TYPE',
 'BT_LOAN_LAN',
 'BT_LOAN_START_DATE',
 'BT_OUTSTANDING',
 'DELAY_REASON',
 'DST_CODE',
 'END_USAGE_FUNDS',
 'END_USE_FOR_TOPUP',
 'FINAL_FOIR',
 'FINAL_LOAN_AMOUNT',
 'FINAL_LTV',
 'FTR',
 'GHFAM',
 'GHFAM_BD',
 'GHFAM_BD_NAME',
 'GHFAM_NAME',
 'GHFSM',
 'GHFSM_NAME',
 'GHF_AM',
 'GRACE_TERMS',
 'INCOME_PROGRAM_TYPE',
 'INDIVIDUAL_DEVIATION_FLAG',
 'LAN_ID',
 'LEADID',
 'LOAN_PURPOSE',
 'LOAN_TYPE',
 'NO_FINANCE_DEVIATIONS',
 'PEP',
 'PRINCIPAL_OUTSTANDING',
 'PSL',
 'RISK_CATEGORIZATION',
 'ROI',
 'SELF_EMPLOYED_RISK',
 'TOTAL_TENOR',
 'TYPES_OF_REJECT']
loan_view=loan_view[lv_cols]
# cv_cols=find_columns(customer_view, dtcron)
cv_cols=['CUSTOMER_CITY_NAME',
 'CUSTOMER_COUNTRY_DESC',
 'CUSTOMER_INDUSTRY_DESC',
 'CUSTOMER_QUALIFICATION_DESCRIPTION',
 'CUSTOMER_RESIDENTIAL_STATUS',
 'CUSTOMER_SECTOR_DESC',
 'CUSTOMER_SUB_SECTOR_CODE',
 'CUSTOMER_SUB_SECTOR_DESC',
 'EMPLOYERNAME',
 'LAN_ID',
 'NUMBER_OF_DEPENDENTS',
 'OCCUPATION_CATEGORY',
 'QUALIFICATION',
 'SCORE',
 'SUB_CATEGORY']
cv_cols.append('CUSTOMER_CIF')
customer_view=customer_view[cv_cols]
# iv_cols=find_columns(insurance_view,dtcron)
iv_cols= ['CHANNEL_CODE', 'LAN_ID', 'NET_PREMIUM']
insurance_view=insurance_view[iv_cols]
insurance_view.set_index('LAN_ID', inplace=True)
insurance_view.reset_index(drop=False, inplace=True)
loan_view.set_index('LAN_ID', inplace=True)
loan_view.reset_index(drop=False, inplace=True)
base_view.set_index('LAN_ID', inplace=True)
base_view.reset_index(drop=False, inplace=True)
customer_view.set_index('LAN_ID', inplace=True)
customer_view.reset_index(drop=False, inplace=True)




insurance_view=remove_error_lans(insurance_view)
loan_view=remove_error_lans(loan_view)
base_view=remove_error_lans(base_view)
customer_view=remove_error_lans(customer_view)
base_view=base_view.merge(customer_view,on='LAN_ID',how='left')
base_view=base_view.merge(loan_view, left_on='LAN_ID',right_on='LAN_ID',how='left')
insurance_view=insurance_view.groupby('LAN_ID',as_index=False).sum('NET_PREMIUM')
base_view=base_view.merge(insurance_view,left_on='LAN_ID',right_on='LAN_ID',how='left')

# base_view.to_excel(r"C:\Users\VAIBHAV.SRIVASTAV01\Desktop\bv.xlsx")
