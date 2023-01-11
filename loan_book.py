# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 22:26:36 2022

@author: vaibhav.srivastav01
"""

import pandas as pd 
import numpy as np
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


ffm=pd.read_sql("select FINANCE_REFERENCE,FINANCE_BRANCH as BRANCH_CODE,finance_type,finance_start_date,maturity_date from prod_da_db.serve.f_finance_main where dh_record_active_flag='Y';",con_dm)
branch=pd.read_sql("select BRANCH_CODE,BRANCH_DESCRIPTION from prod_da_db.serve.rmt_branches",con_dm)
loan_type_name=pd.read_sql("select FINANCE_TYPE,FINANCE_TYPE_DESCRIPTION from prod_da_db.serve.rmt_finance_types where dh_record_active_flag='Y';",con_dm)
cif=pd.read_sql("select customer_cif,customer_shrtname from prod_da_db.serve.d_customer where dh_record_active_flag='Y'",con_dm)
xref_lan_to_cif=pd.read_sql("select customer_cif,lan_id as FINANCE_REFERENCE from prod_da_db.serve.x_ref_lan_to_cif where dh_record_active_flag='Y' and applicant_type='APPLICANT';",con_dm)
cif=cif.merge(xref_lan_to_cif, on='CUSTOMER_CIF',how='left')
cif=cif[['FINANCE_REFERENCE','CUSTOMER_CIF','CUSTOMER_SHRTNAME']]
ghf_pft=pd.read_sql("select finance_reference,tdschedule_pri_balance as Principal_Due,total_pri_balance as Principal_NotDue ,current_od_days as DPD_FOR_LAN from prod_da_db.serve.f_finance_pft_details where dh_record_active_flag='Y';",con_dm)
ghf_prov=pd.read_sql("select f1.finance_reference, f1.provision_rate as Provision_Percent,f1.provision_amount_calculate as Provision_Amount, f2.bucket_description as NPA_STAGE from prod_da_db.serve.f_finance_provisions f1 left join (prod_da_db.serve.npa_buckets) as f2 on  f1.npa_bucket_id = f2.bucket_id where f1.dh_record_active_flag='Y';",con_dm)
ghf_cust_view = pd.read_sql("SELECT NET_ANNUAL,CUSTOMER_CIF,CUSTOMER_CATEGORY_CODE,SCORE,PAN_NUMBER,SUB_CATEGORY,CUSTOMER_CITY_NAME,CUSTOMER_PROVINCE_NAME,CUSTOMER_SUB_SECTOR_DESC as INDUSTRY_CLASSIFICATION,customer_caste_desc as CASTE ,customer_religion_desc as Religion from prod_da_db.serve.customer_base WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm)


ffm=ffm.merge(branch,on='BRANCH_CODE',how='left')
ffm=ffm.merge(loan_type_name,on='FINANCE_TYPE',how='left')
ffm=ffm.merge(cif,on='FINANCE_REFERENCE',how='left')
ffm=ffm.merge(ghf_cust_view,on='CUSTOMER_CIF',how='left')
ffm=ffm.merge(ghf_pft,on='FINANCE_REFERENCE',how='left')
ffm=ffm.merge(ghf_prov,on='FINANCE_REFERENCE',how='left')



ffm_gfl=pd.read_sql("select FINANCE_REFERENCE,FINANCE_BRANCH as BRANCH_CODE,finance_type,finance_start_date,maturity_date from prod_gfl_da_db.serve.f_finance_main where dh_record_active_flag='Y';",con_dm_nbfc)
branch_gfl=pd.read_sql("select BRANCH_CODE,BRANCH_DESC as BRANCH_DESCRIPTION  from prod_gfl_da_db.serve.rmt_branches",con_dm_nbfc)
loan_type_name_gfl=pd.read_sql("select FINANCE_TYPE,FINANCE_TYPE_DESCRIPTION from prod_gfl_da_db.serve.rmt_finance_types where dh_record_active_flag='Y';",con_dm_nbfc)
cif_gfl=pd.read_sql("select customer_cif,customer_shrtname from prod_gfl_da_db.serve.d_customer where dh_record_active_flag='Y';",con_dm_nbfc)
xref_lan_to_cif_gfl=pd.read_sql("select customer_cif,lan_id as FINANCE_REFERENCE from prod_gfl_da_db.serve.x_ref_lan_to_cif where dh_record_active_flag='Y' and applicant_type='APPLICANT';",con_dm_nbfc)
cif_gfl=cif_gfl.merge(xref_lan_to_cif_gfl, on='CUSTOMER_CIF',how='left')
cif_gfl=cif_gfl[['FINANCE_REFERENCE','CUSTOMER_CIF','CUSTOMER_SHRTNAME']]
gfl_pft=pd.read_sql("select finance_reference,tdschedule_pri_balance as Principal_Due,total_pri_balance as Principal_NotDue,current_od_days as DPD_FOR_LAN  from prod_gfl_da_db.serve.f_finance_pft_details where dh_record_active_flag='Y';",con_dm_nbfc)
gfl_prov=pd.read_sql("select f1.finance_reference, f1.provision_rate as Provision_Percent,f1.provision_amount_calculate as Provision_Amount, f2.bucket_description as NPA_STAGE from prod_gfl_da_db.serve.f_finance_provisions f1 left join (prod_gfl_da_db.serve.npa_buckets) as f2 on  f1.npa_bucket_id = f2.bucket_id where f1.dh_record_active_flag='Y';",con_dm_nbfc)
gfl_cust_view = pd.read_sql("SELECT NET_ANNUAL,CUSTOMER_CIF,CUSTOMER_CATEGORY_CODE,SCORE,PAN_NUMBER,SUB_CATEGORY,CUSTOMER_CITY_NAME,CUSTOMER_PROVINCE_NAME,CUSTOMER_SUB_SECTOR_DESC as INDUSTRY_CLASSIFICATION,customer_caste_desc as CASTE ,customer_religion_desc as Religion from prod_gfl_da_db.serve.customer_base WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm_nbfc)



ffm_gfl=ffm_gfl.merge(branch_gfl,on='BRANCH_CODE',how='left')
ffm_gfl=ffm_gfl.merge(loan_type_name_gfl,on='FINANCE_TYPE',how='left')
ffm_gfl=ffm_gfl.merge(cif_gfl,on='FINANCE_REFERENCE',how='left')
ffm_gfl=ffm_gfl.merge(gfl_cust_view,on='CUSTOMER_CIF',how='left')
ffm_gfl=ffm_gfl.merge(gfl_pft,on='FINANCE_REFERENCE',how='left')
ffm_gfl=ffm_gfl.merge(gfl_prov,on='FINANCE_REFERENCE',how='left')
fin_fm=pd.concat([ffm,ffm_gfl])



ghf_loan_view=pd.read_sql("SELECT * from prod_da_db.serve.loan_view WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm)
ghf_base_view = pd.read_sql("SELECT * from prod_da_db.serve.BASE_VIEW WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm)
ghf_base_view['NBFC_FLAG'] = 'N'
ghf_disb=pd.read_sql("select * from  prod_da_db.serve.f_finance_disbursement_details WHERE DH_RECORD_ACTIVE_FLAG = 'Y';", con_dm)
ghf_collat=pd.read_sql("select collateral_reference, concat(property_address_1,property_address_2, property_address_3) as collateral_details, collateral_value, city as property_city, state as property_state from prod_da_db.serve.collateral_view;",con_dm)
ghf_lancollat=pd.read_sql("select lan_id,collateral_reference from prod_da_db.serve.x_ref_lan_to_collat where dh_record_active_flag='Y' ;",con_dm)
ghf_collat=ghf_lancollat.merge(ghf_collat,on='COLLATERAL_REFERENCE',how='left')

gfl_loan_view=pd.read_sql("SELECT * from prod_gfl_da_db.serve.loan_view WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm_nbfc)
gfl_base_view = pd.read_sql("SELECT * from prod_gfl_da_db.serve.BASE_VIEW WHERE DH_RECORD_ACTIVE_FLAG = 'Y'", con_dm_nbfc)
gfl_base_view['NBFC_FLAG'] = 'Y'
gfl_disb=pd.read_sql("select * from  prod_gfl_da_db.serve.f_finance_disbursement_details WHERE DH_RECORD_ACTIVE_FLAG = 'Y';", con_dm_nbfc)
gfl_collat=pd.read_sql("select collateral_reference, concat(property_address_1,property_address_2, property_address_3) as collateral_details, collateral_value, city as property_city, state as property_state from prod_gfl_da_db.serve.collateral_view;",con_dm_nbfc)
gfl_lancollat=pd.read_sql("select lan_id,collateral_reference from prod_gfl_da_db.serve.x_ref_lan_to_collat where dh_record_active_flag='Y';",con_dm_nbfc)
gfl_collat=gfl_lancollat.merge(gfl_collat,on='COLLATERAL_REFERENCE',how='left')





loan_view=pd.concat([ghf_loan_view,gfl_loan_view]) #REFERENCE
base_view=pd.concat([ghf_base_view,gfl_base_view]) # LAN_ID
base_view=base_view[['LAN_ID','BOOKING_DATE','BOOKING_AMOUNT','EOMLOGN','EOMSNCTN','NBFC_FLAG','NET_PREMIUM','STATUS']]
loan_view.rename(columns={'REFERENCE':'LAN_ID'},inplace=True)
loan_view=loan_view[['LAN_ID','LOAN_PURPOSE','ROI','PRINCIPAL_OUTSTANDING','GPL_FLAG','SUB_PRODUCT','FINAL_LTV','CR_EXPOSURE','SCHEME_MORATORIUM']]
base_view=base_view.merge(loan_view, on='LAN_ID',how='left')
disb=pd.concat([ghf_disb,gfl_disb])
disb=disb[['FINANCE_REFERENCE','DISBURSEMENT_DATE', 'DISBURSEMENT_AMOUNT','DISBURSEMENT_SEQUENCE']]
disb.rename(columns={'FINANCE_REFERENCE':'LAN_ID'},inplace=True)
disb_date=disb.copy()
disb_date=disb_date[disb_date['DISBURSEMENT_SEQUENCE']==1]
disb_date=disb_date[['LAN_ID','DISBURSEMENT_DATE']]
collat_view=pd.concat([ghf_collat,gfl_collat])
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


base_view['EOMLOGN'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)
base_view['EOMSNCTN'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)
base_view['BOOKING_DATE'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)
fin_fm=fin_fm.merge(base_view,left_on='FINANCE_REFERENCE',right_on='LAN_ID',how='left')
fin_fm=fin_fm[fin_fm['STATUS'].isin(['Booked','Cancelled'])]
fin_fm=fin_fm.merge(collat_view,left_on='FINANCE_REFERENCE',right_on='LAN_ID',how='left')
disb=disb.groupby('LAN_ID').sum('DISBURSEMENT_AMOUNT')
disb=disb.merge(disb_date,on='LAN_ID',how='left')
fin_fm=fin_fm.merge(disb,left_on='FINANCE_REFERENCE',right_on='LAN_ID',how='left')
fin=fin_fm[['FINANCE_REFERENCE','BRANCH_CODE', 'NET_ANNUAL','FINANCE_TYPE','FINANCE_START_DATE', 'MATURITY_DATE', 'BRANCH_DESCRIPTION','CUSTOMER_CITY_NAME','CUSTOMER_PROVINCE_NAME', 'FINANCE_TYPE_DESCRIPTION', 'CUSTOMER_CIF', 'SCHEME_MORATORIUM','CUSTOMER_SHRTNAME','EOMSNCTN','BOOKING_DATE','CUSTOMER_CATEGORY_CODE','BOOKING_AMOUNT','DISBURSEMENT_AMOUNT','DISBURSEMENT_DATE','ROI','FINAL_LTV','CR_EXPOSURE','LOAN_PURPOSE','SUB_CATEGORY','SCORE','NBFC_FLAG','PAN_NUMBER','CASTE' ,'RELIGION', 'PRINCIPAL_DUE', 'PRINCIPAL_NOTDUE','PROVISION_PERCENT','PROVISION_AMOUNT','DPD_FOR_LAN','NPA_STAGE','COLLATERAL_DETAILS', 'COLLATERAL_VALUE', 'PROPERTY_CITY', 'PROPERTY_STATE','INDUSTRY_CLASSIFICATION','NET_PREMIUM']]
fin['Monthly Return Classification'] = np.where(fin['LOAN_PURPOSE'] == 'Home Loan Resale ', 'Home Loan Resale',   #when... then
                  np.where(fin['FINANCE_TYPE'].isin(['LP','NP','LT','FL','FT']), 'For mortgage/property/home equity loans',  #when... then
                    'Housing loans to individuals for construstion/ purchase of new units'))  
fin['Approval Date'] = fin['BOOKING_DATE']
fin['Undisbursed Amount'] = fin['BOOKING_AMOUNT'] - fin['DISBURSEMENT_AMOUNT']/100

state=pd.read_excel(r"C:\Users\VAIBHAV.SRIVASTAV01\Documents\country_vs_province (1) (1).xlsx")
state=state[['cpprovince', 'cpprovincename']]
city=pd.read_excel(r"C:\Users\VAIBHAV.SRIVASTAV01\Documents\province_vs_city (1).xlsx")
city=city[['pccity','pccityname']]

fin=fin.merge(city,left_on='PROPERTY_CITY',right_on='pccity',how='left')
fin=fin.merge(state,left_on='PROPERTY_STATE',right_on='cpprovince',how='left')
lb=pd.read_excel(r"C:\Users\VAIBHAV.SRIVASTAV01\Desktop\Consolidate_Loan Book_Nov'22.xlsx",sheet_name='Sheet1')
lb=lb[lb['Agreement Number']==' ']

lb['Agreement Number']=fin['FINANCE_REFERENCE']
lb[ 'Branch ID']=fin['BRANCH_CODE']
lb[ 'Branch Name']=fin['BRANCH_DESCRIPTION']
lb[ 'Loan Type']=fin['FINANCE_TYPE']
lb[  'Loan Type Name']=fin['FINANCE_TYPE_DESCRIPTION']
lb[ 'Customer CIF']=fin[ 'CUSTOMER_CIF']
lb[ 'Customer Name']=fin['CUSTOMER_SHRTNAME']
lb[ 'PAN']=fin['PAN_NUMBER']
lb['Sch.Caste / Sch.Tribe']=fin['CASTE']
lb['Religion']=fin['RELIGION']
lb['Loan Start Date']=fin['FINANCE_START_DATE']
lb[ 'Maturity Date']=fin['MATURITY_DATE']
lb[ 'NPA Stage']=fin['NPA_STAGE']
# lb[ 'Tenure']=fin[]
# lb['Balance Tenure']=fin[], ,,
lb[ 'Principal Due']=fin['PRINCIPAL_DUE']/100
lb[ 'Principal Not Due']=fin['PRINCIPAL_NOTDUE']/100
lb[ 'Provision %']=fin['PROVISION_PERCENT']*100
lb[ 'Provision Amount']=fin['PROVISION_AMOUNT']/100
lb[ 'State of Customer']=fin['CUSTOMER_PROVINCE_NAME']
lb[ 'City of Customer']=fin['CUSTOMER_CITY_NAME']
lb[ 'City of Property']=fin['pccityname']
lb[ 'State of Property']=fin[ 'cpprovincename']
lb['Monthly Return Classification']=fin['Monthly Return Classification']
lb[ 'Insurance in Gross Disbursal']=fin['NET_PREMIUM']
lb['Sanction Date']=fin['EOMSNCTN']
lb['Approval Date']=fin['Approval Date']
lb['Undisbursed Amount'] =fin['Undisbursed Amount'] 
lb['Cust Type']=fin['CUSTOMER_CATEGORY_CODE']
lb['Sanctioned Amount']=fin['BOOKING_AMOUNT']
lb['ROI']=fin['ROI']
lb['LTV Ratio (at the time of sanction)']=fin['FINAL_LTV']
lb['CRE cases']=fin['CR_EXPOSURE']
lb['Loan Purpose']=fin['LOAN_PURPOSE']
lb['Employment_Type']=fin['SUB_CATEGORY']
for i in range(len(fin)):
    if fin.loc[i,'SCORE']!=None:
        fin.loc[i,'SCORE']=fin.loc[i,'SCORE'][2:]
fin['SCORE']=np.where(fin['SCORE']=='0-1','000-1',fin['SCORE'])
lb['CIBIL Score']=fin['SCORE']
lb['Morat']=np.where(fin['SCHEME_MORATORIUM']==1,'YES','NO')
lb['NBFC_FLAG']=fin['NBFC_FLAG']
lb['DPD for LAN']=fin['DPD_FOR_LAN']
lb['Collateral details']=fin['COLLATERAL_DETAILS']
lb['Property Value']=fin['COLLATERAL_VALUE']
lb['Industry classification']=fin['INDUSTRY_CLASSIFICATION']
lb['DISB_AMOUNT']=fin['DISBURSEMENT_AMOUNT']/100
lb['First Disbursal Date']=fin['DISBURSEMENT_DATE']
lb['Intrest Accrued & Not Due']=0
lb['Intrest Accrued & Due']=0

lb['Total Principal Outstanding']=lb['Principal Due']+lb['Principal Not Due']

lb['Total Interest Receivable']=lb['Intrest Accrued & Due']+lb['Intrest Accrued & Not Due']

lb['Total Outstanding Amount']=lb['Total Principal Outstanding']+lb['Total Interest Receivable']
lb['Memo Charges (Yet to be Collected)']=0
lb['Total Receivable']=lb['Total Outstanding Amount']+lb['Memo Charges (Yet to be Collected)']
lb['Housing/Non-housing']=np.where(lb['Loan Type'].isin(['HL','HT']),'Housing','Non-housing')
lb['PAN'].replace(to_replace=[np.nan,'',' '],value='00000000',inplace=True)
for i in range(len(lb['PAN'])):
    lb.loc[i,'4th']=lb.loc[i,'PAN'][3]
lb['Individual or Not']=np.where(lb['4th']=="P",'Individual',np.where(lb['4th']=="C",'Company',np.where(lb['4th']=="H" ,'Hindu Undivided Family (HUF)',np.where(lb['4th']=="A",'Association of Persons (AOP)',np.where(lb['4th']=="B" ,'Body of Individuals (BOI)',np.where(lb['4th']=="G" ,'Government Agency',np.where(lb['4th']=="J" ,'Artificial Juridical Person',np.where(lb['4th']=="L" ,'Local Authority',np.where(lb['4th']=="F" ,'Firm/ Limited Liability Partnership',np.where(lb['4th']=="T" ,'Trust',0))))))))))
# condition=[((lb['NPA Stage']=='Standard') & (lb['CRE cases']=='NON CRE') & (lb['Housing/Non-housing']=='Housing') & (lb['Individual or Not']=='Individual')),((lb['NPA Stage']=='Standard') & (lb['CRE cases']=='CRE-RH') & (lb['Housing/Non-housing']=='Housing') & (lb['Individual or Not']=='Individual')),((lb['NPA Stage']=='Standard') & (lb['CRE cases'].isin['CRE-H','CRE']) & (lb['Housing/Non-housing']=='Housing') & (lb['Individual or Not']=='Individual')),((lb['NPA Stage']=='Standard') & (lb['CRE cases']=='NON CRE') & (lb['Housing/Non-housing']=='Housing') & (lb['Individual or Not']=='Non-individual')),((lb['NPA Stage']=='Standard') & (lb['CRE cases'].isin['CRE-H','CRE']) & (lb['Housing/Non-housing']=='Housing') & (lb['Individual or Not']=='Non-individual')),((lb['NPA Stage']=='Standard') & (lb['CRE cases']=='NON CRE') & (lb['Housing/Non-housing']=='Non-housing') & (lb['Individual or Not']=='Individual')),((lb['NPA Stage']=='Standard') & (lb['CRE cases'].isin['CRE','CRE-H']) & (lb['Housing/Non-housing']=='Non-housing') & (lb['Individual or Not']=='Individual')),((lb['NPA Stage']=='Standard') & (lb['CRE cases'].isin['CRE-H','CRE']) & (lb['Housing/Non-housing']=='Non-housing') & (lb['Individual or Not']=='Non-individual')),((lb['NPA Stage']=='Standard') & (lb['CRE cases']=='NON CRE') & (lb['Housing/Non-housing']=='Non-housing') & (lb['Individual or Not']=='Non-individual')),(lb['NPA Stage']=='Sub Standard')]
# value=[0.0025,0.0075,0.01,0.004,0.01,0.004,0.01,0.01,0.004,0.15]
# lb['Provision %']=np.select(condition,value)
# lb['Provision Amount']=lb['Provision %']*lb['Total Receivable']

lb['Additional Provision']=np.where(lb['Provision %']<0.004,(0.004-lb['Provision %'])*lb['Total Receivable']/100,0)

lb['Total Provision']=lb['Additional Provision']+lb['Provision Amount']
lb['Secured / Unsecured']=np.where(lb['Loan Type'].isin(['FL','FT']),'Unsecured','Secured')
lb["Disbursed LTD"]=lb['DISB_AMOUNT']##"Disbursed as on August'22"

lb['Insurance portion %']=lb['Insurance in Gross Disbursal']/lb["Disbursed LTD"]

lb['DPD for CIF']=0
lb['IRR']=0
lb['Insurance Portion on POS']=np.where(lb['Insurance in Gross Disbursal']>0,lb['Total Principal Outstanding']*lb['Insurance portion %'],0)
lb['Insurance Portion on Accrued Interest']=np.where(lb['Insurance in Gross Disbursal']>0,lb['Total Interest Receivable']*lb['Insurance portion %'],0)
lb['Total Insurance']=lb['Insurance Portion on POS']+lb['Insurance Portion on Accrued Interest']
lb['Processing Fee']=np.where(lb['Loan Type'].isin(['HL','HT']),0,0)
lb['PF portion %']=lb['Processing Fee']/lb["Disbursed LTD"]


lb['PF Portion on POS']=np.where(lb['Processing Fee']>0,lb['Total Principal Outstanding']*lb['PF portion %'],0)


lb['PF Portion on Accrued Interest']=np.where(lb['Processing Fee']>0,lb['Total Interest Receivable']*lb['PF portion %'],0)
lb['Total PF']=lb['PF Portion on Accrued Interest']+lb['PF Portion on POS']
lb['Total outstanding (less Insurance)']=lb['Total Outstanding Amount']-lb['Total Insurance']
lb['DPD Stage']=np.where(lb['DPD for CIF']>90, 'Stage 3',np.where(lb['DPD for CIF']>30 ,'Stage 2','Stage 1'))

lb['LTV Ratio total outstanding (excluding insurance) to property value']=lb['Total outstanding (less Insurance)']/lb['Property Value']
lb['Total outstanding for LTV']=np.where(lb['Total Outstanding Amount']<=3000000.0,"Loans upto 30 lacs",np.where((lb['Total Outstanding Amount']>3000000.0) & lb['Total Outstanding Amount']<=7500000.0 ,"Loans upto 75 lacs",np.where(lb['Total Outstanding Amount']>7500000.0,"Loans above 75 lacs",0)))
# condition=[lb['NPA Stage']=='Standard'&]
# value=[]
lb['Risk Weight']=0
lb['RWA']=0
lb['EOM']=lb['Approval Date']
lb['Income of Borrower']=fin['NET_ANNUAL']

###########


import datetime as datetime
###########
lb['Pool Buyout (yes/no)']='No'
lb['Sanction Date']=pd.to_datetime(lb['Sanction Date']).dt.date
year=datetime.datetime.now().year
lb['Sanction Date'].replace(to_replace=[np.nan,'',' '],value=date.date(1900,1,1),inplace=True)
lb['SANC_YEAR']=pd.to_datetime(lb['Sanction Date']).dt.year
lb['SANC_YEAR']=lb['SANC_YEAR'].astype(int)
lb['SANC_MONTH']=pd.to_datetime(lb['Sanction Date']).dt.month
lb['SANC_MONTH']=lb['SANC_MONTH'].astype(int)
lb['SANC_YEAR_MONTH']=lb['SANC_YEAR'].astype(str)+'-'+ lb['SANC_MONTH'].astype(str)
todays_date=datetime.datetime.now()
MTD=[(str(todays_date.year)+"-"+str(todays_date.month))]    
MTD_1=[(str(todays_date.year)+"-"+str(todays_date.month-1))]   
lb['Sanction Amount as on Previous Year']=np.where(lb['Sanction Date']<=datetime.date(year,3,31),lb['Sanctioned Amount'],0)
# # lb['Undisbursed Amount as on 31st March 2022']
# lb["Disbursed as on Mar'22"]=lb['Sanction Amount_MIS as on 31st March, 2022']-lb['Undisbursed Amount as on 31st March 2022']
# # lb["Insurance Portion as on March'22"]???
lb["Sanction Amount YTD"]=np.where(lb['Sanction Date']>datetime.date(year,3,31),lb['Sanctioned Amount'],0)


# lb["Disbursed YTD_August'22"]=np.where(lb['Pool Buyout (yes/no)']=='No',(lb["Disbursed as on August'22"]-lb["Disbursed as on Mar'22"]),0)
# lb["Insurance Portion YTD August'22"]=np.where(lb['Pool Buyout (yes/no)']=='No',(lb['Insurance in Gross Disbursal']-lb["Insurance Portion as on March'22"]),0)
lb['Sanction Amount as on Previous month']=np.where(lb['SANC_YEAR_MONTH']==MTD_1[0],lb['Sanctioned Amount'],0)


# # lb['Undisbursed Amount as on Previous month']=???
# lb['Disbursed as on Previous month']=lb['Sanction Amount_MIS as on Previous month']-lb['Undisbursed Amount as on Previous month']
# #lb['Insurance Portion as on Previous month']=??
lb["Sanction Amount MTD"]=np.where(lb['SANC_YEAR_MONTH']==MTD[0],lb['Sanctioned Amount'],0)



# lb["Disbursed FTM_August'22"]=np.where(lb['Pool Buyout (yes/no)']=='No',lb["Disbursed YTD_August'22"]-lb['Disbursed as on Previous month'],0)
# lb["Insurance Portion FTM_August'22"]=np.where(lb['Pool Buyout (yes/no)']=='No',lb['Insurance in Gross Disbursal']-lb['Insurance Portion as on Previous month'],0)
#lb['Current']=???
# lb['Non-current']=lb['Total Principal Outstanding']-lb['Current']
# lb['Insurance Portion (Current)']=np.where(lb['Total Principal Outstanding']>0,(lb['Insurance Portion on POS']*lb['Current'])/(lb['Current']+lb['Non-current']),0)
# lb['Insurance Portion (Non-Current)']=np.where(lb['Total Principal Outstanding']>0,(lb['Insurance Portion on POS']*lb['Non-current'])/(lb['Current']+lb['Non-current']),0)
# lb['ECL (Current)']=np.where(lb['Current']>0,(lb['Total ECL']*lb['Current'])/(lb['Current']+lb['Non-current']),0)
# lb['ECL (Non-Current)']=lb['Total ECL']-lb['ECL (Current)']
# lb['ECL on Insurance portion (Current)']=np.where(lb['Total Principal Outstanding']>0,lb['Total ECL']*lb['Insurance Portion (Current)']/lb['Total Principal Outstanding'],0)
# lb['ECL on Insurance portion (Non-Current)']=np.where(lb['Total Principal Outstanding']>0,lb['Total ECL']*lb['Insurance Portion (Non-Current)']/lb['Total Principal Outstanding'],0)
# # lb['UW cost']=??
# # lb['Employee incentive']=??
# # lb['CP commission']=??
# # lb['Processing Fee income']=??
# lb['Total Loan acquistion cost']=lb['UW cost']+lb['Employee incentive']+lb['CP commission']+lb['Processing Fee income']
# #lb['UW cost (Current)']=??
# lb['UW cost (Non-current)']=lb['UW cost']-lb['UW cost (Current)']
# # lb['Employee incentive (Current)']=??
# lb['Employee incentive (Non-current)']=lb['Employee incentive']-lb['Employee incentive (Current)']
# # lb['CP commission (Current)']=??
# lb['CP commission (Non-current)']=lb['CP commission']-lb['CP commission (Current)']
# # lb['Processing Fee income (Current)']=??
# lb['Processing Fee income (Non-current)']=lb['Processing Fee income']-lb['Processing Fee income (Current)']
# lb['Total Housing loans']=np.where(lb['Loan Type'].isin(['HL','HT']),(lb['Total Outstanding Amount']-lb['Total Insurance']+lb['Total Loan acquistion cost']),0)
# lb['Total Non-Housing loans']=(lb['Total Outstanding Amount']-lb['Total Housing loans']+lb['Total Loan acquistion cost'])
# lb['Total loans']=lb['Total Non-Housing loans']+lb['Total Housing loans']
lb['Sanctioned Amount'].replace(to_replace=[np.nan,'',' '],value=0,inplace=True)
lb['Sanctioned Amount']=lb['Sanctioned Amount'].astype(int)
condition=[(lb['Sanctioned Amount']<200001),(lb['Sanctioned Amount']>200000)&( lb['Sanctioned Amount']<500001),(lb['Sanctioned Amount']>500000) & (lb['Sanctioned Amount']<1000001),(lb['Sanctioned Amount']>1000000) & (lb['Sanctioned Amount']<2500001),(lb['Sanctioned Amount']>2500000) & (lb['Sanctioned Amount']<5000001),lb['Sanctioned Amount']>5000000]
value=["Upto 2 lakhs","More than 2 lakhs to upto 5 lakhs","More than 5 lakhs upto 10 lakhs","More than 10 lakhs upto 25 lakhs","More than 25 lakhs upto 50 lakhs","More than 50 lakhs"]
lb['Sanction Slabwise']=np.select(condition,value)
condition=[lb['ROI']<=5,(lb['ROI']>5) & (lb['ROI']<=10),(lb['ROI']>10) & (lb['ROI']<=15),(lb['ROI']>15) & (lb['ROI']<=20),(lb['ROI']>20) & (lb['ROI']<=30),lb['ROI']>30]
value=["Upto 5%","Above 5% to 10%","Above 10% to 15%","Above 15% to 20%","Above 20% to 30%","Above 30%"]
lb['Interest range']=np.select(condition,value)

lb['WROI']=lb['ROI']*lb['Total Principal Outstanding']/100
condition=[(lb['Income of Borrower']<300001),(lb['Income of Borrower']>300000 )& (lb['Income of Borrower']<600001),(lb['Income of Borrower']>600000) & (lb['Income of Borrower']<1800001),(lb['Income of Borrower']>1800000)]
value=["Upto 3 lakhs","More than 3 lakhs to upto 6 lakhs","More than 6 lakhs upto 18 lakhs","More than 18 lakhs"]
lb['Income slab wise']=np.select(condition,value)
lb['Exposure']=np.where(lb['Total Outstanding Amount']<=1500000.0,"Upto 15 Lakhs","More than 15 lakhs")
lb.replace(to_replace=[np.nan,'',' '],value=' ',inplace=True)
lb.to_excel(r"C:\Users\VAIBHAV.SRIVASTAV01\Desktop\Loan Book_test_1.xlsx")
#lb['Revised Maturity Date']=??
# lb['Check']
# lb['Top 20']
# lb["Principal Out As on July'22"]
# lb["Capitalisation August'22"]
# lb['Repayment ']
# lb["Principal Out as on July'22"]
# lb["Repayment in Aug'22"]
# lb['Final Maturity Date']
# lb=lb[['Agreement Number', 'Branch ID', 'Branch Name', 'Channel Code',
#        'Channel Description', 'Loan Type', 'Loan Type Name', 'Customer CIF',
#        'Customer Name', 'PAN', 'Religion', 'CASTE','Sanction Date','Approval Date',
#        'Loan Start Date', 'Maturity Date','NPA Stage','Cust Type' ,'Tenure','Balance Tenure' ,'Interest Type',
#        'ROI','Sanctioned Amount','Income of Borrower', 'Principal Due', 'Principal Not Due',
#        'Total Principal Outstanding', 'Intrest Accrued & Due',
#        'Intrest Accrued & Not Due', 'Total Interest Receivable',
#        'Total Outstanding Amount', 'Memo Charges (Yet to be Collected)',
#        'Total Receivable', 'Provision %', 'Provision Amount',
#        'Additional Provision', 'Total Provision', 'Undisbursed Amount',
#        'Secured / Unsecured', 'NPA Date', 'Intrest Accrued & Due Reversal',
#        'Intrest Accrued & Not Due Reversal', 'Pool Buyout (yes/no)',
#        'Collateral details', 'Property Size', 'State of Property','State of Customer',
#        'City of Property','City of Customer',  'Metro/ Urban/Rural', 'Property Value','LTV Ratio (at the time of sanction)',
#        'Monthly Return Classification',
#        'CRE cases','NBFC_FLAG','Insurance in Gross Disbursal', 'Insurance portion %','Loan Purpose',
#        'Collateral Type', 'Applicant Type', 'Co-Applicant Name',
#        'Co-applicant Type', 'Restructuring Date', 'Fraud Cases', 'DPD for LAN',
#        'DPD for CIF', 'Sell Down', 'Technical W/off', 'Processing Fee', 'IRR',
#        'Priority Sector Lending','Employment_Type', 'CIBIL Score','Morat (yes/no)',
#        'MSME (Micro & Small, Medium,Large)','Industry classification']]
