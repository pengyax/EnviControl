import pandas as pd
import glob
import traceback
import numpy as np

def combine_purchasing(path,year):
    try:
        finalexcelsheet = pd.DataFrame()
        filenames = glob.glob(path + "\*.xlsx")
        for file in filenames:
            df = pd.read_excel(file,skiprows=7,sheet_name=0)
            df.columns = df.columns.str.replace('\n','',regex=True)
            df = df.iloc[:,:5]
            df.columns = ['Div','Vendor Code','Vendor Name','Purchasing','YTD Purchasing']
            df['Month'] = file[-7:-5]
            df['Year'] = year
            finalexcelsheet = pd.concat([finalexcelsheet,df],ignore_index=True)
            finalexcelsheet.Month = finalexcelsheet.Month.map(int)
        return finalexcelsheet
    except Exception as e:
            traceback.print_exc()


def complaints_envic(path,vendor_list,year_list = [2022,2023],div_list = [21,22,51]):
    try:
        df_complaints_detail = pd.read_excel(path,sheet_name=1)
        df_complaints_detail['Year'] = df_complaints_detail['Notification Created Date'].dt.year
        df_complaints_detail['Quarter'] = df_complaints_detail['Notification Created Date'].dt.quarter
        df_complaints_query = (df_complaints_detail['Vendor Name'].isin(vendor_list)) & (df_complaints_detail['Year'].isin(year_list)) & (df_complaints_detail['If Manufacturing Complaint'] == 'Y') \
    & (df_complaints_detail['Division'].isin(div_list)) & (df_complaints_detail['Short Text For Cause Code'] == 'MFG: Environmental Controls')
        df_complaints_detail = df_complaints_detail.loc[df_complaints_query]
        df_complaints_envic_gy = df_complaints_detail.groupby(['Vendor Name','Year','Month','Division'],as_index=False).size()
        df_complaints_envic_gy.rename(columns = {'size':'Complaints #: Foreign Particulates'},inplace=True)
        return df_complaints_envic_gy
        
    except Exception as e:
            traceback.print_exc()
            
            
def inspection_envic(path,vendor_list,year_list = [2022,2023]):
    try:
        df_inspection_detail = pd.read_excel(path,sheet_name=0)
        df_inspection_detail_query = (df_inspection_detail['Vendor'].isin(vendor_list)) & (df_inspection_detail['Year'].isin(year_list)) \
    & (df_inspection_detail.Results.isin(['A','R'])) & (df_inspection_detail['Division'].isin(['21','22','51']))
        df_inspection_detail.rename(columns={'Vendor':'Vendor Name','Mon':'Month'},inplace=True)
        
        df_inspection_detail['Rework #: Foreign Particulate'] = 0
        df_inspection_detail.loc[(df_inspection_detail['Reject Code'] == 'Foreign Particulate') & df_inspection_detail['Rework'] == 1,'Rework #: Foreign Particulate'] = 1
        df_inspection_detail_envic = df_inspection_detail.copy().loc[df_inspection_detail_query]
        df_inspection_detail_envic['Division'] = df_inspection_detail_envic['Division'].map(int)
        df_inspection_detail_envic_gy = df_inspection_detail_envic.groupby(['Vendor Name','Division','Year','Month'],as_index=False).size()
        df_inspection_detail_envic_gy.rename(columns = {'size':'Inspection'},inplace=True)
        
        df_inspection_detail_ForeignParticulate = df_inspection_detail_envic.groupby(['Vendor Name','Division','Year','Month'],as_index=False)['Rework #: Foreign Particulate'].sum()
        
        return df_inspection_detail_envic_gy,df_inspection_detail_ForeignParticulate
    except Exception as e:
        traceback.print_exc()

if __name__ == "__main__":
    
    vendor_mapping = pd.read_excel(r'C:\Medline\CPM\data\vendor_mapping\Vendor _mapping 2022_v12.xlsx')
    vendor_mapping_dict = dict(zip(vendor_mapping['Vendor Number'],vendor_mapping['Cleaned Vendor Name']))
    df_vendor_list = pd.read_excel('../EnviControlData/datasource/vendor_list.xlsx',sheet_name=0)
    vendor_list = df_vendor_list['vendor_name'].to_list()
    purchasing_path = r'C:\Medline\CPM\finance\US\2023'
    df_purchasing_2023_detail =  combine_purchasing(purchasing_path,2023)
    df_purchasing_2023_detail['Vendor Name'] = df_purchasing_2023_detail['Vendor Code'].apply(lambda x : vendor_mapping_dict.get(x,np.nan))
    df_purchasing_2023 = df_purchasing_2023_detail.groupby(['Vendor Name','Year','Month'],as_index=False)[['Purchasing','YTD Purchasing']].sum()
    df_purchasing_2023 = df_purchasing_2023.iloc[:,:4]
    
    
    df_purchasing_2022 = pd.read_excel('../EnviControlData/datasource/2022 Purchasing.xlsx',sheet_name=0)
    df_purchasing = pd.concat([df_purchasing_2023,df_purchasing_2022],ignore_index=True)
    df_purchasing = df_purchasing.loc[(df_purchasing['Vendor Name'].isin(vendor_list)) &(df_purchasing['Purchasing'] != 0)]
    print(len(df_purchasing))
    
    complaints_path = r'../EnviControlData/datasource/US Complaint Data 201901-202302.xlsx'
    df_complaints_foreignParticulate = complaints_envic(complaints_path,vendor_list)
    print(df_complaints_foreignParticulate)
    
    insepction_path = r'../EnviControlData/datasource/InspectionDatabase_duplicateCombine_2022&2023.xlsx'
    df_inspection,df_inspection_foreignParticulate = inspection_envic(insepction_path,vendor_list)
    print(len(df_inspection))
    print(len(df_inspection_foreignParticulate))
    
    with pd.ExcelWriter('EnviControl_02.xlsx') as writer:
        df_vendor_list.to_excel(writer,index=False,sheet_name='vendor_list')
        df_complaints_foreignParticulate.to_excel(writer,index=False,sheet_name='env_complaints')
        df_purchasing.to_excel(writer,index=False,sheet_name='purchasing')
        df_inspection.to_excel(writer,index=False,sheet_name='inspection')
        df_inspection_foreignParticulate.to_excel(writer,index=False,sheet_name='foreignParticulate')
        
    Year = [2022,2023]
    Month = [1,2,3,4,5,6,7,8,9,10,11,12]
    Division = [21,22,51]
    index = pd.MultiIndex.from_product([vendor_list, Year, Month,Division], names=['Vendor Name', 'Year', 'Month','Division'])
    df_base = pd.DataFrame(index=index)
    df_base = df_base.reset_index()
    
    df = df_base.merge(df_inspection_foreignParticulate,how='left',on=['Vendor Name', 'Year', 'Month','Division']).merge(df_inspection,how='left',on=['Vendor Name', 'Year', 'Month','Division'])\
    .merge(df_complaints_foreignParticulate,how='left',on=['Vendor Name', 'Year', 'Month','Division']).merge(df_purchasing,how='left',on=['Vendor Name', 'Year', 'Month']).fillna(0)
    df = df.loc[~((df['Year'] == 2023) & (df['Month'] >=3))]
    query_zero = df['Rework #: Foreign Particulate'] + df['Inspection'] + df['Complaints #: Foreign Particulates'] + df['Purchasing'] == 0
    df.drop(df[query_zero].index,inplace=True)
    df.drop(columns='Purchasing',axis=1,inplace=True)
    
    with pd.ExcelWriter('EnviControl_gy_02.xlsx') as writer:
        df_vendor_list.to_excel(writer,index=False,sheet_name='vendor_list')
        df.to_excel(writer,index=False,sheet_name='sheet_all')
        df_purchasing.to_excel(writer,index=False,sheet_name='purchasing')