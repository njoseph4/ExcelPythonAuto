#!/usr/bin/env python
# coding: utf-8

# In[1]:


# In[2]:
import os
exec(open("Generic.py").read()) ## this will bring all libraries and credentials in place 


import time
import sys 
import pandas as pd

import glob
import os

import numpy as np

from datetime import timedelta
import datetime 


import os
from sqlalchemy import create_engine


if datetime.datetime.today().weekday() in [3]: ## which is a thursday 

    




    if os.system('start "excel" "Management_Upper_US.xls"')==0:
        os.system('start "excel" "Management_Upper_US.xls"')
    print("Already open so keep moving")
    template='Management_Upper_US.xls'

    time.sleep(60) 







        
    import xlwings as xw
    wbtest=xw.Book(template)
    wb1=wbtest.sheets['Sheet1']
    df_port=wb1.range('A1:E5000').options(pd.DataFrame).value
    df_port.reset_index(inplace=True)
    df_port=df_port[df_port['Ticker'].notnull()]

    df_port.to_sql('Management_Ticker_Wkl_Hist',schema ='ams',con=engine,chunksize=100,method='multi',index=False,if_exists='append')


    wbtest.close()

import pyautogui

pyautogui.FAILSAFE = False
import webbrowser


import glob
print("job started")
print(datetime.datetime.now())


files = glob.glob(r'\Desktop\Importgeniusdata\*')  ##careful here   ## clear all json files 
for f in files:
    print(f)
#     os.remove(f)

#     print("removed all files from root")




url='https://app.importgenius.com/'
webbrowser.open_new_tab(url)
print("opened url")
time.sleep(10) # time to run your search 
n=int(input("# of Pages"))

for i in range(n):


    time.sleep(10)
    

    
    time.sleep(2)
   
    
    key = pyautogui.locateOnScreen('next.PNG', grayscale=True, confidence=.9)
    pyautogui.moveTo(key.left+12, key.top+23)
    # 
    pyautogui.hotkey('ctrl', 's')
    pyautogui.typewrite(str(i) + '.html')
    pyautogui.press('enter')
    time.sleep(5)
    
    pyautogui.click()
    time.sleep(5)

    
    
    ### Anaomalous Updates 
    
       import os
    import time 
    ##parameters 
    import sys 

    region='US'

    import json

    import pandas as pd

    import glob
    import os

    import numpy as np

    import xlwings as xw

    import datetime
    import time
    import re 
    import os

    from datetime import timedelta
    import logging
    import pathlib
    import time


    exec(open(r"C:\Users\NaveenJoseph\Hunter Capital Limited Partnership\Firm - Documents\QUANTITATIVE\ScheduledJobs\cred.py").read())


    import os 
    from pandas import json_normalize

    template='HistoricBeatReport_Comparables.xlsx' 

    from sqlalchemy import create_engine
    from pyodbc import ProgrammingError
    import numpy as np

    from numpy.lib.function_base import append 
    if os.system('start "excel" "HistoricBeatReport_Comparables.xlsx"')==0:
        os.system('start "excel" "HistoricBeatReport_Comparables.xlsx"')
    print("Already open so keep moving")
    time.sleep(3)

    sector_list=['Consumer Discretionary','Materials','Information Technology','Utilities', 'Communication Services' ,'Industrials','Consumer Staples','Health Care']




#     # In[45]:

#     ##Now see what did they send upto now

#     conn = get_connection()
#     raw = conn.raw_connection()
#     cursor = raw.cursor()
#     cursor.execute('''


#     select distinct *
#     from ams."quant_revisions_vw"
#     where cast("load_date" as date) > current_date - 90




        
#     ''')

#     alldf=[]
#     for row in cursor:
#         df_h=pd.DataFrame(row)
#         alldf.append(df_h.T)

#     df_loaded = pd.concat(alldf)
#     df_loaded.columns=['sector','Tickers','Rev_revision','EBITDA_Rev','EPS_Rev','Score','Load_Date']

#     already_send_tickers = df_loaded['Tickers'].unique().tolist()








    #template='Template10SurpriseVal.xlsx' 

    def fetchdatafromtemplate(year):

        wbtest=xw.Book(template)
        wb1=wbtest.sheets['EBITDA_Spread']
        df=wb1.range('A4:S5000').options(pd.DataFrame).value
        df.columns=['ticker','sector',1,2,3,4,5,6,7,8,9,10,11,12,'Latest_Revenue']
        
        df['forecastyear']=df[year]
        
        return df 


# In[8]:


def getvaluesfromtemplate(target_year):

    wbtest=xw.Book(template)
    wb1=wbtest.sheets['EBITDA_Spread']
    wb1.range('C1').value=target_year
    time.sleep(30)
    
    df=wb1.range('A3:P5000').options(pd.DataFrame).value
    df.reset_index(inplace=True)
    df.columns=['ticker','sector',1,2,3,4,5,6,7,8,9,10,11,12,13,'Latest_Revenue']
    df=df[df['ticker'].notnull()]
    df=df.melt(id_vars=["ticker","Latest_Revenue","sector"],var_name="Trailing_Months",value_name="Estimate_MoM_Var")
    df['Forecast_Year']=target_year
    return df 



# In[7]:


allcalyr=[]
for i in range(2010,2030,1):
    if i < int(datetime.datetime.today().strftime('%Y'))+2:
        allcalyr.append(i)
allcalyr


# In[9]:


alldf=[]
for cal_yr in allcalyr:
    print(cal_yr)
    time.sleep(10)
    df_samp=getvaluesfromtemplate(cal_yr)
    alldf.append(df_samp)
    
final_df=pd.concat(alldf)
final_df


# In[12]:


final_df.to_csv("BackupMoMEstimate.csv")


# In[19]:


df_q_comp= pd.read_excel(r"Quick_Comps.xls",sheet_name='Sheet1')
df_q_comp


# In[31]:


fyear= int(datetime.datetime.today().strftime('%Y'))+1
curr_month=int(datetime.datetime.today().strftime('%m'))+1
fyear,curr_month


# In[32]:


final_df[final_df['ticker']=='NasdaqGS:AAPL']


# In[84]:


final_df_cal= final_df[['Forecast_Year','Trailing_Months']].drop_duplicates().sort_values(by=['Forecast_Year','Trailing_Months']).reset_index(drop=True).reset_index()
final_df_cal.columns=['RelativePeriod','Forecast_Year','Trailing_Months']
final_df_cal


# In[86]:


final_df=final_df.merge(final_df_cal,on=['Forecast_Year','Trailing_Months'])


# In[87]:


final_df


# In[97]:


border_value=final_df[(final_df['Forecast_Year']==fyear)&(final_df['Trailing_Months']==curr_month)]['RelativePeriod'].unique().tolist()[0]


# In[101]:


relperiodlist=[]
for i in range(border_value-2,border_value+1,1):
    print(i)
    relperiodlist.append(i)


# In[120]:


final_df['Estimate_MoM_Var']=final_df['Estimate_MoM_Var'].fillna(0)


# In[171]:


## write the function so that for every ticker - look up the peers from q comp and then go back to the final_df and filter for the dataframe with the comps and take the array and calculate the z score of the ticker for the current month 

import statistics
# period=
allzscore=[]
alltickers=[]
allperiods=[]
for tic in final_df['ticker'].unique().tolist():
    print(tic)
    
    for rp in relperiodlist:
        print(rp)
        df_t_samp = final_df[(final_df['ticker']==tic)&(final_df['RelativePeriod']<=rp)]
        if (len(df_t_samp)>12 and  len(df_t_samp['Estimate_MoM_Var'].unique().tolist())>2 ): ## this ensure that the we have enough data and also it eliminates one time mass downgrades 
            
            x=df_t_samp[(df_t_samp['RelativePeriod']==rp)]['Estimate_MoM_Var'].values[0] ## this is the x 

            all_T_values=df_t_samp['Estimate_MoM_Var'].tolist()
            
            med=statistics.median(all_T_values)
            stdev=statistics.stdev(all_T_values)
            try:
                z_score=(x-med)/(stdev)
            except ZeroDivisionError:
                pass 
                z_score=0
                
            allzscore.append(z_score)
            alltickers.append(tic)
            allperiods.append(rp)

        
        
    
    
    
    
    
    
    
#     break
df_z_score=pd.DataFrame()
df_z_score['Tickers']=alltickers
df_z_score['Periods']=allperiods
df_z_score['ZScore']=allzscore
df_z_score   
    
    
    


# In[172]:


all_T_values


# In[173]:


df_z_score_summary=df_z_score.groupby(['Tickers'])['ZScore'].mean().reset_index()
df_z_score_summary=df_z_score_summary.sort_values(by=['ZScore'])
df_z_score_summary


# In[174]:


df_z_score_summary.dropna(inplace=True)


# In[175]:


df_z_score[df_z_score['Tickers']=='NasdaqGM:EMBK']


# In[210]:





# In[136]:


df_short_score = pd.read_csv(r'Output.csv')
df_short_score


# In[176]:


df_z_score_plot=df_z_score_summary.rename(columns={'Tickers':'Ticker'}).merge(df_short_score[['Ticker','stock_price_ytd_return','short_score_sum']])
df_z_score_plot


# In[215]:


df_z_score_plot.to_csv(r"LatestRevisions_ZScore.csv")


# In[198]:


df_z_score_plot.loc[(df_z_score_plot['short_score_sum']>3)].sort_values(by=['ZScore'],ascending=True).head(20)


# In[214]:


df_z_score_plot.head(50).to_csv(r'AnomalousRev.csv')


# In[211]:


df_z_score_plot[df_z_score_plot['Ticker']=='NasdaqGS:GTLB']


# In[153]:


ticker_list=df_z_score_plot['Ticker'].tolist()


# In[178]:


import plotly.express as px
fig = px.scatter(df_z_score_plot,y='ZScore', x='short_score_sum')
fig.show()


# In[179]:


import plotly.graph_objects as go
import numpy as np


fig = go.Figure(data=go.Scatter(y=df_z_score_plot['ZScore'], x=df_z_score_plot['short_score_sum'], mode='markers',text=df_z_score_plot['Ticker']))
fig.show()


# In[163]:


dfss= pd.read_excel(r'Quick_Comps.xls',sheet_name='Sheet1')

                   


