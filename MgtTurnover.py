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


