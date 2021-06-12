#!/usr/bin/env python
# coding: utf-8

# In[4]:


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from os import path


# In[45]:

def run(mainfile, dayRange):
    if not path.exists("data.pkl"):
        df = pd.read_excel(mainfile,header=2,skiprows=[3],skipfooter=2,engine='openpyxl')
        df.to_pickle("data.pkl")
    else:
        df = pd.read_pickle("data.pkl")
    plt.rcParams['font.sans-serif']=['simhei']


    # In[48]:


    test = df.groupby(['生产企业','居委会'])['受种者编码'].agg('count')
    test.to_csv("疫苗品种_居委会.csv")


    # In[49]:


    df.columns


    # In[50]:


    df.接种日期 = pd.to_datetime(df.接种日期)


    # In[51]:


    df.接种日期.value_counts().plot(figsize=(10, 8))


    # In[52]:



    df.接种医生.value_counts().head(10).plot.bar(figsize=(10, 8))


    # In[53]:


    df.生产企业.value_counts().plot.bar(figsize=(10, 8))


    # In[54]:


    df.居委会.value_counts()


    # In[55]:


    # fixed variable
    dose1 = "新型冠状疫苗①"
    dose2 = "新型冠状疫苗②"
    today = pd.to_datetime("today")


    # In[73]:


    df_jici = df.groupby(['生产企业','疫苗种类/剂次'])['受种者编码'].agg('count')
    df_jici.rename("人数").to_csv('疫苗品种_疫苗计次.csv')


    # In[58]:


    a = df.生产企业 == "北京科兴中维"
    b = df.生产企业 == "北京生物"
    doses = df[a | b]
    doses = doses.sort_values("接种日期")
    doses_1_2 = doses.drop_duplicates(subset=['受种者编码'],keep='last')
    doses_1 = doses_1_2[doses_1_2['疫苗种类/剂次'] == dose1]
    doses_2 = doses_1_2[doses_1_2['疫苗种类/剂次'] == dose2]


    # In[59]:




    # In[65]:


    out = doses_1[(today - doses_1.接种日期).astype('timedelta64[D]') > dayRange]


    # In[63]:


    doses_2["疫苗种类/剂次"].value_counts()


    # In[67]:

    # In[64]:


    df["疫苗种类/剂次"].value_counts()


    # In[66]:


    out.to_csv('脱漏30_1.csv')


    # In[ ]:




