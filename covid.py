import tkinter as tk
from tkinter import ttk
# from numpy.random import *

import pandas as pd
# import numpy as np
# import matplotlib.pyplot as plt
from os import path


def run(mainfile, dayRange):
    if not path.exists("data.pkl"):
        df = pd.read_excel(mainfile, header=2, skiprows=[3], skipfooter=2, engine='openpyxl')
        df.to_pickle("data.pkl")
    else:
        df = pd.read_pickle("data.pkl")
    # plt.rcParams['font.sans-serif'] = ['simhei']

    # In[48]:

    test = df.groupby(['生产企业', '居委会'])['受种者编码'].agg('count')
    test.to_csv("疫苗品种_居委会.csv")

    df.接种日期 = pd.to_datetime(df.接种日期)

    # df.接种日期.value_counts().plot(figsize=(10, 8))

    # df.接种医生.value_counts().head(10).plot.bar(figsize=(10, 8))

    # df.生产企业.value_counts().plot.bar(figsize=(10, 8))

    df.居委会.value_counts()

    # fixed variable
    dose1 = "新型冠状疫苗①"
    dose2 = "新型冠状疫苗②"
    today = pd.to_datetime("today")

    df_jici = df.groupby(['生产企业', '疫苗种类/剂次'])['受种者编码'].agg('count')
    df_jici.rename("人数").to_csv('疫苗品种_疫苗计次.csv')

    a = df.生产企业 == "北京科兴中维"
    b = df.生产企业 == "北京生物"
    doses = df[a | b]
    doses = doses.sort_values("接种日期")
    doses_1_2 = doses.drop_duplicates(subset=['受种者编码'], keep='last')
    doses_1 = doses_1_2[doses_1_2['疫苗种类/剂次'] == dose1]
    doses_2 = doses_1_2[doses_1_2['疫苗种类/剂次'] == dose2]

    out = doses_1[(today - doses_1.接种日期).astype('timedelta64[D]') > dayRange]

    doses_2["疫苗种类/剂次"].value_counts()

    df["疫苗种类/剂次"].value_counts()

    # In[66]:

    out.to_csv('脱漏30_1.csv')

    # In[ ]:

def main():
    def click():
        run(main_name.get(), int(dayRange.get()))
        print("Hi," + str(int(dayRange.get())))  # Textbox widget

    win = tk.Tk()  # Application Name
    win.geometry("500x200")
    win.title("新冠统计")  # Label
    lbl = ttk.Label(win, text="主要文件名").grid(column=0, row=0)
    lbl2 = ttk.Label(win, text="单日文件名").grid(column=0, row=1)
    lbl3 = ttk.Label(win, text="间隔天数").grid(column=0, row=2)
    lbl4 = ttk.Label(win, text="疫苗种类").grid(column=0, row=3)
    main_name = tk.StringVar()
    singleDay = tk.StringVar()
    dayRange = tk.StringVar()
    mainfile = ttk.Entry(win, width=40, textvariable=main_name)  # Button widget
    mainfile.grid(column=1, row=0)
    singleDayFile = ttk.Entry(win, width=40, textvariable=singleDay).grid(column=1, row=1)
    daterange = ttk.Entry(win, width=40, textvariable=dayRange).grid(column=1, row=2)
    tkvar = tk.StringVar()
    choices = ['选择种类', '全部', '北京科兴中维', '北京生物']
    tkvar.set(choices[0])  # set the default option
    popupMenu = ttk.OptionMenu(win, tkvar, *choices)
    popupMenu.grid(column=1, row=3)
    button = ttk.Button(win, text="submit", command=click).grid(column=1, row=4)
    win.mainloop()





if __name__ == '__main__':
    main()
