import tkinter as tk
from tkinter import ttk
# from numpy.random import *

import pandas as pd
# import numpy as np
# import matplotlib.pyplot as plt
from os import path

newMenu = []


# df = []
def run_import(win, tkvar, mainfile,menzheng):
    # popMenu.destroy()
    df = []
    if not path.exists("data.pkl"):
        df = pd.read_excel(mainfile, header=2, skiprows=[3], skipfooter=2, engine='openpyxl')
        df.生产企业 = df.生产企业.fillna('未注明厂商')
        df.to_pickle("data.pkl")
        print("新建数据存档data.pkl")
    else:
        print("读取已存记录")
        df = pd.read_pickle("data.pkl")

        if mainfile != "":
            df_single = pd.read_excel(mainfile, header=2, skiprows=[3], skipfooter=2, engine='openpyxl')
            df = pd.concat([df, df_single], axis=0, ignore_index=True)
            df.生产企业 = df.生产企业.fillna('未注明厂商')
            df = df.drop_duplicates()
            print("写入新数据")
            df.to_pickle("data.pkl")
            print("添加新纪录")
    choices_new = df.生产企业.unique().tolist()
    choices_new.insert(0, '选择种类')
    choices_new.insert(1, '全部')
    menzheng.set(df.归属门诊.value_counts().index[0])
    tkvar.set(choices_new[0])
    newMenu = ttk.OptionMenu(win, tkvar, *choices_new)
    newMenu.grid(column=1, row=4)
    return


def runExport(dayRange, singleDay, doseType, menzheng):
    if menzheng != "":
        with open('title.txt', 'w') as myfile:
            myfile.seek(0)
            myfile.write(menzheng)
            myfile.truncate()
            myfile.close()
        df = pd.read_pickle("data.pkl")
        df.接种日期 = pd.to_datetime(df.接种日期)
        df_jiezhong = df[df.接种门诊名称 == menzheng]
        # 单日统计
        if singleDay != "":
            selectedDay = pd.to_datetime(singleDay)
            df_danrijici = df_jiezhong[df_jiezhong.接种日期 == selectedDay]
            print(selectedDay)
            print("开始单日统计")

            df_danrijici = df_danrijici.groupby(['生产企业', '疫苗种类/剂次'])['受种者编码'].agg('count')
            df_danrijici = df_danrijici.rename("人数")
            df_danrijici.to_csv('单日' + singleDay + '.csv')
            print("单日统计完成")
        print(df.tail(5))
        dose1 = "新型冠状疫苗①"
        dose2 = "新型冠状疫苗②"
        today = pd.to_datetime("today")
        df_jici = df_jiezhong.groupby(['生产企业', '疫苗种类/剂次'])['受种者编码'].agg('count')
        df_jici = df_jici.rename("人数")
        df_jici.to_csv('全部疫苗品种_疫苗计次.csv')
        if dayRange != "":
            df_guishu = df[df.归属门诊 == menzheng]
            dayRange = int(dayRange)
            # fixed variable
            if doseType == "选择种类":
                print("未选择种类")
            else:
                doses = df_guishu.sort_values("接种日期")
                doses_1_2 = doses.drop_duplicates(subset=['受种者编码'], keep='last')
                doses_1 = doses_1_2[doses_1_2['疫苗种类/剂次'] == dose1]
                out = doses_1[(today - doses_1.接种日期).astype('timedelta64[D]') >= dayRange]
                if doseType != "全部":
                    out = out[out.生产企业 == doseType]
                out.to_csv(doseType + '接种超过' + str(dayRange) + '天.csv')
                quanbu_out_renshu = out.生产企业.value_counts().rename_axis('生产企业').reset_index(name='人数')
                quanbu_out_renshu.to_csv(doseType + '接种超过' + str(dayRange) + '天_疫苗人数.csv')

        print("已完成统计")  # Textbox widget

    return


def main():
    def click():
        runExport(dayRange.get(), singleDay.get(), tkvar.get(), menzheng.get())

    def click_import():
        run_import(win, tkvar, main_name.get(),menzheng)

    win = tk.Tk()  # Application Name
    win.geometry("500x200")
    win.title("新冠统计")  # Label
    lbl = ttk.Label(win, text="添加文件").grid(column=0, row=0)
    lbl2 = ttk.Label(win, text="日期：年-月-日，2021-01-01").grid(column=0, row=1)
    lbl3 = ttk.Label(win, text="间隔天数").grid(column=0, row=2)
    lbl5 = ttk.Label(win, text="门诊名称").grid(column=0, row=3)
    lbl4 = ttk.Label(win, text="疫苗种类(脱漏)").grid(column=0, row=4)
    lbl5 = ttk.Label(win, text="黏贴名称使用ctrl + v").grid(column=0, row=6)
    main_name = tk.StringVar()
    singleDay = tk.StringVar()
    dayRange = tk.StringVar()
    menzheng = tk.StringVar()
    if path.exists("title.txt"):
        with open('title.txt', 'r') as file:
            data = file.read().replace('\n', '')
            menzheng.set(data)
            print(data)
            file.close()

    mainfile = ttk.Entry(win, width=40, textvariable=main_name)  # Button widget
    mainfile.grid(column=1, row=0)
    selectedDate = ttk.Entry(win, width=40, textvariable=singleDay).grid(column=1, row=1)
    daterange = ttk.Entry(win, width=40, textvariable=dayRange).grid(column=1, row=2)

    menzhengEntry = ttk.Entry(win, width=40, textvariable=menzheng).grid(column=1, row=3)

    tkvar = tk.StringVar()
    if path.exists("data.pkl"):
        df = pd.read_pickle("data.pkl")
        choices_new = df.生产企业.unique().tolist()
        choices_new.insert(0, '选择种类')
        choices_new.insert(1, '全部')

        tkvar.set(choices_new[0])
        newMenu = ttk.OptionMenu(win, tkvar, *choices_new)
        newMenu.grid(column=1, row=4)

    button = ttk.Button(win, text="导出", command=click).grid(column=1, row=5)
    button_import = ttk.Button(win, text="导入", command=click_import).grid(column=0, row=5)

    win.mainloop()


if __name__ == '__main__':
    main()
