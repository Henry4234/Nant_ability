#authorised by Henry Tsai
import os
from pdb import line_prefix
from sysconfig import get_paths
import tkinter as tk
from tkinter import CENTER, messagebox,ttk, StringVar
import pandas as pd
from pandas import DataFrame, value_counts
import numpy as np
from pkg_resources import register_finder
from ttkbootstrap import Style
import customtkinter  as ctk
from collections import defaultdict
import subprocess
import pymysql
import matplotlib.pyplot as plt
#建立與mySQL連線資料
db_settings = { 
    "host": "127.0.0.1",
    "port": 3306,
    "user": "root",
    "password": "ROOT",
    "db": "nantou db",
    "charset": "utf8"
    }
try:
    conn = pymysql.connect(**db_settings)
except pymysql.err.OperationalError:
    db_settings.update({"host": "192.168.0.120","port": 3307})
    del db_settings["password"]
finally:
    conn = pymysql.connect(**db_settings)

with conn.cursor() as cursor:
    cursor.execute("SELECT `測試件項目`.測試件名稱, `測試件分項目`.測試項目_分項 FROM `測試件項目`, `測試件分項目` WHERE (`測試件項目`.編號 = `測試件分項目`.編號);")
result_subtestname = cursor.fetchall()
with conn.cursor() as cursor:
    cursor.execute("SELECT DISTINCT `測試件結果`.`年份`FROM `測試件結果`ORDER BY `測試件結果`.`年份` ASC")
year = list(cursor.fetchall())
with conn.cursor() as cursor:
    cursor.execute("SELECT `測試件名稱`, `編號` FROM `測試件項目`;")
testname_dict = cursor.fetchall()
result_subtestname=tuple((y, x) for x, y in result_subtestname)
testname_dict = dict((x,y)for x,y in testname_dict)
subtestname = defaultdict(list)
for i,j in result_subtestname:
    subtestname[j].append(i)
sorted(subtestname.items())
result=list(subtestname.keys())

class dashboard(object):

    def __init__(self):
        #建立資料輸入介面  
        self.root = ctk.CTk()
        # ctk.set_default_color_theme("green")  
        # 給主視窗設定標題內容  
        self.root.title("能力試驗結果查看")
        self.root.geometry('1200x800')
        style = Style(theme='cyborg') 
        self.root.config(background='#323232') #設定背景色  
        self.hellow_label = ctk.CTkLabel(
            self.root, 
            text = "能力試驗結果查看",
            fg_color='#323232',
            text_font=('微軟正黑體',20),
            text_color="#00F5FF",
            width=260
            )
        self.label_year = ctk.CTkLabel(
            self.root, 
            text = "年份:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            # width=120,
            anchor=tk.E
            )
        self.label_testname = ctk.CTkLabel(
            self.root, 
            text = "測試件名稱:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            # width=150
            )
        self.label_testnum = ctk.CTkLabel(
            self.root, 
            text = "年度第幾次:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            # width=150
            )
        self.label_testobj = ctk.CTkLabel(
            self.root, 
            text = "能力試驗項目:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            # width=150
            )
        self.label_main = ctk.CTkLabel(
            self.root, 
            text = "主備機:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            # width=150
            )
        self.input_year = ttk.Combobox(
            master = self.root, 
            values=year,
            width=15,height=40,
            state="readonly"
            )
        self.input_year.bind("<<ComboboxSelected>>", self.update)
        self.input_testnum = ttk.Combobox(
            master = self.root, 
            values=[1,2,3],
            width=15,height=40,
            state="readonly"
            )
        self.input_testnum.bind("<<ComboboxSelected>>", self.update)
        self.input_testname = ttk.Combobox(
            master = self.root, 
            values=result,
            width=15,height=40,
            state="readonly"
            )
        self.input_testname.bind("<<ComboboxSelected>>", lambda event:(self.callback(event),self.update(event)))
        self.input_testobj = ttk.Combobox(
            master = self.root, 
            # textvariable=self.input_testname.get(),
            values=subtestname[self.input_testname.get()],
            width=15,height=40,
            state="readonly"
            )
        self.input_testobj.bind("<<ComboboxSelected>>", self.update)
        self.input_main = ttk.Combobox(
            master = self.root, 
            values=["主機","備機"],
            width=15,height=40,
            state="readonly"
            )
        self.input_main.bind("<<ComboboxSelected>>", self.update)

        self.frame1 = tk.LabelFrame(self.root, text="Raw Data", foreground="#323232")
        self.frame1.place(height=200, width=700)

        ###建立treeview(可視覺化excel)###
        self.tv1 = ttk.Treeview(self.frame1)
        self.tv1.place(relheight=1, relwidth=1)
        treescrolly = ttk.Scrollbar(self.frame1,orient="vertical", command=self.tv1.yview) # command means update the yaxis view of the widget
        treescrollx = ttk.Scrollbar(self.frame1,orient="horizontal", command=self.tv1.yview) # command means update the xaxis view of the widget
        self.tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # assign the scrollbars to the Treeview Widget
        treescrollx.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
        treescrolly.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget
    def update(self,event):
        year = self.input_year.get()
        testnum =  int(self.input_testnum.get())
        testname = testname_dict[self.input_testname.get()]
        testobj = self.input_testobj.get()
        main = self.input_main.get()
        if main =="主機":
            srch_db="""SELECT `測試件結果`.`結果編號`, `測試件結果`.`測試件序號`,`測試件結果`.`測試件結果`, `能力試驗結果`.`能力試驗數值`,`測試件結果`.`不等判讀` 
                    FROM `測試件結果`
                    JOIN `測試件項目`
                    ON `測試件結果`.`測試件項目編號` = `測試件項目`.`編號` 
                    JOIN `測試件分項目`
                    ON`測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號`
                    JOIN `能力試驗結果`
                    ON `測試件結果`.`結果編號` = `能力試驗結果`.`測試件結果編號`
                    WHERE `測試件結果`.`測試件分項目編號` IN  (SELECT `分項編號` FROM `測試件分項目` WHERE `編號` = %d AND `測試項目_分項` = '%s')
                    AND `測試件結果`.`年份` = %s 
                    AND `測試件結果`.`年度次數` = %d
                    ORDER BY `測試件序號` ASC;"""%(testname,testobj,year,testnum)
        else:
            srch_db="""SELECT `測試件結果`.`結果編號`, `測試件結果`.`測試件序號`,`測試件結果`.`測試件結果_備機`, `能力試驗結果`.`能力試驗數值`,`測試件結果`.`不等判讀_備機` 
                    FROM `測試件結果`
                    JOIN `測試件項目`
                    ON `測試件結果`.`測試件項目編號` = `測試件項目`.`編號` 
                    JOIN `測試件分項目`
                    ON`測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號`
                    JOIN `能力試驗結果`
                    ON `測試件結果`.`結果編號` = `能力試驗結果`.`測試件結果編號`
                    WHERE `測試件結果`.`測試件分項目編號` IN  (SELECT `分項編號` FROM `測試件分項目` WHERE `編號` = %d AND `測試項目_分項` = '%s')
                    AND `測試件結果`.`年份` = %s 
                    AND `測試件結果`.`年度次數` = %d
                    ORDER BY `測試件序號` ASC;"""%(testname,testobj,year,testnum)
        with conn.cursor() as cursor:
            cursor.execute(srch_db)
        df = cursor.fetchall()
        df = pd.DataFrame(df)
        title = ["結果編號","測試件序號","測試件結果","能力試驗結果","不等判讀"]
        try:
            df.columns=title
        except ValueError:
            self.clear_data()
        print(df)
        self.clear_data()
        self.tv1["column"] = list(df.columns)
        self.tv1["show"] = "headings"
        for column in self.tv1["columns"]:
            self.tv1.heading(column, text=column) # let the column heading = column name
            self.tv1.column(column, width= 80, anchor=CENTER)
        df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
        for row in df_rows:
            self.tv1.insert("", "end", values=row)
    def callback(self,event):  #combobox雙層列表
        self.testobj = StringVar()
        self.testobj=subtestname[self.input_testname.get()]
        self.input_testobj.configure(values = subtestname[self.input_testname.get()])
    def clear_data(self):
        self.tv1.delete(*self.tv1.get_children())
        return None
        	
    def gui_arrang(self):
        self.hellow_label.grid(column=0, row=0, columnspan=2, rowspan=2, ipadx=5, ipady=15, sticky=tk.W+tk.N)
        self.label_year.grid(column=0, row=2, padx=15, pady=15, sticky=tk.E+tk.N)
        self.label_testnum.grid(column=0, row=3, ipadx=15, pady=15, sticky=tk.E+tk.N)
        self.label_testname.grid(column=0, row=4, ipadx=15, pady=15, sticky=tk.E+tk.N)
        self.label_testobj.grid(column=0, row=5, ipadx=15, pady=15, sticky=tk.E+tk.N)
        self.label_main.grid(column=0, row=6, ipadx=15, pady=15, sticky=tk.E+tk.N)
        self.input_year.grid(column=1, row=2, padx=20, pady=15,)
        self.input_testnum.grid(column=1, row=3, padx=20, pady=15,)
        self.input_testname.grid(column=1, row=4, padx=20, pady=15,)
        self.input_testobj.grid(column=1, row=5, padx=20, pady=15,)
        self.input_main.grid(column=1, row=6, padx=20, pady=15,)
        self.frame1.place(relx=0.65,rely=0.25,anchor=tk.CENTER)
def main():  
    D = dashboard()
    D.gui_arrang()
    # 主程式執行  
    tk.mainloop()

if __name__ == '__main__':  
    main()  