import os
from pdb import line_prefix
from sysconfig import get_paths
import tkinter as tk
from tkinter import CENTER, END, messagebox,ttk, StringVar
from turtle import bgcolor
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

class Modify(object):
    
    def __init__(self):
        #建立資料輸入介面  
        self.root = ctk.CTk()
        # ctk.set_default_color_theme("green")  
        # 給主視窗設定標題內容  
        self.root.title("能力試驗結果查看")
        self.root.geometry('1200x800')
        style = Style(theme='cyborg') 
        self.root.config(background='#323232') #設定背景色 
        ##設定選取資料後填入框架
        self.frame2 = tk.LabelFrame(self.root, text="選取資料", foreground="#323232")
        self.frame2.config(background="#474747")
        self.frame2.place(height=300, width=500)
        ##設定label中變數
        self.s_testname = tk.StringVar()
        self.s_testnum_1 = tk.StringVar()
        self.s_testresult = tk.StringVar()
        self.s_testmean = tk.StringVar()
        self.hellow_label = ctk.CTkLabel(
            self.root, 
            text = "能力試驗結果修改",
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
        self.input_testname.bind("<<ComboboxSelected>>", self.callback)
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
        ##選擇/修改按鈕
        self.button_select = ctk.CTkButton(
            self.root, 
            command = self.clicktreeview, 
            text = "選擇", 
            fg_color='#666666',
            width=180,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.button_modify = ctk.CTkButton(
            self.root, 
            command = self.modify, 
            text = "修改", 
            fg_color='#666666',
            width=80,height=90,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        ##選擇後標籤
        self.l_serial_number = ctk.CTkLabel(
            self.frame2, 
            text = "編號:", 
            fg_color='#474747',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            # width=120,
            anchor=tk.E
            )
        self.s_N = ctk.CTkLabel(
            self.frame2, 
            textvariable=self.s_testname, 
            fg_color='#F0F0F0',
            text_font=('微軟正黑體',16),
            text_color="#000000",
            width=220
            )
        self.l_testnum_1 = ctk.CTkLabel(
            self.frame2, 
            text = "序號:", 
            fg_color='#474747',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            # width=120,
            anchor=tk.E
            )
        self.s_tn_1 = ctk.CTkLabel(
            self.frame2, 
            textvariable=self.s_testnum_1, 
            fg_color='#F0F0F0',
            text_font=('微軟正黑體',16),
            text_color="#000000",
            width=220
            )
        self.l_testnum_2 = ctk.CTkLabel(
            self.frame2, 
            text = "測試件結果:", 
            fg_color='#474747',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=170,
            anchor=tk.E
            )
        self.s_tn_2 = ctk.CTkEntry(
            self.frame2, 
            fg_color='#F0F0F0',
            text_color="#000000",
            width=220
            )
        self.l_testnum_3 = ctk.CTkLabel(
            self.frame2, 
            text = "能力試驗結果:", 
            fg_color='#474747',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=170,
            anchor=tk.E
            )
        self.s_tn_3 = ctk.CTkEntry(
            self.frame2, 
            fg_color='#F0F0F0',
            text_color="#000000",
            width=220
            )
        self.frame1 = tk.LabelFrame(self.root, text="Raw Data", foreground="#323232")
        self.frame1.place(height=200, width=700)
        
        ###建立treeview(可視覺化excel)###
        self.tv1 = ttk.Treeview(self.frame1)
        self.tv1.place(relheight=1, relwidth=1)
        treescrolly = ttk.Scrollbar(self.frame1,orient="vertical", command=self.tv1.yview) # command means update the yaxis view of the widget
        treescrollx = ttk.Scrollbar(self.frame1,orient="horizontal", command=self.tv1.yview) # command means update the xaxis view of the widget
        self.tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # assign the scrollbars to the Treeview Widget
        treescrollx.pack(side="bottom", fill="x")
        treescrolly.pack(side="right", fill="y")
    
    def update(self,*event):
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
        df.columns=title
        # print(df)
        self.clear_data()
        self.tv1["column"] = list(df.columns)
        self.tv1["show"] = "headings"
        for column in self.tv1["columns"]:
            self.tv1.heading(column, text=column) # let the column heading = column name
            self.tv1.column(column, width= 80, anchor=CENTER)
        df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
        for row in df_rows:
            self.tv1.insert("", "end", values=row)
        self.tv1.bind('<Double-Button-1>', self.clicktreeview)
    def callback(self,event):  #combobox雙層列表
        self.testobj = StringVar()
        self.testobj=subtestname[self.input_testname.get()]
        self.input_testobj.configure(values = subtestname[self.input_testname.get()])
    def select(self,event):
        # print("YYY")
        self.s_tn_2.select_range(0, END)
        self.s_tn_2.focus()
        pass
    def clear_data(self):
        self.tv1.delete(*self.tv1.get_children())
        return None
    def clicktreeview(self,*event):
        for item in self.tv1.selection():
            item_data = self.tv1.item(item,"values")
            print(item_data)
        self.s_testname.set(item_data[0]) #更改label_year中內容
        self.s_testnum_1.set(item_data[1])		#更改label_testname中內容
        self.s_tn_2.delete(0,"end")
        self.s_testresult.set(item_data[2])	#更改label_testnum1(年度次數)中內容
        self.s_tn_2.insert(0,self.s_testresult.get()) 
        self.s_tn_3.delete(0,"end")
        self.s_testmean.set(item_data[3])	#更改label_testnum2(測試件序號)中內容
        self.s_tn_3.insert(0,self.s_testmean.get())
        # pastyear = int(self.year.get())	#取得今年年份
    
    def modify(self):
        ##取得年份/年度次數/能力試驗項目/能力試驗分項目
        year = int(self.input_year.get())   #年份
        year_num = int(self.input_testnum.get())    #年度次數
        testname = str(self.input_testname.get())   #測試件項目
        testobj = str(self.input_testobj.get()) #測試件分項目
        testnum = int(self.s_testnum_1.get())   #測試件序號
        main = str(self.input_main.get())   #主備機
        result_val = float(self.s_tn_2.get())   #測試件結果數值
        ability_val = float(self.s_tn_3.get())   #能力試驗結果數值
        s_testname = str(self.s_testname.get()) #結果編號
        msg = """確認修改以下內容:
年份: %s
年度次數: %d
能力試驗項目: %s
能力試驗分項目: %s
序號: %s
主備機: %s
測試件數值: %.5f
能力試驗數值: %.5f"""%(year,year_num,testname,testobj,testnum,main,result_val,ability_val)
        if main =="主機":
            main=""
        elif main=="備機":
            main="_備機"
        if tk.messagebox.askyesno(title='南投署立醫院檢驗科', message=msg):
            with conn.cursor() as cursor: 
                sql_valmodify = """UPDATE `測試件結果` SET `測試件結果`.`測試件結果%s`=%.5f WHERE `測試件結果`.`結果編號`='%s';"""%(main,result_val,s_testname) 
                cursor.execute(sql_valmodify)
                sql_ablmodify = """UPDATE `能力試驗結果` SET `能力試驗結果`.`能力試驗數值`=%.5f WHERE `能力試驗結果`.`測試件結果編號`='%s';"""%(ability_val,s_testname)
                cursor.execute(sql_ablmodify)
            conn.commit()
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='修改成功!')
            self.update()
        else:
            self.update()
            return

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
        self.l_serial_number.grid(column=0, row=25, padx=5, pady=15, sticky=tk.E+tk.S)
        self.s_N.grid(column=1, row=25, padx=5, pady=15, sticky=tk.E+tk.S)
        self.l_testnum_1.grid(column=0, row=26, padx=5, pady=15, sticky=tk.E+tk.S)
        self.s_tn_1.grid(column=1, row=26, padx=5, pady=15, sticky=tk.E+tk.S)
        self.l_testnum_2.grid(column=0, row=27, padx=5, pady=15, sticky=tk.E+tk.S)
        self.s_tn_2.grid(column=1, row=27, padx=5, pady=15, sticky=tk.E+tk.S)
        self.l_testnum_3.grid(column=0, row=28, padx=5, pady=15, sticky=tk.E+tk.S)
        self.s_tn_3.grid(column=1, row=28, padx=5, pady=15, sticky=tk.E+tk.S)
        self.button_select.place(relx=0.65,rely=0.42,anchor=tk.CENTER)
        self.button_modify.place(relx=0.5,rely=0.7,anchor=tk.CENTER)
        self.frame1.place(relx=0.65,rely=0.25,anchor=tk.CENTER)
        self.frame2.place(relx=0.25,rely=0.7,anchor=tk.CENTER)
def main():  
    M = Modify()
    M.gui_arrang()
    # 主程式執行  
    tk.mainloop()

if __name__ == '__main__':  
    main()  