from collections import defaultdict
from multiprocessing.dummy import current_process
from sqlite3 import connect
import pymysql
import tkinter as tk
import customtkinter as ctk
from tkinter import ttk, messagebox,CENTER
from ttkbootstrap import Style
import pandas as pd

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
    cursor.execute("SELECT `clsi規則`.編號,`clsi規則`.規則內容 FROM `clsi規則`;")
clia_rules = cursor.fetchall()
clia_rules = dict((y, x) for x, y in clia_rules)
clia_names = list(clia_rules.keys())
# print(clia_rules)
with conn.cursor() as cursor:
    cursor.execute("SELECT `測試件名稱`, `編號` FROM `測試件項目`;")
testnamedict = cursor.fetchall()
testnamedict = dict((x,y)for x,y in testnamedict)
class keyin_clia(object):
    def __init__(self):  
        
        #建立資料輸入介面  
        self.root = ctk.CTk()
        # ctk.set_default_color_theme("green")  
        # 給主視窗設定標題內容  
        self.root.title("能力試驗資料輸入")  
        self.root.geometry('800x800')
        style = Style(theme='cyborg') 
        self.root.config(background='#323232') #設定背景色
        self.root.bind('<Return>', self.return_click)
        #設定label中變數
        self.year = tk.StringVar()
        self.testname = tk.StringVar()
        self.testnum_1 = tk.StringVar()
        self.testnum_2 = tk.StringVar()
        self.testobj = tk.StringVar()
        self.previous = tk.StringVar()
        #設定觀看視窗框架
        self.frame1 = tk.LabelFrame(self.root, text="Raw Data", foreground="#323232")
        self.frame1.place(height=400, width=750)
        self.label_year = ctk.CTkLabel(
            self.root, 
            text = "年份:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=120
            )
        self.tv_year = ctk.CTkLabel(
            self.root, 
            textvariable=self.year, 
            fg_color='#F0F0F0',
            text_font=('微軟正黑體',16),
            text_color="#000000",
            width=120
            )
        self.tv_year.pack()
        self.label_testname = ctk.CTkLabel(
            self.root, 
            text = "測試件名稱:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=150
            )
        self.tv_testname = ctk.CTkLabel(
            self.root, 
            textvariable=self.testname, 
            fg_color='#F0F0F0',
            text_font=('微軟正黑體',16),
            text_color="#000000",
            width=120
            )
        self.tv_testname.pack()
        self.label_testnum_1 = ctk.CTkLabel(
            self.root, 
            text = "年度第幾次:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=150
            )
        self.tv_testnum_1 = ctk.CTkLabel(
            self.root, 
            textvariable=self.testnum_1, 
            fg_color='#F0F0F0',
            text_font=('微軟正黑體',16),
            text_color="#000000",
            width=120
            )    
        self.label_testnum_2 = ctk.CTkLabel(
            self.root, 
            text = "測試件序號:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=150
            )
        self.tv_testnum_2 = ctk.CTkLabel(
            self.root, 
            textvariable=self.testnum_2, 
            fg_color='#F0F0F0',
            text_font=('微軟正黑體',16),
            text_color="#000000",
            width=120
            )
        self.label_testobj = ctk.CTkLabel(
            self.root, 
            text = "測試分項目:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=150
            )
        self.tv_testobj = ctk.CTkLabel(
            self.root, 
            textvariable=self.testobj, 
            fg_color='#F0F0F0',
            text_font=('微軟正黑體',16),
            text_color="#000000",
            width=120
            )
        self.label_previous = ctk.CTkLabel(
            self.root, 
            text = "上次能力試驗範圍", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=200
            )
        self.input_previous = ctk.CTkLabel(
            self.root, 
            textvariable=self.previous, 
            fg_color='#F0F0F0',
            text_font=('微軟正黑體',16),
            text_color="#000000",
            width=220
            )
        self.label_claival = ctk.CTkLabel(
            self.root, 
            text = "CLAI範圍:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=150
            )
        self.input_claival = ttk.Combobox(
            master = self.root, 
            # textvariable=self.input_testname.get(),
            values=clia_names,
            width=15,height=40,
            state="readonly"
            )
        self.button_OK=ctk.CTkButton(
            self.root, 
            command = self.OK_interface, 
            text = "確定", 
            fg_color='#666666',
            width=180,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.button_OK.pack()
        self.button_back=ctk.CTkButton(
            self.root, 
            command = self.back_interface, 
            text = "返回", 
            fg_color='#666666',
            width=180,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.button_back.pack()
        self.button_OKs=ctk.CTkButton(
            self.root, 
            command = self.OKs_interface, 
            text = "批次確定", 
            fg_color='#666666',
            width=180,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.button_OKs.pack()
        ###建立treeview(可視覺化excel)###
        self.tv1 = ttk.Treeview(self.frame1)
        self.tv1.place(relheight=1, relwidth=1)
        treescrolly = ttk.Scrollbar(self.frame1,orient="vertical", command=self.tv1.yview) 
        treescrollx = ttk.Scrollbar(self.frame1,orient="horizontal", command=self.tv1.xview)
        self.tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) 
        treescrollx.pack(side="bottom", fill="x")
        treescrolly.pack(side="right", fill="y")
        
    def add_data(self):	#更新treeview頁面
        for row in self.tv1.get_children():
            self.tv1.delete(row)
        with conn.cursor() as cursor:
            clai = """SELECT `能力試驗結果`.`測試件結果編號`,`測試件結果`.`年份`,`測試件結果`.`年度次數`,`測試件項目`.`測試件名稱`,`測試件分項目`.`測試項目_分項`,`測試件結果`.`測試件序號`,`能力試驗結果`.`能力試驗數值`,`能力試驗結果`.`能力試驗標準差`,`能力試驗結果`.`CLSI規則` FROM `能力試驗結果` 
                    JOIN `測試件結果`
                    ON `測試件結果`.`結果編號` = `能力試驗結果`.`測試件結果編號` 
                    JOIN `測試件項目`
                    ON `測試件結果`.`測試件項目編號` = `測試件項目`.`編號`
                    JOIN `測試件分項目`
                    ON`測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號`
                    WHERE `CLSI規則` IS NULL;"""
            cursor.execute(clai)
        clainull = cursor.fetchall()
        df = pd.DataFrame(clainull)
        df.columns = ["編號","年份","年度次數","測試件名稱","測試件分項目","測試件結果編號","能力試驗數值","能力試驗標準差","CLSI規則"]
        # print(df)
        # print(clainull)
        self.tv1["column"] = list(df.columns)
        self.tv1["show"] = "headings"
        for column in self.tv1["columns"]:
            self.tv1.heading(column, text=column) # let the column heading = column name
            self.tv1.column(column, width= 80, anchor=CENTER)
        df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
        for row in df_rows:
            self.tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
        self.root.bind('<ButtonRelease-1>', self.clicktreeview)

    def gui_arrang(self):
        self.frame1.place(relx=0.5,rely=0.3,anchor=tk.CENTER)
        self.label_year.place(relx=0.085, rely=0.6, anchor=tk.W)
        self.tv_year.place(relx=0.25, rely=0.6,anchor=tk.W)
        self.label_testname.place(relx=0.025,rely=0.675,anchor=tk.W)
        self.tv_testname.place(relx=0.25,rely=0.675,anchor=tk.W)
        self.label_testobj.place(relx=0.01,rely=0.75,anchor=tk.W)
        self.tv_testobj.place(relx=0.25,rely=0.75,anchor=tk.W)
        self.label_testnum_1.place(relx=0.025,rely=0.825,anchor=tk.W)
        self.tv_testnum_1.place(relx=0.25,rely=0.825,anchor=tk.W)
        self.label_testnum_2.place(relx=0.025,rely=0.9,anchor=tk.W)
        self.tv_testnum_2.place(relx=0.25,rely=0.9,anchor=tk.W)
        self.label_previous.place(relx=0.84,rely=0.6,anchor=tk.E)
        self.input_previous.place(relx=0.85,rely=0.65,anchor=tk.E)
        self.label_claival.place(relx=0.81,rely=0.7,anchor=tk.E)
        self.input_claival.place(relx=0.8,rely=0.75,anchor=tk.E)
        self.button_OK.place(relx=0.825,rely=0.825,anchor=tk.E)
        self.button_back.place(relx=0.825,rely=0.895,anchor=tk.E)
        self.button_OKs.place(relx=0.825,rely=0.965,anchor=tk.E)
    def clicktreeview(self,event):
        for item in self.tv1.selection():
            item_data = self.tv1.item(item,"values")
            # print(item_data)
        self.year.set(item_data[1]) #更改label_year中內容
        self.testname.set(item_data[3])		#更改label_testname中內容
        self.testnum_1.set(item_data[2])	#更改label_testnum1(年度次數)中內容
        self.testnum_2.set(item_data[5])	#更改label_testnum2(測試件序號)中內容
        self.testobj.set(item_data[4])		#更改label_testobj中內容
        pastyear = int(self.year.get())	#取得今年年份
        pastyear = pastyear - 1	#去年 
        with conn.cursor() as cursor:
            find_privious = """SELECT `測試件結果`.`年份`,`clsi規則`.`規則內容`
                                FROM (`測試件結果` INNER JOIN `測試件分項目` ON `測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號`)
                                JOIN `能力試驗結果`
                                ON `測試件結果`.`結果編號` = `能力試驗結果`.`測試件結果編號`
                                JOIN `clsi規則`
                                ON `能力試驗結果`.`CLSI規則` = `clsi規則`.`編號`
                                WHERE `能力試驗結果`.`測試件結果編號` IN  (
                                    SELECT `結果編號` FROM `測試件結果` WHERE`測試件結果`.`測試件分項目編號`IN(
                                        SELECT `分項編號` FROM `測試件分項目` WHERE `編號` = %d AND `測試項目_分項` = '%s'))
                                AND `測試件結果`.`年份` = '%s'; """%(testnamedict[self.testname.get()],self.testobj.get(),pastyear)
            cursor.execute(find_privious)	#查找去年度能力試驗之可容許範圍
        find_privious = cursor.fetchone()
        past = find_privious[1]	#去年度能力試驗可容許範圍值
        self.previous.set(past)
    def OK_interface(self):
        curItem = self.tv1.focus()
        rowval = self.tv1.item(curItem)['values']
        # print (rowval)
        claival = str(self.input_claival.get())
        try: 
            claiidx = clia_rules[claival]
        except KeyError:
            tk.messagebox.showerror(title='南投醫院檢驗科', message='尚未選擇可容許範圍!')
            return
        askmsg="""確認上傳CLAI可容許範圍
年份: %d
測試件名稱: %s
測試件分項目: %s
年度第幾次: %d
測試件序號: %d
CLAI可容許範圍: %s"""%(rowval[1],rowval[3],rowval[4],rowval[2],rowval[5],claival)
        if tk.messagebox.askyesno(title='南投署立醫院檢驗科', message = askmsg):
            with conn.cursor() as cursor:
                clai_update = """UPDATE `能力試驗結果` 
                SET `CLSI規則` = %d 
                WHERE `測試件結果編號` = '%s';"""%(claiidx,rowval[0])
                cursor.execute(clai_update)
            conn.commit()
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message = "上傳成功!")
            self.add_data()
        else :
            return
    def OKs_interface(self):
        curItem = self.tv1.focus()
        rowval = self.tv1.item(curItem)['values']
        claival = str(self.previous.get())
        claiidx = clia_rules[claival]
        for row in self.tv1.get_children():
            self.tv1.delete(row)
        askmsg="""確認沿用上一次CLAI可容許範圍?
測試件名稱: %s
測試件分項目: %s
CLAI可容許範圍: %s"""%(rowval[3],rowval[4],claival)
        with conn.cursor() as cursor:
            filt = """SELECT `能力試驗結果`.`測試件結果編號`,`測試件結果`.`年份`,`測試件結果`.`年度次數`,`測試件項目`.`測試件名稱`,`測試件分項目`.`測試項目_分項`,`測試件結果`.`測試件序號`,`能力試驗結果`.`能力試驗數值`,`能力試驗結果`.`能力試驗標準差`,`能力試驗結果`.`CLSI規則` FROM `能力試驗結果` 
                    JOIN `測試件結果`
                    ON `測試件結果`.`結果編號` = `能力試驗結果`.`測試件結果編號` 
                    JOIN `測試件項目`
                    ON `測試件結果`.`測試件項目編號` = `測試件項目`.`編號`
                    JOIN `測試件分項目`
                    ON`測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號`
                    WHERE `CLSI規則` IS NULL AND `測試件項目`.`測試件名稱`='%s' AND `測試件分項目`.`測試項目_分項`='%s';"""%(rowval[3],rowval[4])
            cursor.execute(filt)
            filt = cursor.fetchall()
        df = pd.DataFrame(filt)
        df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
        for row in df_rows:
            self.tv1.insert("", "end", values=row)
        print(filt)
        if tk.messagebox.askyesno(title='南投署立醫院檢驗科', message = askmsg):
            for row in range(0,len(df)):
                # print(df.iat[row,0])
                with conn.cursor() as cursor:
                    clai_update = """UPDATE `能力試驗結果` 
                    SET `CLSI規則` = %d 
                    WHERE `測試件結果編號` = '%s';"""%(claiidx,df.iat[row,0])
                # print(clai_update)
                    cursor.execute(clai_update)
                conn.commit()
            self.add_data()
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message = "上傳成功!")
        else :
            self.add_data()
            return
    def back_interface(self):
        self.root.destroy()    
    def return_click(self, event):  #按Enter鍵自動連結登入
        self.OKs_interface()

def main():  
    K = keyin_clia()
    K.gui_arrang()
    K.add_data()
    # 主程式執行  
    tk.mainloop()

if __name__ == '__main__':  
    main()  