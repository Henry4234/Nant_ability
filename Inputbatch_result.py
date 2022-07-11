#authorised by Henry Tsai
import sys
import tkinter as tk
from tkinter import CENTER, ttk, StringVar
from tkinter import filedialog
from turtle import width
from ttkbootstrap import Style
import customtkinter  as ctk
from collections import defaultdict
import subprocess
import pymysql
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

global account
account = sys.argv[1]
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

#與mySQL建立連線，取出測試件項目工作表中的測試件名稱以及編號
with conn.cursor() as cursor:
    test_db = """SELECT  `測試件結果`.`結果編號`,`測試件結果`.`年份`,`測試件結果`.`年度次數`,`測試件項目`.`測試件名稱`, `測試件分項目`.`測試項目_分項`, `測試件結果`.`測試件序號`
                FROM `測試件結果`
                JOIN `測試件項目`
                ON `測試件結果`.`測試件項目編號` = `測試件項目`.`編號` 
                JOIN `測試件分項目`
                ON`測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號` 
                ORDER BY `測試件結果`.`結果編號` ASC;"""
    cursor.execute(test_db)
testname = cursor.fetchall()
df = pd.DataFrame(testname,columns=["結果編號","年份","年度次數","測試件名稱","測試項目_分項","測試件序號"])
resultnum = []
class loadbatch_mySQL(object):
    def __init__(self):  
        #建立資料輸入介面  
        self.root = ctk.CTk()
        # ctk.set_default_color_theme("green")  
        # 給主視窗設定標題內容  
        self.root.title("能力試驗結果批次輸入")  
        self.root.geometry('800x1000')
        style = Style(theme='cyborg') 
        ttk.Style().configure("Treeview", fieldbackground = "#323232")
        self.root.config(background='#323232') #設定背景色
        self.root.bind('<Return>', self.return_click)
        ###建立框架###
        self.frame1 = tk.LabelFrame(self.root, text="Excel Data", foreground="#323232")
        self.frame1.place(height=400, width=750)
        
        ###建立圖片###
        self.canvas = tk.Canvas(self.root, height=97, width=689,background="#323232",highlightthickness=0)#建立畫布
        self.image_file = tk.PhotoImage(file='batchinput_result.png')#載入圖片檔案  
        self.image = self.canvas.create_image(0,0, anchor='nw', image=self.image_file)#將圖片置於畫布上  
        self.canvas.pack(side='top')#放置畫布（為上端）  
        
        ###建立treeview(可視覺化excel)###
        self.tv1 = ttk.Treeview(self.frame1)
        self.tv1.place(relheight=1, relwidth=1)
        treescrolly = ttk.Scrollbar(self.frame1,orient="vertical", command=self.tv1.yview) # command means update the yaxis view of the widget
        treescrollx = ttk.Scrollbar(self.frame1,orient="horizontal", command=self.tv1.xview) # command means update the xaxis view of the widget
        self.tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # assign the scrollbars to the Treeview Widget
        treescrollx.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
        treescrolly.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget
        
        ###建立標籤及按鈕###
        self.hellow_label = ctk.CTkLabel(
            self.root, 
            text = "能力試驗結果批次輸入:",
            fg_color='#323232',
            text_font=('微軟正黑體',20),
            text_color="#00F5FF",
            width=350
            )
        self.name_label = ctk.CTkLabel(
            self.root, 
            text = "支援格式: .csv  .xls  .xlsx",
            fg_color='#323232',
            text_font=('微軟正黑體',14),
            text_color="#00F5FF",
            width=500
            )
        self.msg_label = ctk.CTkLabel(
            self.root, 
            text = """請依下列順序放入Excel中:

年份、年度次數、測試件項目、測試件分項目、測試件序號、能力試驗結果、能力試驗標準差(選填)""",
            fg_color='#323232',
            text_font=('微軟正黑體',13),
            text_color="#00F5FF",
            width=1000,
            height=80
            )
        self.selectfile_btn = ctk.CTkButton(
            self.root, 
            command = self.selectfile, 
            text = "選擇檔案", 
            fg_color='#666666',
            width=240,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.selectfile_btn.pack()
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
        self.button_verify=ctk.CTkButton(
            self.root, 
            command = self.verify_interface, 
            text = "資料驗證", 
            fg_color='#666666',
            width=180,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.button_verify.pack()
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
        self.cc = ctk.CTkLabel(
            self.root, 
            fg_color="#323232",
            text='@Design by Henry Tsai',
            text_color="#8E8E8E",
            text_font="Calibri",
            width=170)
    def gui_arrang(self):
        self.hellow_label.place(relx=0.01, rely=0.05, anchor=tk.W)
        self.frame1.place(relx=0.5,rely=0.3,anchor=tk.CENTER)
        self.msg_label.place(relx=0.5,rely=0.55,anchor=tk.CENTER)
        self.name_label.place(relx=0.5, rely=0.81, anchor=tk.CENTER)
        self.selectfile_btn.place(relx=0.5,rely=0.85,anchor=tk.CENTER)
        self.canvas.place(relx=0.5,rely=0.7,anchor=tk.CENTER)
        self.button_back.place(relx=0.2,rely=0.95,anchor=tk.CENTER)
        self.button_verify.place(relx=0.5,rely=0.95,anchor=tk.CENTER)
        self.button_OK.place(relx=0.8,rely=0.95,anchor=tk.CENTER)
        self.cc.place(relx=1, rely=1,anchor=tk.SE) 
    def selectfile(self):
        self.filename  = filedialog.askopenfilename(initialdir="E:\python\Tkinter",title="能力試驗資料批次輸入",filetypes=(("Excel","*.xlsx"),("CSV UTF-8","*.csv"),("Excel 2003","*.xls"),("all files","*.*")))
        self.selectfile_path = ctk.CTkLabel(
            self.root, 
            text = self.filename,
            fg_color='#323232',
            text_font=('微軟正黑體',12),
            text_color="#FFFFFF",
            width=800
            )
        self.selectfile_path.pack()
        self.selectfile_path.place(relx=0.5,rely=0.79,anchor=tk.CENTER)
        tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='請先驗證資料後，再按確定上傳!')
    def verify_interface(self):
        file_path = self.filename
        global df_raw,upload_df
        try:
            excel_filename = r"{}".format(file_path)
            if excel_filename[-4:] == ".csv":
                df_raw = pd.read_csv(excel_filename,skiprows=2)
            else:
                df_raw = pd.read_excel(excel_filename,skiprows=2)
        except ValueError:
            tk.messagebox.showerror('南投署立醫院檢驗科', message='資料格式不符，請重新選擇檔案')
            return None
        except FileNotFoundError:
            tk.messagebox.showerror("Information", message='未選擇檔案，請選擇檔案後再重新驗證!')
            return None
        titlecompair = df_raw.columns
        title = ['年份', '年度次數', '測試件項目', '測試件分項目', '測試件序號','能力試驗結果','能力試驗標準差']
        for j in range(0,len(titlecompair)):
            if titlecompair[j] == title[j]:
                continue
            else:
                tk.messagebox.showerror(title='南投署立醫院檢驗科', message="檔案錯誤!請檢查檔案後重新輸入!!")
                return None
        self.clear_data()
        self.tv1["column"] = list(df_raw.columns)
        self.tv1["show"] = "headings"
        for column in self.tv1["columns"]:
            self.tv1.heading(column, text=column) # let the column heading = column name
            self.tv1.column(column, width= 80, anchor=CENTER)
        df_rows = df_raw.to_numpy().tolist() # turns the dataframe into a list of lists
        for row in df_rows:
            self.tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
        for i in range(0,len(df_raw)):
            filt = (df["年份"]== str(df_raw.at[i,"年份"])) & (df["年度次數"]== int(df_raw.at[i,"年度次數"])) & (df["測試件名稱"]== str(df_raw.at[i,"測試件項目"])) & (df["測試項目_分項"] == str(df_raw.at[i,"測試件分項目"])) & (df["測試件序號"] == int(df_raw.at[i,"測試件序號"]))
            aa = df.loc[filt,["結果編號"]]
            # print(aa)
            if aa.empty:
                tk.messagebox.showerror(title='南投署立醫院檢驗科', message="能力試驗資料尚未上傳!!")
                return None
            else:
                aa.index = pd.Series([0])
                bb = aa.at[0,"結果編號"]
                # print(bb)
                resultnum.append(bb)
        upload_df = df_raw.drop(["年份","年度次數","測試件項目","測試件分項目","測試件序號"],axis=1)
        upload_df.insert(0,column="結果編號", value = resultnum)
        uploadtestname = df_raw.at[0,"測試件項目"]
        tk.messagebox.showinfo(title='南投署立醫院檢驗科', message="""驗證成功!!
上傳能力試驗項目: %s
上傳筆數: %d 筆
請按確定繼續上傳!"""%(uploadtestname,len(df_raw)))
        return None
        # print("Verify")
    def OK_interface(self):
        # try:  upload_df已經沒有"測試件項目"
        #     upload_df = df.replace({"測試件項目": testname})
        # except NameError:
        #     tk.messagebox.showerror('南投署立醫院檢驗科', "檔案尚未驗證!請先驗證後再按確定!")
        #     return None
        rawdata = upload_df.to_numpy().tolist()
        for i in range(0, len(upload_df)):
            with conn.cursor() as cursor:
                val = "INSERT INTO `能力試驗結果`(`測試件結果編號`,`能力試驗數值`,`能力試驗標準差`) VALUES ('%s', %.5f, %.5f);" %(rawdata[i][0],rawdata[i][1],rawdata[i][2])
                cursor.execute(val)
            conn.commit()
        conn.close()
        tk.messagebox.showinfo('南投署立醫院檢驗科', "上傳成功!總共新增%d筆資料!"%(len(rawdata)))
    def clear_data(self):
        self.tv1.delete(*self.tv1.get_children())
        return None
    def back_interface(self):
        self.root.destroy()
    def return_click(self, event):  #按Enter鍵自動連結登入
        self.verify_interface()
    
def main():  
    L = loadbatch_mySQL()
    L.gui_arrang()
    # 主程式執行  
    tk.mainloop()

if __name__ == '__main__':  
    main()  