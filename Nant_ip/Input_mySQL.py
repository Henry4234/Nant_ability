#authorised by Henry Tsai
import sys
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk, StringVar
from ttkbootstrap import Style
import customtkinter  as ctk
from collections import defaultdict
import subprocess
import pymysql

global account
account = sys.argv[1]
#建立與mySQL連線資料
# db_settings = { 
#     "host": "192.168.0.120",
#     "port": 3307,
#     "user": "root",
#     "db": "nantou db",
#     "charset": "utf8"
#     }
db_settings = { 
    "host": "127.0.0.1",
    "port": 3306,
    "user": "root",
    "password": "ROOT",
    "db": "nantou db",
    "charset": "utf8"
    }
conn = pymysql.connect(**db_settings)

#與mySQL建立連線，取出測試件項目工作表中的測試件名稱以及編號
with conn.cursor() as cursor:
    cursor.execute("SELECT `測試件項目`.測試件名稱, `測試件分項目`.測試項目_分項 FROM `測試件項目`, `測試件分項目` WHERE (`測試件項目`.編號 = `測試件分項目`.編號);")
result_subtestname = cursor.fetchall()
result_subtestname=tuple((y, x) for x, y in result_subtestname)
subtestname = defaultdict(list)
for i,j in result_subtestname:
    subtestname[j].append(i)
sorted(subtestname.items())
result=list(subtestname.keys())
# conn.close()

class load_mySQL(object):
    def __init__(self):  
        #建立資料輸入介面  
        self.root = ctk.CTk()
        # ctk.set_default_color_theme("green")  
        # 給主視窗設定標題內容  
        self.root.title("能力試驗資料輸入")  
        self.root.geometry('800x600')
        style = Style(theme='cyborg') 
        self.root.config(background='#323232') #設定背景色
        self.root.bind('<Return>', self.return_click)

        self.hellow_label = ctk.CTkLabel(
            self.root, 
            text = "能力試驗資料輸入",
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
            width=120
            )
        self.label_testname = ctk.CTkLabel(
            self.root, 
            text = "測試件名稱:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=150
            )
        self.label_testnum_1 = ctk.CTkLabel(
            self.root, 
            text = "年度第幾次:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=150
            )
        self.label_testnum_2 = ctk.CTkLabel(
            self.root, 
            text = "測試件序號:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=150
            )
        self.label_testobj = ctk.CTkLabel(
            self.root, 
            text = "能力試驗項目:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=150
            )
        self.label_testval = ctk.CTkLabel(
            self.root, 
            text = "試驗結果:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=150
            )
        self.input_year = ctk.CTkEntry(
            self.root, 
            fg_color='#666666',
            text_color="#FFFFFF",
            width=150,height=35
            )
        self.input_year.pack()
        self.input_testname = ttk.Combobox(
            master = self.root, 
            values=result,
            width=15,height=40,
            state="readonly"
            )
        self.input_testname.bind("<<ComboboxSelected>>", self.callback)
        self.input_testnum_1 = ctk.CTkEntry(
            self.root, 
            fg_color='#666666',
            text_color="#FFFFFF",
            width=150,height=35
            )
        self.input_testnum_2 = ctk.CTkEntry(
            self.root, 
            fg_color='#666666',
            text_color="#FFFFFF",
            width=150,height=35
            )
        self.input_testobj = ttk.Combobox(
            master = self.root, 
            # textvariable=self.input_testname.get(),
            values=subtestname[self.input_testname.get()],
            width=15,height=40,
            state="readonly"
            )
        self.input_testval = ctk.CTkEntry(
            self.root, 
            fg_color='#666666',
            text_color="#FFFFFF",
            width=150,height=35
            )
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

    def callback(self,event):  #combobox雙層列表
        self.testobj = StringVar()
        self.testobj=subtestname[self.input_testname.get()]
        self.input_testobj.configure(values = subtestname[self.input_testname.get()])
    
    def back_interface(self):
        self.root.destroy()
    
    def OK_interface(self):
        input_year = self.input_year.get()
        input_testname = self.input_testname.get()  #需要比對'測試件項目'
        input_testnum_1 = self.input_testnum_1.get()
        input_testnum_2 = self.input_testnum_2.get()
        input_testobj = self.input_testobj.get()    #需要比對'測試件分項目'
        input_testval = self.input_testval.get()
        if input_year == "" or input_testname == "" or input_testnum_1 =="" or input_testnum_2 =="" or input_testobj =="" or input_testval =="":
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='輸入不完全!請重新輸入!')
        else:
            input_testnum_1 = int(input_testnum_1)
            input_testnum_2 = int(input_testnum_2)
            input_testval = float(input_testval)
            with conn.cursor() as cursor:
                cursor.execute("SELECT `測試件項目`.`編號`, `測試件項目`.`測試件名稱` FROM `測試件項目` WHERE `測試件項目`.`測試件名稱`= %s;",input_testname)
            name = cursor.fetchone()
            testname_num = int(name[0])
            with conn.cursor() as cursor:
                cursor.execute("SELECT `測試件分項目`.`分項編號`, `測試件分項目`.`測試項目_分項` FROM `測試件分項目` WHERE `測試件分項目`.`測試項目_分項`= %s;",input_testobj)
            name = cursor.fetchone()
            testobj_num = int(name[0])
            with conn.cursor() as cursor:
                val = "INSERT INTO `測試件結果`(`年度次數`, `測試件項目編號`, `測試件分項目編號`, `年份`, `測試件序號`, `測試件結果`,`新增人員`) VALUES (%d, %d, %d, %s, %d, %.5f,'%s');" %(input_testnum_1,testname_num,testobj_num,input_year,input_testnum_2,input_testval,account)
                cursor.execute(val)
            conn.commit()
            conn.close()
            # print(input_year,input_testname,input_testnum,input_testobj,input_testval)
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='新增成功!')
            if tk.messagebox.askyesno(title='南投署立醫院檢驗科', message='要繼續輸入數值?', ):
                self.clear_data()
            else:
                self.root.destroy()
                
    def clear_data(self):
        # self.input_year.delete(0,"end")
        # self.input_testname.set("")
        # self.input_testnum_1.delete(0,"end")
        # self.input_testnum_2.delete(0,"end")
        # self.input_testobj.set("")
        self.input_testval.delete(0,"end")

    def gui_arrang(self):
        self.hellow_label.place(relx=0, rely=0.1, anchor=tk.W)
        self.label_year.place(relx=0.31, rely=0.2, anchor=tk.W)
        self.label_testname.place(relx=0.25,rely=0.3,anchor=tk.W)
        self.label_testnum_1.place(relx=0.25,rely=0.4,anchor=tk.W)
        self.label_testnum_2.place(relx=0.25,rely=0.5,anchor=tk.W)
        self.label_testobj.place(relx=0.23,rely=0.6,anchor=tk.W)
        self.label_testval.place(relx=0.26,rely=0.7,anchor=tk.W)
        self.input_year.place(relx=0.43, rely=0.2, anchor=tk.W)
        self.input_testname.place(relx=0.45,rely=0.3,anchor=tk.W)
        self.input_testnum_1.place(relx=0.43, rely=0.4, anchor=tk.W)
        self.input_testnum_2.place(relx=0.43, rely=0.5, anchor=tk.W)
        self.input_testval.place(relx=0.43,rely=0.7,anchor=tk.W)
        self.input_testobj.place(relx=0.45,rely=0.6,anchor=tk.W)
        self.button_back.place(relx=0.35,rely=0.8,anchor=tk.CENTER)
        self.button_OK.place(relx=0.65,rely=0.8,anchor=tk.CENTER)
        self.cc.place(relx=1, rely=1,anchor=tk.SE) 
    def return_click(self, event):  #按Enter鍵自動連結登入
        self.OK_interface()
def main():  
    L = load_mySQL()
    L.gui_arrang()
    # 主程式執行  
    tk.mainloop()

if __name__ == '__main__':  
    main()  