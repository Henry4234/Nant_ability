#authorised by Henry Tsai
import sys
from threading import activeCount
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from tkinter import messagebox
import customtkinter as ctk
import subprocess

from setuptools import Command
from verifyAccount import changepw
global account
account = sys.argv[1]
class basedesk:
    def __init__(self):  
        # 建立登入後視窗  
        self.root = ctk.CTk()
        # ctk.set_default_color_theme("green")  
        # 給主視窗設定標題內容  
        self.root.title("能力試驗")  
        self.root.geometry('800x600')
        self.root.config(background='#323232') #設定背景色
        global account_1
        s = ttk.Style()
        s.configure('Red.TLabelframe.Label', font=('微軟正黑體', 12))
        s.configure('Red.TLabelframe.Label', foreground ='#FFFFFF')
        s.configure('Red.TLabelframe.Label', background='#323232')
        self.hellow_label = ctk.CTkLabel(
            self.root, 
            # text = "歡迎回來",
            text = "歡迎回來%s"%(account),
            fg_color='#323232',
            text_font=('微軟正黑體',20),
            text_color="#00F5FF",
            width=240
            )
        self.hellow_label.pack()
        self.label_1 = ctk.CTkLabel(
            self.root, 
            text = "請選擇需要使用的功能:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=300
            )
        self.label_1.pack()
        ##框架設置
        self.labelframe_1 = tk.LabelFrame(
            self.root,
            text='1. 操作能力試驗後',
            foreground="#FFFFFF",
            background="#323232")
        self.labelframe_1.pack()
        self.labelframe_2 = tk.LabelFrame(
            self.root,
            text='2. 能力試驗結果回來',
            foreground="#FFFFFF",
            background="#323232")
        self.labelframe_2.pack()
        self.labelframe_3 = tk.LabelFrame(
            self.root,
            text='3. 結果匯出',
            foreground="#FFFFFF",
            background="#323232")
        self.labelframe_3.pack()
        self.labelframe_4 = tk.LabelFrame(
            self.root,
            text='4. 其他',
            foreground="#FFFFFF",
            background="#323232")
        self.labelframe_4.pack()
        ##各類功能選單
        self.button_1=ctk.CTkButton(
            self.labelframe_1, 
            command = self.keyin_interface, 
            text = "能力試驗資料輸入",
            fg_color='#666666', 
            width=160,height=40,
            text_font=('微軟正黑體',12),
            text_color="#00F5FF",
            )
        self.button_1.pack(padx=10, pady=15)
        self.button_backup=ctk.CTkButton(
            self.labelframe_1, 
            command = self.keyin_backup, 
            text = "能力試驗資料輸入(備機)",
            fg_color='#666666', 
            width=200,height=40,
            text_font=('微軟正黑體',12),
            text_color="#00F5FF",
            )
        self.button_backup.pack(padx=10, pady=15)
        self.button_1s=ctk.CTkButton(
            self.labelframe_1, 
            command = self.keyin_batch, 
            text = "能力試驗資料批次輸入",
            fg_color='#666666', 
            width=190,height=40,
            text_font=('微軟正黑體',12),
            text_color="#00F5FF"
            )
        self.button_1s.pack(padx=10, pady=15)
        self.button_backupbatch=ctk.CTkButton(
            self.labelframe_1, 
            command = self.keyinbatch_backup, 
            text = "能力試驗資料批次輸入(備機)",
            fg_color='#666666', 
            width=230,height=40,
            text_font=('微軟正黑體',12),
            text_color="#00F5FF",
            )
        self.button_backupbatch.pack(padx=10, pady=15)
        self.button_2=ctk.CTkButton(
            self.labelframe_2, 
            command = self.keyinresult_interface, 
            text = "能力試驗結果輸入",
            fg_color='#666666', 
            width=160,height=40,
            text_font=('微軟正黑體',12),
            text_color="#00F5FF",
            )
        self.button_2.pack(padx=10, pady=25)
        self.button_2s=ctk.CTkButton(
            self.labelframe_2, 
            command = self.keyinresult_batch, 
            text = "能力試驗結果批次輸入",
            fg_color='#666666', 
            width=190,height=40,
            text_font=('微軟正黑體',12),
            text_color="#00F5FF"
            )
        self.button_2s.pack(padx=10, pady=25)
        self.button_3=ctk.CTkButton(
            self.labelframe_3,
            command = self.output_interface, 
            text = "能力試驗結果匯出", 
            fg_color='#666666',
            width=160,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.button_3.pack(padx=10, pady=25)
        self.button_3s=ctk.CTkButton(
            self.labelframe_3,
            command = self.output_batch, 
            text = "能力試驗結果批次匯出", 
            fg_color='#666666',
            width=190,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.button_3s.pack(padx=10, pady=25)
        self.button_4=ctk.CTkButton(
            self.labelframe_4, 
            command = "", 
            text = "能力試驗檢體存放位置查詢", 
            fg_color='#666666',
            width=220,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.button_4.pack(padx=10, pady=25)
        self.button_5=ctk.CTkButton(
            self.root, 
            command = self.logout_interface, 
            text = "登出", 
            fg_color='#408080',
            width=180,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.button_6=ctk.CTkButton(
            self.root, 
            command = self.exit_interface, 
            text = "結束使用", 
            fg_color='#408080',
            width=180,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.button_6.pack()
        self.button_7=ctk.CTkButton(
            self.labelframe_4, 
            command = self.dashboard, 
            text = "查詢上傳結果", 
            fg_color='#666666',
            width=180,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.button_7.pack()
        self.button_changepw=ctk.CTkButton(
            self.root, 
            command = self.changepw, 
            text = "更改密碼", 
            fg_color='#408080',
            width=180,height=40,
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.cc = ctk.CTkLabel(
            self.root, 
            fg_color="#323232",
            text='@Design by Henry Tsai',
            text_color="#8E8E8E",
            text_font="Calibri",
            width=170)
    def gui_arrang(self):
        self.hellow_label.place(relx=0.5, rely=0.1, anchor=tk.CENTER)
        self.label_1.place(relx=0.5, rely=0.2, anchor=tk.CENTER)
        # self.button_1.place(relx=0.2,rely=0.4,anchor=tk.CENTER)
        # self.button_1s.place(relx=0.2,rely=0.55,anchor=tk.CENTER)
        # self.button_2.place(relx=0.5,rely=0.4,anchor=tk.CENTER)
        # self.button_2s.place(relx=0.5,rely=0.55,anchor=tk.CENTER)
        # self.button_3.place(relx=0.8,rely=0.4,anchor=tk.CENTER)
        self.button_5.place(relx=0.2,rely=0.9,anchor=tk.CENTER)
        self.button_6.place(relx=0.5,rely=0.9,anchor=tk.CENTER)
        self.button_changepw.place(relx=0.8,rely=0.9,anchor=tk.CENTER)
        self.cc.place(relx=1, rely=1,anchor=tk.SE) 
        self.labelframe_1.place(relx=0.17,rely=0.48, anchor=tk.CENTER)
        self.labelframe_2.place(relx=0.52,rely=0.4, anchor=tk.CENTER)
        self.labelframe_3.place(relx=0.85,rely=0.4, anchor=tk.CENTER)
        self.labelframe_4.place(relx=0.83,rely=0.7, anchor=tk.CENTER)
    def changepw(self):
        label_1 = tk.simpledialog.askstring(
            title = '南投署立醫院檢驗科',
            show='*',
            prompt='請輸入新密碼：')
        while label_1 is not None:
            if label_1 != "":
                label_2 = tk.simpledialog.askstring(
                title = '南投署立醫院檢驗科',
                show='*',
                prompt='請再次輸入新密碼：')
                if label_2 == label_1 :
                    verifyResult = changepw(account,label_2)  
                    tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='密碼修改成功!')
                    break
                else:
                    tk.messagebox.showwarning(title='南投署立醫院檢驗科', message='兩次密碼不同!請重新輸入!')
                    label_1 = tk.simpledialog.askstring(
                    title = '南投署立醫院檢驗科',
                    prompt='請輸入新密碼：')
            else:
                tk.messagebox.showwarning(title='南投署立醫院檢驗科', message='密碼不得為空白!')
                label_1 = tk.simpledialog.askstring(
                title = '南投署立醫院檢驗科',
                prompt='請輸入新密碼：')
        
    def keyin_interface(self):
        command = "python Input_mySQL.py " + account
        subprocess.run(command, shell=True)
    def keyin_batch(self):
        command = "python Inputbatch_mySQL.py " + account
        subprocess.run(command, shell=True)
    def keyin_backup(self):
        command = "python Input_backup.py " + account
        subprocess.run(command, shell=True)
    def keyinbatch_backup(self):
        command = "python Inputbatch_backup.py " + account
        subprocess.run(command, shell=True)
    def keyinresult_interface(self):
        command = "python Input_result.py " + account
        subprocess.run(command, shell=True)
    def keyinresult_batch(self):
        command = "python Inputbatch_result.py " + account
        subprocess.run(command, shell=True)
    def output_interface(self):
        subprocess.run("python output_mySQL.py", shell=True)
    def output_batch(self):
        subprocess.run("python outputbatch_mySQL.py", shell=True)
    def dashboard(self):
    	subprocess.run("python dashboard.py", shell=True)
    def logout_interface(self):
        if tk.messagebox.askyesno(title='南投署立醫院檢驗科', message='確定要登出嗎?', ):
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='已登出!')
            self.root.destroy()
            subprocess.run("python main.py", shell=True) 
        else:
            return
    def exit_interface(self):
        if tk.messagebox.askyesno(title='南投署立醫院檢驗科', message='確定要離開嗎?', ):
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='結束能力試驗!')
            self.root.destroy()
        else:
            return

def main():  
    B = basedesk()
    B.gui_arrang()
    # 主程式執行  
    tk.mainloop()  
  
  
if __name__ == '__main__':  
    main()  