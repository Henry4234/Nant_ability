#authorised by Henry Tsai
import pickle
import subprocess
import tkinter as tk
from tkinter import messagebox
from turtle import goto
import customtkinter  as ctk
import verifyAccount


class Login:    #建立登入介面  
    #初始化設定__init__
    def __init__(self):
        # 建立主視窗,用於容納其它元件  
        self.root = ctk.CTk()
        ctk.set_default_color_theme("green")  
        # 給主視窗設定標題內容  
        self.root.title("南投署立醫院檢驗科")  
        self.root.geometry('600x300')
        self.root.config(background='#323232')
        # self.account_2 = None
        self.root.bind('<Return>', self.callback)
        #建立圖片
        self.canvas = tk.Canvas(self.root, height=125, width=500,background="#323232",highlightthickness=0)#建立畫布
        # self.canvas.comfig(highlightthickness=0)  
        self.image_file = tk.PhotoImage(file='logo.png')#載入圖片檔案  
        self.image = self.canvas.create_image(0,0, anchor='nw', image=self.image_file)#將圖片置於畫布上  
        self.canvas.pack(side='top')#放置畫布（為上端）  
        #建立一個`label`名為`Account: `  
        self.label_account = ctk.CTkLabel(self.root, text='Account: ')  
        #建立一個`label`名為`Password: `  
        self.label_password = ctk.CTkLabel(self.root, text='Password: ')       
        # 建立一個賬號輸入框,並設定尺寸  
        self.input_account = ctk.CTkEntry(self.root, width=120)
        self.cc = ctk.CTkLabel(
            self.root, 
            fg_color="#323232",
            text='@Design by Henry Tsai',
            text_color="#8E8E8E",
            text_font="Calibri",
            width=170)  
        # 建立一個密碼輸入框,並設定尺寸  
        self.input_password = ctk.CTkEntry(self.root, show='*',  width=120)  
        # 建立一個登入系統的按鈕  
        self.login_button = ctk.CTkButton(self.root, command = self.backstage_interface, text = "登入", width=60,text_font='微軟正黑體')
        self.login_button.pack()
        # 建立一個退出系統的按鈕  
        self.exit_button = ctk.CTkButton(self.root, command = self.exit_interface, text = "退出", width=60,text_font='微軟正黑體')
        self.exit_button.pack()
        # 完成佈局
    def gui_arrang(self):  
        self.label_account.place(relx=0.4, rely=0.5, anchor=tk.CENTER)  
        self.label_password.place(relx=0.4, rely=0.7, anchor=tk.CENTER)  
        self.input_account.place(relx=0.6, rely=0.5, anchor=tk.CENTER)  
        self.input_password.place(relx=0.6, rely=0.7, anchor=tk.CENTER)  
        self.login_button.place(relx=0.4, rely=0.9, anchor=tk.CENTER)  
        self.exit_button.place(relx=0.6, rely=0.9, anchor=tk.CENTER)
        self.cc.place(relx=1, rely=1,anchor=tk.SE)  
    # 退出介面  
    def exit_interface(self):  
        self.root.destroy()  
    # 進行登入資訊驗證  
    def backstage_interface(self):  
        # with open('pw.pickle','wb') as usr_file:
        #     usrs_info={'admin':'admin','t001':'t001'}
        #     pickle.dump(usrs_info,usr_file)
        # # global account
        global account
        # global account_1
        account = self.input_account.get()
        password = self.input_password.get()
        # idreturn(account)
        #對賬戶資訊進行驗證，普通使用者返回user，管理員返回master，賬戶錯誤返回noAccount，密碼錯誤返回noPassword  
        verifyResult = verifyAccount.verifyAccountData(account,password)  
        if verifyResult=='master':  
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='進入管理介面')
            self.loginuseradmin()
        elif verifyResult=='user':  
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='進入使用者介面')
            self.loginuser()   
        elif verifyResult=='noAccount':  
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='該賬號不存在請重新輸入!')  
        elif verifyResult=='noPassword':  
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='賬號/密碼錯誤請重新輸入!')
        elif verifyResult=='empty':
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='未輸入賬號/密碼!')
    def callback(self, event):  #按Enter鍵自動連結登入
        self.backstage_interface()
    def loginuser(self):
        self.root.destroy()
        command = "python basedesk.py " + account
        subprocess.run(command, shell=True)
    def loginuseradmin(self):
        self.root.destroy()
        command = "python basedesk_admin.py " + account
        subprocess.run(command, shell=True)
def main():  
    # 初始化物件  
    L = Login()  
    # 進行佈局 
    L.gui_arrang()  
    # 主程式執行
    tk.mainloop()  
  
if __name__ == '__main__':  
    main()