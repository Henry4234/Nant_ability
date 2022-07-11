import tkinter as tk
import customtkinter as ctk
class output_mySQL(object):
    val_2 = 0
    def __init__(self):
        self.root = ctk.CTk()
        self.input_testnum = ctk.CTkEntry(
            self.root, 
            fg_color='#666666',
            text_color="#FFFFFF",
            )
        self.button_back=ctk.CTkButton(
            self.root, 
            command = self.getvalue, 
            text = "確定", 
            fg_color='#666666',
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.button_back_2=ctk.CTkButton(
            self.root, 
            command = lambda: self.getvalue(val = 123), 
            text = "確定_2", 
            fg_color='#666666',
            text_font='微軟正黑體',
            text_color="#00F5FF"
            )
        self.input_testnum.place(relx=0.5, rely=0,anchor = tk.N)
        self.button_back.place(relx=0.5, rely=0.2,anchor = tk.N)
        self.button_back_2.place(relx=0.5, rely=0.4,anchor = tk.N)
    def getvalue(self,*arg):
        val = self.input_testnum.get()
        print(val)
        print(arg)
    # def getval_2(self):
    #     output_mySQL.getvalue(self, val_2 = 1234)
def main():  
    L = output_mySQL()
    # 主程式執行  
    tk.mainloop()
if __name__ == "__main__":
    main()