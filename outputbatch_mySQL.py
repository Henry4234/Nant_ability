#authorised by Henry Tsai
import os
from logging import root
from pdb import line_prefix
from sysconfig import get_paths
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk, StringVar
import pandas as pd
from pandas import value_counts
from ttkbootstrap import Style
import customtkinter  as ctk
from collections import defaultdict
import subprocess
import pymysql
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart,Reference
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph,CharacterProperties,ParagraphProperties,Font
#建立與mySQL連線資料
db_settings = { 
    "host": "192.168.0.120",
    "port": 3307,
    "user": "root",
    "db": "nantou db",
    "charset": "utf8"
    }
# db_settings = { 
#     "host": "192.168.53.167",
#     "port": 3306,
#     "user": "root",
#     "password": "ROOT",
#     "db": "nantou db",
#     "charset": "utf8"
#     }
conn = pymysql.connect(**db_settings)

#與mySQL建立連線，取出測試件項目工作表中的測試件名稱以及編號
with conn.cursor() as cursor:
    cursor.execute("SELECT `測試件項目`.測試件名稱, `測試件分項目`.測試項目_分項 FROM `測試件項目`, `測試件分項目` WHERE (`測試件項目`.編號 = `測試件分項目`.編號);")
result_subtestname = cursor.fetchall()
with conn.cursor() as cursor:
    cursor.execute("SELECT `測試件名稱`, `編號` FROM `測試件項目`;")
testname = cursor.fetchall()
with conn.cursor() as cursor:
    cursor.execute("SELECT `測試項目_分項`, `分項編號` FROM `測試件分項目`;")
testobj = cursor.fetchall()
result_subtestname=tuple((y, x) for x, y in result_subtestname)
testname_dict = dict((x,y)for x,y in testname)
# result_testobj = tuple((y,x)for x,y in testobj)
# print(testname_dict)
# print(result_testobj)
subtestname = defaultdict(list)
for i,j in result_subtestname:
    subtestname[j].append(i)
sorted(subtestname.items())
result=list(subtestname.keys())
# conn.close()

class output_mySQL(object):

    def __init__(self):  
        #建立資料輸入介面  
        self.root = ctk.CTk()
        # ctk.set_default_color_theme("green")  
        # 給主視窗設定標題內容  
        self.root.title("能力試驗結果匯出")  
        self.root.geometry('800x600')
        style = Style(theme='cyborg') 
        self.root.config(background='#323232') #設定背景色
        global testobj_values
        self.root.bind('<Return>', self.return_click)
        self.hellow_label = ctk.CTkLabel(
            self.root, 
            text = "能力試驗結果匯出",
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
        self.label_testnum = ctk.CTkLabel(
            self.root, 
            text = "年度第幾次:", 
            fg_color='#323232',
            text_font=('微軟正黑體',18),
            text_color="#00F5FF",
            width=150
            )
        # self.label_testobj = ctk.CTkLabel(
        #     self.root, 
        #     text = "能力試驗項目:", 
        #     fg_color='#323232',
        #     text_font=('微軟正黑體',18),
        #     text_color="#00F5FF",
        #     width=150
        #     )
        self.input_year = ctk.CTkEntry(
            self.root, 
            fg_color='#666666',
            text_color="#FFFFFF",
            width=150,height=35
            )
        self.input_testnum = ctk.CTkEntry(
            self.root, 
            fg_color='#666666',
            text_color="#FFFFFF",
            width=150,height=35
            )
        self.input_testname = ttk.Combobox(
            master = self.root, 
            values=result,
            width=15,height=40,
            state="readonly"
            )
        # self.input_testname.bind("<<ComboboxSelected>>", self.callback)
        # self.input_testobj = ttk.Combobox(
        #     master = self.root, 
        #     # textvariable=self.input_testname.get(),
        #     values=subtestname[self.input_testname.get()],
        #     width=15,height=40,
        #     state="readonly"
        #     )
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
        self.button_back.pack()
        self.cc = ctk.CTkLabel(
            self.root, 
            fg_color="#323232",
            text='@Design by Henry Tsai',
            text_color="#8E8E8E",
            text_font="Calibri",
            width=170)

    # def callback(self,event):  #combobox雙層列表
    #     self.testobj = StringVar()
    #     self.testobj=subtestname[self.input_testname.get()]
    #     self.input_testobj.configure(values = subtestname[self.input_testname.get()])
    
    def back_interface(self):
        self.root.destroy()
    
    def OK_interface(self):
        input_year = self.input_year.get()
        input_testnum = self.input_testnum.get()
        # input_testobj = self.input_testobj.get()
        if input_year=="" or input_testnum=="":
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='輸入不完全!請重新輸入!')
        else:
            input_testname = testname_dict[self.input_testname.get()]
            input_testname = int(input_testname)
            input_testnum = int(input_testnum)
            # input_testobj = str(input_testobj)
            with conn.cursor() as cursor:
                    srch_objnum = "SELECT `測試件數` FROM `測試件項目` WHERE `測試件項目`.`編號` = %d;"%(input_testname)
                    cursor.execute(srch_objnum)
            objnum = cursor.fetchone()
            objnum = objnum[0]
            with conn.cursor() as cursor:
                srch_db="""SELECT `測試件結果`.`年份`,  `測試件結果`.`年度次數`, `測試件結果`.`測試件項目編號`, `測試件分項目`.`測試項目_分項`, `測試件結果`.`測試件序號`,`測試件結果`.`測試件結果`, `能力試驗結果`.`能力試驗數值` 
                        FROM `測試件結果`
                        JOIN `測試件項目`
                        ON `測試件結果`.`測試件項目編號` = `測試件項目`.`編號` 
                        JOIN `測試件分項目`
                        ON`測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號`
                        JOIN `能力試驗結果`
                        ON `測試件結果`.`結果編號` = `能力試驗結果`.`測試件結果編號`
                        WHERE `測試件結果`.`測試件分項目編號` IN  (SELECT `分項編號` FROM `測試件分項目` WHERE `編號` = %d)
                        AND `測試件結果`.`年份` = %s 
                        AND `測試件結果`.`年度次數` = %d ;
                        """%(input_testname,input_year,input_testnum)
                cursor.execute(srch_db)
            name = cursor.fetchall()
            cnt = cursor.rowcount
            if cnt == 0:
                tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='查無此筆資料!請確定年份或次數是否輸入正確!')
                self.clear_data()
                return
            name = pd.DataFrame(name)
            testobj_1 = name[3].unique()
            testobj_1.tolist()
            wb = Workbook()
            m = 0
            for q in testobj_1:
                with conn.cursor() as cursor:
                    srch_clsi = """SELECT `能力試驗結果`.`測試件結果編號`, `測試件分項目`.`測試項目_分項`,`clsi規則`.`規則內容`, `clsi規則`.`實際數值`,`clsi規則`.`實際數值_1`
                                    FROM (`測試件結果` INNER JOIN `測試件分項目` ON `測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號`)
                                    JOIN `能力試驗結果`
                                    ON `測試件結果`.`結果編號` = `能力試驗結果`.`測試件結果編號`
                                    JOIN `clsi規則`
                                    ON `能力試驗結果`.`CLSI規則` = `clsi規則`.`編號`
                                    WHERE `能力試驗結果`.`測試件結果編號` IN  (
                                        SELECT `結果編號` FROM `測試件結果` WHERE`測試件結果`.`測試件分項目編號`IN(
                                            SELECT `分項編號` FROM `測試件分項目` WHERE `編號` = %d AND `測試項目_分項` = '%s'));
                                """%(input_testname,q)
                    cursor.execute(srch_clsi)
                rules = cursor.fetchone()
                aa = rules[2]
                # print(rules[3])
                if m == 0:
                    ws = wb.active
                else:
                    ws = wb.create_sheet()
                ws.title = "%s_%d_%s_%s"%(input_year,input_testnum,self.input_testname.get(),q)
                # ws.title()
                title=["年份","年度次數","測試件項目","測試件分項目","測試件序號","測試件結果","目標值"]
                ws.append(title)
                #放入目標數值
                for j in range(m, objnum+m):
                    ws.append(name.iloc[j].tolist())
                    m += 1
                #刪除不必要資訊
                ws.delete_cols(1,amount = 4)

                ##計算與peer差異值
                ws['D1'] = "差異值"
                for x in range(2,objnum+2):
                    if ws['B' + str(x)].value == None:
                        tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='測試件無結果!，無法自動計算可容許範圍')    
                    else:
                        actualmean = ws['B' + str(x)].value
                        peermean =  ws['C' + str(x)].value
                        ws['D' + str(x)].value = actualmean - peermean
                ##新增標題
                ws.insert_rows(1)
                ws['A1'] = "%s年第%d次%s_%s"%(input_year,input_testnum,self.input_testname.get(),q)
                ws.merge_cells('A1:B1')
                ws.merge_cells('A23:B23')
                ws['E2'] = "可容許差異高值"
                ws['F2'] = "可容許差異低值"
                ws['G2'] = "差異百分比"
                ##計算可容許差異高/低值
                if "or" in aa:  #如果多於兩個變數
                    if "%" and "SD" in aa:  #如果同時出現SD跟%
                        percentage =[]
                        with conn.cursor() as cursor:   #取得SD值
                            srch_sd = """SELECT `能力試驗結果`.`測試件結果編號`, `能力試驗結果`.`能力試驗標準差`
                                        FROM (`測試件結果` INNER JOIN `測試件分項目` ON `測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號`)
                                        JOIN `能力試驗結果`
                                        ON `測試件結果`.`結果編號` = `能力試驗結果`.`測試件結果編號`
                                        WHERE `能力試驗結果`.`測試件結果編號` IN  (
                                            SELECT `結果編號` FROM `測試件結果` WHERE`測試件結果`.`測試件分項目編號`IN(
                                                SELECT `分項編號` FROM `測試件分項目` WHERE `編號` = %d AND `測試項目_分項` = '%s')
                                                AND `測試件結果`.`年份` = '%s' 
                                                AND `測試件結果`.`年度次數` = %d);
                            """%(input_testname,q,input_year,input_testnum)
                            cursor.execute(srch_sd)
                        sd = cursor.fetchall()
                        for i in range(3, objnum + 3):
                            if ws['B' + str(i)].value == None:
                                percentage.append(None)
                                continue
                            else:
                                realamount_1 = rules[3]
                                realamount_2 = rules[4]
                                realamount_1 = ws['C' + str(i)].value * realamount_1
                                realamount_2 = sd[i-3][1] * realamount_1
                                if realamount_1 > realamount_2:
                                    ws['E' + str(i)].value = realamount_1
                                    ws['F' + str(i)].value = (-realamount_1)
                                    ws['G' + str(i)].value = ws['D' + str(i)].value / ws['E' + str(i)].value
                                    ws['G' + str(i)].number_format = "0.00%"
                                else: 
                                    ws['E' + str(i)].value = realamount_2
                                    ws['F' + str(i)].value = (-realamount_2)
                                    ws['G' + str(i)].value = ws['D' + str(i)].value / ws['E' + str(i)].value
                                    ws['G' + str(i)].number_format = "0.00%"
                                percent = ws['G' + str(i)].value * 100
                                percentage.append(percent)
                    elif "%" in aa:
                        percentage =[]
                        for z in range(3, objnum + 3):
                            if ws['B' + str(i)].value == None:
                                percentage.append(None)
                                continue
                            else:
                                realamount_1 = rules[3]
                                realamount_2 = rules[4]
                                realamount_1 = ws['C' + str(z)].value * realamount_1
                                if realamount_1 > realamount_2:
                                    ws['E' + str(z)].value = realamount_1
                                    ws['F' + str(z)].value = (-realamount_1)
                                    ws['G' + str(z)].value = ws['D' + str(z)].value / ws['E' + str(z)].value
                                    ws['G' + str(z)].number_format = "0.00%"
                                else: 
                                    ws['E' + str(z)].value = realamount_2
                                    ws['F' + str(z)].value = (-realamount_2)
                                    ws['G' + str(z)].value = ws['D' + str(z)].value / ws['E' + str(z)].value
                                    ws['G' + str(z)].number_format = "0.00%"
                                percent = ws['G' + str(z)].value * 100
                                percentage.append(percent)
                    elif "SD" in aa:
                        percentage =[]
                        with conn.cursor() as cursor:   #取得SD值
                            srch_sd = """SELECT `能力試驗結果`.`測試件結果編號`, `能力試驗結果`.`能力試驗標準差`
                                        FROM (`測試件結果` INNER JOIN `測試件分項目` ON `測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號`)
                                        JOIN `能力試驗結果`
                                        ON `測試件結果`.`結果編號` = `能力試驗結果`.`測試件結果編號`
                                        WHERE `能力試驗結果`.`測試件結果編號` IN  (
                                            SELECT `結果編號` FROM `測試件結果` WHERE`測試件結果`.`測試件分項目編號`IN(
                                                SELECT `分項編號` FROM `測試件分項目` WHERE `編號` = %d AND `測試項目_分項` = '%s')
                                                AND `測試件結果`.`年份` = '%s' 
                                                AND `測試件結果`.`年度次數` = %d);
                            """%(input_testname,q,input_year,input_testnum)
                            cursor.execute(srch_sd)
                        sd = cursor.fetchall()
                        for i in range(3, objnum + 3):
                            if ws['B' + str(i)].value == None:
                                percentage.append(None)
                                continue
                            else:
                                realamount_1 = rules[3]
                                realamount_2 = rules[4]
                                realamount_1 = sd[i-3][1] * realamount_1
                                if realamount_1 > realamount_2:
                                    ws['E' + str(i)].value = realamount_1
                                    ws['F' + str(i)].value = (-realamount_1)
                                    ws['G' + str(i)].value = ws['D' + str(i)].value / ws['E' + str(i)].value
                                    ws['G' + str(i)].number_format = "0.00%"
                                else: 
                                    ws['E' + str(i)].value = realamount_2
                                    ws['F' + str(i)].value = (-realamount_2)
                                    ws['G' + str(i)].value = ws['D' + str(i)].value / ws['E' + str(i)].value
                                    ws['G' + str(i)].number_format = "0.00%"
                                percent = ws['G' + str(i)].value * 100
                                percentage.append(percent)
                elif "SD" in aa:   #利用標準差計算高低值
                    percentage =[]
                    realamount = rules[3]
                    with conn.cursor() as cursor:
                        srch_sd = """SELECT `能力試驗結果`.`測試件結果編號`, `能力試驗結果`.`能力試驗標準差`
                                    FROM (`測試件結果` INNER JOIN `測試件分項目` ON `測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號`)
                                    JOIN `能力試驗結果`
                                    ON `測試件結果`.`結果編號` = `能力試驗結果`.`測試件結果編號`
                                    WHERE `能力試驗結果`.`測試件結果編號` IN  (
                                        SELECT `結果編號` FROM `測試件結果` WHERE`測試件結果`.`測試件分項目編號`IN(
                                            SELECT `分項編號` FROM `測試件分項目` WHERE `編號` = %d AND `測試項目_分項` = '%s')
                                            AND `測試件結果`.`年份` = '%s' 
                                            AND `測試件結果`.`年度次數` = %d);
                        """%(input_testname,q,input_year,input_testnum)
                        cursor.execute(srch_sd)
                    sd = cursor.fetchall()
                    # print(sd)
                    for i in range(3,objnum + 3):
                        if ws['B' + str(i)].value == None:
                            percentage.append(None)
                            continue
                        else:
                            ws['E' + str(i)].value = sd[i-3][1] * realamount
                            ws['F' + str(i)].value = sd[i-3][1] * (-realamount)
                            if ws['E' + str(i)].value != 0:
                                ws['G' + str(i)].value = ws['D' + str(i)].value / ws['E' + str(i)].value
                                ws['G' + str(i)].number_format = "0.00%"
                            else:
                                ws['G' + str(i)].value = 0
                            percent = ws['G' + str(i)].value * 100
                            percentage.append(percent)
                elif "%" in aa: #利用百分比計算高低值
                    percentage =[]
                    realamount = rules[3]
                    for i in range(3,objnum + 3):
                        if ws['B' + str(i)].value == None:
                            percentage.append(None)
                            continue
                        else:
                            ws['E' + str(i)].value = ws['C' + str(i)].value * realamount
                            ws['F' + str(i)].value = ws['C' + str(i)].value * (-realamount)
                            ws['G' + str(i)].value = ws['D' + str(i)].value / ws['E' + str(i)].value
                            ws['G' + str(i)].number_format = "0.00%"
                            percent = ws['G' + str(i)].value * 100
                            percentage.append(percent)
                else:   #高低值為固定數值
                    percentage =[]
                    realamount = rules[3]
                    for i in range(3,objnum + 3):
                        if ws['B' + str(i)].value == None:
                            percentage.append(None)
                            continue
                        else:
                            ws['E' + str(i)].value = realamount
                            ws['F' + str(i)].value = (-realamount)
                            ws['G' + str(i)].value = ws['D' + str(i)].value / ws['E' + str(i)].value
                            ws['G' + str(i)].number_format = "0.00%"
                            percent = ws['G' + str(i)].value * 100
                            percentage.append(percent)

                ##標題置中對齊 & 自動適配欄寬
                for col in range(1,7):
                    char = get_column_letter(col)
                    for row in range(2,objnum + 3):
                        ws[char + str(row)].alignment = Alignment(horizontal='center',vertical='center')
                    ws.column_dimensions[get_column_letter(col)].auto_size = True
                ##繪製差異圖表
                chart = LineChart()
                chart.title = ws['A1'].value    #圖表標題
                cp = CharacterProperties(ea= Font(typeface='標楷體'), sz = 1400, b = False) #設定標題字型
                lp = CharacterProperties(ea= Font(typeface='標楷體'), sz = 1000, b = False) #設定圖例字型
                pp = ParagraphProperties(defRPr=lp)
                rlp = RichText(p=[Paragraph(pPr=pp, endParaRPr=cp)])
                chart.title.tx.rich.p[0].pPr.defRPr = cp     #標題設定
                chart.legend.textProperties = rlp       #圖例設定
                # chart.legend.textProperties.rich.p[0].pPr.defRPr = lp    #圖例設定
                chart.legend.position = "b"     #設定圖例放置位置
                ydata = Reference(ws, min_col=4, min_row=2, max_col=6, max_row=objnum+2)
                xvalue = Reference(ws, min_col=1, min_row=3, max_col=1, max_row=objnum+2)
                chart.add_data(ydata, titles_from_data=True)
                chart.set_categories(xvalue)
                s1 = chart.series[1]
                s1.marker.symbol = "circle"
                s1.marker.graphicalProperties.solidFill = "D44E2A" # Marker filling
                s1.marker.graphicalProperties.line.solidFill = "D44E2A" # Marker outline
                s1.graphicalProperties.line.solidFill = "D44E2A"
                s2 = chart.series[0]
                s2.marker.symbol = "circle"
                s2.marker.graphicalProperties.solidFill = "3580EF" # Marker filling
                s2.marker.graphicalProperties.line.solidFill = "3580EF" # Marker outline
                s2.graphicalProperties.line.solidFill = "3580EF"
                s3 = chart.series[2]
                s3.marker.symbol = "circle"
                s3.marker.graphicalProperties.solidFill = "71B62C" # Marker filling
                s3.marker.graphicalProperties.line.solidFill = "71B62C" # Marker outline
                s3.graphicalProperties.line.solidFill = "71B62C"
                ws.add_chart(chart, "A9")
                ##計算差異百分比後自動回傳mySQL
                #先檢查是否為自己判定數值
                if ws['E3'].value =="計算複雜，請自行計算":
                    ws['C1'].value = aa
                    tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='%s可容許範圍判定大於兩個變數，請自行判斷'%(q))
                else:
                    if tk.messagebox.askyesno(title='南投署立醫院檢驗科', message='是否上傳計算後可容許百分比?', ):
                        with conn.cursor() as cursor:
                            srch_num ="""SELECT `測試件結果`.`結果編號` 
                                        FROM `測試件結果`
                                        WHERE `測試件結果`.`測試件分項目編號` IN  (
                                            SELECT `分項編號` 
                                            FROM `測試件分項目` 
                                            WHERE `編號` = %d 
                                            AND `測試項目_分項` = '%s')
                                        AND `測試件結果`.`年份` = %s 
                                        AND `測試件結果`.`年度次數` = %d ;"""%(input_testname,q,input_year,input_testnum)
                            cursor.execute(srch_num)
                        srch_number = [srch_num_1[0] for srch_num_1 in cursor.fetchall()]
                        # print(srch_number)
                        for l in range(0,len(srch_number)):
                            with conn.cursor() as cursor:
                                input_percent="""UPDATE `能力試驗結果` SET`差異百分比`=%.5f WHERE `測試件結果編號`='%s';"""%(percentage[l],srch_number[l])
                                # print(input_percent)
                                cursor.execute(input_percent)
                            conn.commit()
                ###擷取近五次資料
                ##所有能力試驗結果，每年測試次數皆為2或3
                with conn.cursor() as cursor:
                    srch_objnum = "SELECT `測試件數` FROM `測試件項目` WHERE `測試件項目`.`編號` = %d;"%(input_testname)
                    cursor.execute(srch_objnum)
                objnum = cursor.fetchone()
                objnum = objnum[0]
                with conn.cursor() as cursor:   #取得近五年
                    srch_year="""SELECT DISTINCT `測試件結果`.`年份`,`測試件結果`.`年度次數`
                    FROM (`測試件結果` INNER JOIN `測試件分項目` ON `測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號`)
                    JOIN `能力試驗結果`
                    ON `測試件結果`.`結果編號` = `能力試驗結果`.`測試件結果編號`
                    WHERE `能力試驗結果`.`測試件結果編號` IN  (
                        SELECT `結果編號` FROM `測試件結果` WHERE`測試件結果`.`測試件分項目編號`IN(
                            SELECT `分項編號` FROM `測試件分項目` WHERE `編號` = %d AND `測試項目_分項` = '%s'))
                    ORDER BY `測試件結果`.`年份` ASC, `測試件結果`.`年度次數` ASC
                    LIMIT %d;"""%(input_testname,q,objnum*5)
                    cursor.execute(srch_year)
                year = cursor.fetchall()
                title =['測試件名稱']
                for i in range(0,len(year)):
                    title.append('%s年第%d次'%(year[i][0],year[i][1]))
                # print(title)
                with conn.cursor() as cursor:
                    get_past = """SELECT `能力試驗結果`.`測試件結果編號`, `測試件結果`.`年份`,`測試件結果`.`年度次數`,`測試件結果`.`測試件序號`,`能力試驗結果`.`差異百分比`
                                FROM (`測試件結果` INNER JOIN `測試件分項目` ON `測試件結果`.`測試件分項目編號` = `測試件分項目`.`分項編號`)
                                JOIN `能力試驗結果`
                                ON `測試件結果`.`結果編號` = `能力試驗結果`.`測試件結果編號`
                                WHERE `能力試驗結果`.`測試件結果編號` IN  (
                                    SELECT `結果編號` FROM `測試件結果` WHERE`測試件結果`.`測試件分項目編號`IN(
                                        SELECT `分項編號` FROM `測試件分項目` WHERE `編號` = %d AND `測試項目_分項` = '%s'))
                                ORDER BY `測試件結果`.`年份` ASC, `測試件結果`.`年度次數` ASC, `測試件結果`.`測試件序號` ASC
                                LIMIT %d;"""%(input_testname,q,objnum*5)
                    cursor.execute(get_past)
                ans = cursor.fetchall()
                ans = pd.DataFrame(ans)
                ws['A23'] = "能力試驗歷次監控"
                ws.merge_cells('A23:B23')
                ws.append(title)
                for i in range(3,3+objnum):    #新增能力試驗序號
                    ws['A' + str(i+22)].value = ws['A'+str(i)].value  #複製測試件項目
                ws['A'+ str(25 + objnum)] = "平均"
                if len(title) == 6:
                    n = 0   #設定pd變數
                    for col in range(2,7):  #五年固定
                        for row in range(25,objnum+25): #利用測試件數建立迴圈
                            char = get_column_letter(col)
                            ws[char + str(row)].value = ans[4][n]
                            n+=1
                        # ws[get_column_letter(col) + str(objnum+25)].value = "=AVERAGE(B25:B29)"
                        ws[get_column_letter(col) + str(objnum+25)].value = "=ROUND(AVERAGE(%s%d:%s%d)/100,5)"%(get_column_letter(col),25,get_column_letter(col),objnum+24)
                        ws[get_column_letter(col) + str(objnum+25)].number_format = "0.00%"
                else:
                    # tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='歷年能力試驗不足五年!')
                    n = 0   #設定pd變數
                    for col in range(2,len(title) + 1):  #不確定幾年
                        for row in range(25,objnum+25): #利用測試件數建立迴圈
                            char = get_column_letter(col)
                            ws[char + str(row)].value = ans[4][n]
                            n+=1
                        # ws[get_column_letter(col) + str(objnum+25)].value = "=AVERAGE(B25:B29)"
                        ws[get_column_letter(col) + str(objnum+25)].value = "=ROUND(AVERAGE(%s%d:%s%d)/100,5)"%(get_column_letter(col),25,get_column_letter(col),objnum+24)
                        ws[get_column_letter(col) + str(objnum+25)].number_format = "0.00%"
                
                ##繪製近五年圖表
                chart1 = LineChart()
                chart1.title = ws['A1'].value    #圖表標題
                cp = CharacterProperties(ea= Font(typeface='標楷體'), sz = 1400, b = False) #設定標題字型
                lp = CharacterProperties(ea= Font(typeface='標楷體'), sz = 1000, b = False) #設定圖例字型
                pp = ParagraphProperties(defRPr=lp)
                rlp = RichText(p=[Paragraph(pPr=pp, endParaRPr=cp)])
                chart1.title.tx.rich.p[0].pPr.defRPr = cp     #標題設定
                chart1.legend.textProperties = rlp
                chart1.legend.position = "b"     #設定圖例放置位置
                chart1.y_axis.scaling.min = -50      # 設置y軸座標最小的值
                chart1.y_axis.majorUnit = 10       
                chart1.y_axis.scaling.max = 50    

                ydata = Reference(ws, min_col=1, min_row=25, max_col=6, max_row=objnum+25)
                xvalue = Reference(ws, min_col=2, max_col=6, min_row=24, max_row=24)
                chart1.add_data(ydata, from_rows = True, titles_from_data=True)
                chart1.set_categories(xvalue)
                s1 = chart1.series[1]
                s1.marker.symbol = "circle"
                s1.marker.size = 7
                s1.marker.graphicalProperties.solidFill = "D44E2A" # Marker filling
                s1.marker.graphicalProperties.line.solidFill = "D44E2A"
                s1.graphicalProperties.line.noFill = True 
                s2 = chart1.series[0]
                s2.marker.symbol = "triangle"
                s2.marker.size = 7
                s2.marker.graphicalProperties.solidFill = "3580EF" # Marker filling
                s2.marker.graphicalProperties.line.solidFill = "3580EF"
                s2.graphicalProperties.line.noFill = True 
                s3 = chart1.series[2]
                s3.marker.symbol = "square"
                s3.marker.size = 7  
                s3.marker.graphicalProperties.solidFill = "6FB7B7" # Marker filling
                s3.marker.graphicalProperties.line.solidFill = "6FB7B7"
                s3.graphicalProperties.line.noFill = True 
                if objnum == 2:
                    s3.graphicalProperties.line.noFill = False
                    s3.graphicalProperties.line.solidfill = "6FB7B7"
                elif  objnum == 3:
                    s4 = chart1.series[3]
                    s4.marker.symbol = "diamond"
                    s4.marker.size = 7
                    s4.graphicalProperties.line.solidfill = "71B62C"
                elif objnum == 5 :
                    s4 = chart1.series[3]
                    s4.marker.symbol = "diamond"
                    s4.marker.size = 7
                    s4.marker.graphicalProperties.solidFill = "71B62C" # Marker filling
                    s4.marker.graphicalProperties.line.solidFill = "71B62C"
                    s4.graphicalProperties.line.noFill = True 
                    s5 = chart1.series[4]
                    s5.marker.symbol = "triangle"
                    s5.marker.size = 7
                    s5.marker.graphicalProperties.solidFill = "004B97" # Marker filling
                    s5.marker.graphicalProperties.line.solidFill = "004B97"
                    s5.graphicalProperties.line.noFill = True 
                    s6 = chart1.series[5]
                    s6.marker.symbol = "triangle"
                    s6.marker.size = 7
                    s6.graphicalProperties.solidFill = "FF9224" # Marker filling
                    s6.graphicalProperties.line.solidFill = "FF9224"
                ws.add_chart(chart1, "A32")

                wb.save("%s年第%d次%s.xlsx"%(input_year,input_testnum,self.input_testname.get()))   #xlsx檔案存檔
                filepath=".//%s年第%d次%s.xlsx"%(input_year,input_testnum,self.input_testname.get())
        if os.path.isfile(filepath):
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='檔案新增成功!')
            if tk.messagebox.askyesno(title='南投署立醫院檢驗科', message='是否開啟檔案?'):
                command = "start " + filepath
                subprocess.run(command, shell=True)
            else:
                self.clear_data()
                return
                # conn.close()
        else:
            tk.messagebox.showinfo(title='南投署立醫院檢驗科', message='檔案新增失敗QQ')
            # conn.close()
    def clear_data(self):
        self.input_year.delete(0,"end")
        self.input_testname.set("")
        self.input_testnum.delete(0,"end")
    def gui_arrang(self):
        self.hellow_label.place(relx=0, rely=0.1, anchor=tk.W)
        self.label_year.place(relx=0.31, rely=0.2, anchor=tk.W)
        self.label_testnum.place(relx=0.25,rely=0.3,anchor=tk.W)
        self.label_testname.place(relx=0.25,rely=0.4,anchor=tk.W)
        self.input_year.place(relx=0.43, rely=0.2, anchor=tk.W)
        self.input_testnum.place(relx=0.43, rely=0.3, anchor=tk.W)
        self.input_testname.place(relx=0.45,rely=0.4,anchor=tk.W)
        self.button_back.place(relx=0.35,rely=0.8,anchor=tk.CENTER)
        self.button_OK.place(relx=0.65,rely=0.8,anchor=tk.CENTER)
        self.cc.place(relx=1, rely=1,anchor=tk.SE)
    def return_click(self, event):  #按Enter鍵自動連結登入
        self.OK_interface()
def main():  
    L = output_mySQL()
    L.gui_arrang()
    # 主程式執行  
    tk.mainloop()

if __name__ == '__main__':  
    main()  