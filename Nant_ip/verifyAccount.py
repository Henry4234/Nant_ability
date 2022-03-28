from multiprocessing import connection
import pickle
import pymysql

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
with conn.cursor() as cursor:
    cursor.execute("SELECT `id_name`, `pw` FROM `id`;")

a = cursor.fetchall()
a=dict((x, y) for x, y in a)

def verifyAccountData(account,password):
    # with open('pw.pickle','rb') as usr_file:
    #         usrs_info=pickle.load(usr_file)
    if account in a:
        if account == "admin"and password == a[account]:
            return "master"
        elif password == a[account]:
            return "user"
        else:
            return "noPassword"
    #使用者名稱密碼不能為空
    elif account=='' or password=='' :
        return "empty"
    #不在資料庫中彈出是否註冊的框
    else:
        return "noAccount"

def changepw(account,password):
    if account in a:
        with conn.cursor() as cursor:
            ch = "UPDATE `id` SET `pw`=%s WHERE `id_name`='%s';"%(password,account)
            cursor.execute(ch)
        conn.commit()
        conn.close()
        return "changesuccess"
    else:
        return "noAccount"