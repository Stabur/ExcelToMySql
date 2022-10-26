# Конвертирует excel-файлы в MySql таблицу с последующим её заполнением.
# Удобен если необходимо создать/загрузить прайс-лист в базу MySql.
from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import *
import MySQLdb
import os
os.environ["CUDA_VISIBLE_DEVICES"] = "-1"
import os.path
import pandas as pd
from sshtunnel import SSHTunnelForwarder
from transliterate import translit
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
from sbcripto import *

Form, Window = uic.loadUiType('forma.ui')

app = QApplication([])
window = Window()
form = Form()
form.setupUi(window)
window.show()

server = []
dbconnect = []
ssh_pass = []
mysql_password = []

form.lineEdit_4.setEchoMode(QtWidgets.QLineEdit.Password)

def on_click_testssh():
    global server, ssh_host, ssh_port, ssh_login, ssh_pass, ssh_mysql_host, ssh_mysql_port
    ssh_host = form.lineEdit.text()
    ssh_port = form.lineEdit_2.text()
    ssh_login = form.lineEdit_3.text()
    ssh_pass = form.lineEdit_4.text()
    ssh_mysql_host = form.lineEdit_5.text()
    ssh_mysql_port = form.lineEdit_6.text()

    if not ssh_host:
        form.label_7.setText("<font color=red>Не заполнено поле Адрес SSH сервера!</font>")
    elif not ssh_port:
        form.label_7.setText("<font color=red>Не заполнено поле Порт SSH сервера!</font>")
    elif not ssh_login:
        form.label_7.setText("<font color=red>Не заполнено поле Логин SSH сервера!</font>")
    elif not ssh_pass:
        form.label_7.setText("<font color=red>Не заполнено поле Пароль SSH сервера!</font>")
    elif not ssh_mysql_host:
        form.label_7.setText("<font color=red>Не заполнено поле Адрес MySql сервера!</font>")
    elif not ssh_mysql_port:
        form.label_7.setText("<font color=red>Не заполнено поле Порт Mysql сервера!</font>")

    new_ssh_pass = str(ssh_pass)
    hash_ssh_pass = stabur_cripto(new_ssh_pass)
    #print('Пароль SSH : ' + new_ssh_pass)
    #print('Кэш SSH : ' + hash_ssh_pass)
    #hash_ssh_pass_len = len(hash_ssh_pass)
    #print('Кол-во символов в Кэше SSH : ' + str(hash_ssh_pass_len))

    if (ssh_host and ssh_port and ssh_login and ssh_pass and ssh_mysql_host and ssh_mysql_port):
        try:
            server = SSHTunnelForwarder(
                (ssh_host, int(ssh_port)),
                ssh_username=ssh_login,
                ssh_password=ssh_pass,
                remote_bind_address=(ssh_mysql_host, int(ssh_mysql_port))
            )
            if (server):
                print(server)
                print("Соединение c SSH-сервером установленно!")
                form.label_7.setText("<font color=green>Соединение c SSH-сервером установленно!</font>")
            else:
                print("Проблема с подключением к SSH-серверу!")
            server.start()
        except Exception as e:
            form.label_7.setText("<font color=red>Соединение c SSH-сервером НЕ установленно!</font>")
            print(f"""Ошибка: {e}""")

form.lineEdit_9.setEchoMode(QtWidgets.QLineEdit.Password)

def on_click_testbd():
    global server, dbconnect, mysql_host, mysql_port, mysql_login, mysql_password, mysql_db_name

    if (server):
        mysql_host = form.lineEdit_7.text()
        mysql_login = form.lineEdit_8.text()
        mysql_password = form.lineEdit_9.text()
        mysql_db_name = form.lineEdit_10.text()
        mysql_port = server.local_bind_port

        if not mysql_host:
            form.label_7.setText("<font color=red>Не заполнено поле Хост MySql!</font>")
        elif not mysql_login:
            form.label_7.setText("<font color=red>Не заполнено поле Логин MySql!</font>")
        elif not mysql_password:
            form.label_7.setText("<font color=red>Не заполнено поле Пароль MySql!</font>")
        elif not mysql_db_name:
            form.label_7.setText("<font color=red>Не заполнено поле База Данных MySql!</font>")

        sql_pass = str(mysql_password)
        hash_sql_pass = stabur_cripto(sql_pass)
        #print('Пароль MySQL : ' + sql_pass)
        #print('Кэш MySQL : ' + hash_sql_pass)
        #hash_sql_pass_len = len(hash_sql_pass)
        #print('Кол-во символов в Кэше MySQL : ' + str(hash_sql_pass_len))

        if (mysql_host and mysql_login and mysql_password and mysql_db_name):
            try:
                dbconnect = MySQLdb.connect(mysql_host, mysql_login, mysql_password, mysql_db_name, mysql_port)
                form.label_7.setText("<font color=green>Соединение c MySql установленно!</font>")
                dbconnect.commit()
                dbconnect.rollback()
                dbconnect.close()
            except:
                form.label_7.setText("<font color=red>Соединение c MySql НЕ установленно!</font>")
    else:
        print("Переменные 'server' или 'dbconnect' не найдены!")
        form.label_7.setText("<font color=red>Нет подключения к SSH серверу!</font>")

def on_click_excfile():
    global mysql_table_name, mysql_name_col, mysql_ef2, mysql_name_col_insert
    excelfile = QFileDialog.getOpenFileName(None, 'OpenFile', '', 'Excel file(*.xlsx)')
    excelfilePath = excelfile[0]
    # print(excelfilePath) # полный путь к файлу на ПК
    exPath = os.path.dirname(excelfilePath)
    exFile = os.path.basename(excelfilePath)
    exFile = translit(exFile, reversed=True)
    exFile = exFile.replace(" ", "_")
    exFile = exFile.replace("'", "")
    # print(exPath) # директория файла на ПК
    # print(exFile) # название файла с расширением
    mysql_table_name = os.path.splitext(exFile)[0]
    #print(mysql_table_name)  # название файла без расширения
    form.label_6.setText(f"<font color=green>Файл <font color=blue>{exFile}</font> прикреплён!</font>")
    ef = pd.read_excel(excelfilePath)
    mysql_ef2 = pd.read_excel(excelfilePath, dtype=object)
    # tosql = ef.head()
    # tosql.index = tosql.index + 1  # Содержимое файла с разбитием на нумерацию с 1-го
    # print(tosql.index)
    strhead = ef.columns.ravel()  # Массив, заголовки столбцов
    # print(strhead)
    # num_strhead = len(strhead) # Кол-во элементов в массиве
    i = (",".join(strhead))
    strhead_en = translit(i, reversed=True)  # Транслит (если вдруг русские названия) заголовков столбцов
    # print(strhead_en)
    strhead_en = strhead_en.replace(" ", "_")
    strhead_en = strhead_en.replace(",", ", ")
    strhead_en = strhead_en.replace(".", "")

    strhead_en_arr = strhead_en.split(', ')  # Добавляем элементы в массив и разделяем их запятыми
    # print(strhead_en_arr)
    name_col = [i + ' VARCHAR(255)  not null' for i in strhead_en_arr]
    mysql_name_col = (', '.join(map(str, name_col)))

    name_col_insert = [i for i in strhead_en_arr]
    name_col_insert = (','.join(map(str, name_col_insert)))
    name_col_insert = name_col_insert.replace(" ", "_")
    name_col_insert = name_col_insert.replace(",", ", ")
    mysql_name_col_insert = name_col_insert.replace(".", "")
    #print(name_col)
    #print(name_col_insert)

def on_click_create():
    global server, mysql_port, dbconnect, mysql_host, mysql_login, mysql_password, mysql_db_name, mysql_table_name, mysql_name_col, mysql_name_col_insert, mysql_ef2

    if (server and dbconnect):
        if not mysql_host:
            form.label_7.setText("<font color=red>Не заполнено поле Хост MySql!</font>")
        elif not mysql_login:
            form.label_7.setText("<font color=red>Не заполнено поле Логин MySql!</font>")
        elif not mysql_password:
            form.label_7.setText("<font color=red>Не заполнено поле Пароль MySql!</font>")
        elif not mysql_db_name:
            form.label_7.setText("<font color=red>Не заполнено поле База Данных MySql!</font>")
        elif not mysql_port:
            form.label_7.setText("<font color=red>Не заполнено поле Порт MySql!</font>")

        if (mysql_host and mysql_login and mysql_password and mysql_db_name):
            dbcon = MySQLdb.connect(mysql_host, mysql_login, mysql_password, mysql_db_name, mysql_port)
            cur = dbcon.cursor()
            cur2 = dbcon.cursor()
            try:
                cur.execute(f"""CREATE TABLE IF NOT EXISTS {mysql_table_name} (id INT(12) auto_increment not null primary key, {mysql_name_col})""")
                form.label_7.setText(f"<font color=green>Таблица <font color=blue>{mysql_table_name}</font> в базе данных создана!!!</font>")
            except Exception as e:
                form.label_7.setText("<font color=red>Соединение c MySql-сервером НЕ установленно!</font>")
                print(f"""Ошибка: {e}""")
                dbcon.rollback()

            for body in range(len(mysql_ef2)):
                mysql_rt = tuple(mysql_ef2.iloc[body].to_list())
                print(mysql_rt)
                cur2.execute(f"""INSERT INTO {mysql_table_name} ({mysql_name_col_insert}) VALUES {mysql_rt}""")

            dbcon.commit()
            dbcon.close()
        else:
            form.label_7.setText("<font color=red>Ошибка заполнения!!!</font>")
    else:
        print("Переменные 'server' или 'dbconnect' не найдены!")
        form.label_7.setText("<font color=red>Нет подключения к SSH и/или MySql серверу!</font>")

def on_click_windows():
    print("Закрыть Окно")
    quit()

form.pushButton.clicked.connect(on_click_windows)
form.pushButton_2.clicked.connect(on_click_testbd)
form.pushButton_3.clicked.connect(on_click_excfile)
form.pushButton_4.clicked.connect(on_click_create)
form.pushButton_5.clicked.connect(on_click_testssh)

app.exec_()

server.stop()
