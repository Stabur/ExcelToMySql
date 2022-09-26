from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QFileDialog
import MySQLdb
import os
import os.path
import pandas as pd
import openpyxl
from sshtunnel import SSHTunnelForwarder
from transliterate import translit, get_available_language_codes

Form, Window = uic.loadUiType("forma.ui")

app = QApplication([])
window = Window()
form = Form()
form.setupUi(window)
window.show()

server = SSHTunnelForwarder(
    ('192.168.1.35', 22),
    ssh_username='burnov',
    ssh_password='01021989',
    remote_bind_address=('127.0.0.1', 3306)
)

if (server):
    print(server)
    print("Соединение c ssh-сервером установленно!")
else:
    print("Проблема с подключением к ssh-серверу!")
server.start()

def on_click_testbd():
    global host, login, password, db_name, table_name

    host = form.lineEdit.text()
    login = form.lineEdit_2.text()
    password = form.lineEdit_3.text()
    db_name = form.lineEdit_4.text()
    port = server.local_bind_port

    if not host:
        form.label_7.setText("<font color=red>Не заполнено поле Хост!</font>")
    elif not login:
        form.label_7.setText("<font color=red>Не заполнено поле Логин!</font>")
    elif not password:
        form.label_7.setText("<font color=red>Не заполнено поле Пароль!</font>")
    elif not db_name:
        form.label_7.setText("<font color=red>Не заполнено поле База Данных!</font>")

    if (host and login and password and db_name):
        try:
            dbconnect = MySQLdb.connect(host, login, password, db_name, port)
            form.label_7.setText("<font color=green>Соединение установленно!</font>")
            dbconnect.commit()
            dbconnect.rollback()
            dbconnect.close()
        except:
            form.label_7.setText("<font color=red>Соединение НЕ установленно!</font>")

def on_click_excfile():
    global host, login, password, db_name, table_name, name_col, ef2, name_col_insert
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
    table_name = os.path.splitext(exFile)[0]
    #print(table_name)  # название файла без расширения
    form.label_6.setText(f"<font color=green>Файл <font color=blue>{exFile}</font> прикреплён!</font>")
    ef = pd.read_excel(excelfilePath)
    ef2 = pd.read_excel(excelfilePath, dtype=object)
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
    name_col = (', '.join(map(str, name_col)))

    name_col_insert = [i for i in strhead_en_arr]
    name_col_insert = (','.join(map(str, name_col_insert)))
    name_col_insert = name_col_insert.replace(" ", "_")
    name_col_insert = name_col_insert.replace(",", ", ")
    name_col_insert = name_col_insert.replace(".", "")
    # print(name_col)
    #print(name_col_insert)

def on_click_create():
    global host, login, password, db_name, table_name, name_col, ef2
    host = form.lineEdit.text()
    login = form.lineEdit_2.text()
    password = form.lineEdit_3.text()
    db_name = form.lineEdit_4.text()
    port = server.local_bind_port
    if not host:
        form.label_7.setText("<font color=red>Не заполнено поле Хост!</font>")
    elif not login:
        form.label_7.setText("<font color=red>Не заполнено поле Логин!</font>")
    elif not password:
        form.label_7.setText("<font color=red>Не заполнено поле Пароль!</font>")
    elif not db_name:
        form.label_7.setText("<font color=red>Не заполнено поле База Данных!</font>")

    if (host and login and password and db_name and table_name and name_col):
        dbconnect = MySQLdb.connect(host, login, password, db_name, port)
        cur = dbconnect.cursor()
        cur2 = dbconnect.cursor()
        try:
            cur.execute(
                f"""CREATE TABLE IF NOT EXISTS {table_name} (id INT(12) auto_increment not null primary key, {name_col})""")
            form.label_7.setText(
                f"<font color=green>Таблица <font color=blue>{table_name}</font> в базе данных создана!!!</font>")
        except:
            dbconnect.rollback()

        for body in range(len(ef2)):
            rt = tuple(ef2.iloc[body].to_list())
            print(rt)
            cur2.execute(f"""INSERT INTO {table_name} ({name_col_insert}) VALUES {rt}""")

        dbconnect.commit()
        dbconnect.close()
    else:
        form.label_7.setText("<font color=red>Ошибка заполнения!!!</font>")

def on_click_windows():
    print("Закрыть Окно")
    quit()

form.pushButton.clicked.connect(on_click_windows)
form.pushButton_2.clicked.connect(on_click_testbd)
form.pushButton_3.clicked.connect(on_click_excfile)
form.pushButton_4.clicked.connect(on_click_create)

app.exec_()

server.stop()
