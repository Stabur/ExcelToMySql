"""
   Автор: StaBur
   Дата: 19.12.2022; 6:05 GMT-3;
   Конвертирует excel-файл (прайс-лист или перечень товара с описанием) в MySql таблицу.
   Удобен если необходимо создать/загрузить прайс-лист||перечень товара  с большим объемом данных в базу MySql.
"""

import MySQLdb
from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QApplication, QFileDialog

import os

# os.environ["CUDA_VISIBLE_DEVICES"] = "-1"
import os.path
import warnings
import pandas as pd
import numpy as np
from sshtunnel import SSHTunnelForwarder
import re
from translit import transliterate

warnings.simplefilter(action='ignore', category=FutureWarning)

Form, Window = uic.loadUiType("forma.ui")

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

        if (mysql_host and mysql_login and mysql_password and mysql_db_name):
            try:
                dbconnect = MySQLdb.connect(mysql_host, mysql_login, mysql_password, mysql_db_name, mysql_port)
                form.label_7.setText("<font color=green>Соединение c MySql установленно!</font>")
                print(f"Соединение c MySql установленно!\n порт: {mysql_port}")
                dbconnect.commit()
                dbconnect.rollback()
                dbconnect.close()
            except:
                form.label_7.setText("<font color=red>Соединение c MySql НЕ установленно!</font>")
    else:
        print("Переменные 'server' или 'dbconnect' не найдены!")
        form.label_7.setText("<font color=red>Нет подключения к SSH серверу!</font>")


def on_click_excfile():
    global mysql_table_name, mysql_name_col, mysql_name_col_insert, mysql_ef2
    excelfile = QFileDialog.getOpenFileName(None, 'OpenFile', '', 'Excel file(*.xlsx)')
    excelfilePath = excelfile[0]
    exFile = os.path.basename(excelfilePath) # Название файла как есть
    table_name = os.path.splitext(exFile)[0]  # Убераем расширение

    table_nameEng = transliterate(table_name) # Транслит, если имя файла на русском

    mysql_table_name = re.sub("[^0-9,A-Za-z]", "", table_nameEng) # Оставляем в названии только буквы и цифры
    print(f"Название таблицы в MySql : {mysql_table_name}")
    form.label_6.setText(f"<font color=green>Файл <font color=blue>{exFile}</font> прикреплён!</font>")

    ef = pd.read_excel(excelfilePath) # Считываем загруженный файл

    strhead = ef.columns.ravel()  # Массив, заголовки столбцов
    strhead = [x.replace(",", "") for x in strhead] # Убираем запятые из элементов массива
    strhead = [x.strip(' ') for x in strhead]  # Удаляем пробелы вначале и в конце элементов массива

    strhead_str = (", ".join(strhead)) # формируем список из массива

    strhead_str = re.sub('[^а-яА-Я,ёЁ,йЙ,0-9a-zA-Z,#,№]', '', strhead_str) # Убераем все спец.символы, оставляем лишь буквы и цифры

    strhead_en = transliterate(strhead_str)  # Транслит (если вдруг русские названия) заголовков столбцов

    strhead_en_arr = strhead_en.split(',')  # Разделяем строку
    print(f"Наименование колонок для таблицы {mysql_table_name} : {strhead_en_arr}")
    name_col = [i + ' TEXT not null' for i in strhead_en_arr] # Добавляем ' TEXT not null' к каждому элементу
    
    mysql_name_col = (', '.join(name_col)) # Формируем список для sql create table
    print(f"Для Create таблицы {mysql_table_name} : {mysql_name_col}")
    mysql_name_col_insert = strhead_en # Cписок для sql insert
    print(f"Для Insert таблицы {mysql_table_name} : {mysql_name_col_insert}")
    # Создаем массив mysql_ef2 из содержимого таблицы, с последующей  передачей его в def on_click_create
    mysql_ef2 = pd.read_excel(excelfilePath)  # Извлекаем содержимое файлa
    mysql_ef2 = mysql_ef2.astype(object).replace(np.nan, 'None') # Пустые ячейки помечаем как None
    print(f"Содержание 'mysql_ef2' : {mysql_ef2}")

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
                print(f"Таблица {mysql_table_name} создана")
            except Exception as e:
                form.label_7.setText("<font color=red>Соединение c MySql-сервером НЕ установленно!</font>")
                print(f"""Ошибка: {e}""")
                dbcon.rollback()

            for body in range(len(mysql_ef2)):
                mysql_rt = tuple(mysql_ef2.iloc[body].to_list())
                try:
                    cur2.execute(f"""INSERT INTO {mysql_table_name} ({mysql_name_col_insert}) VALUES {mysql_rt}""")
                    print(f"{body+1}-я строка таблицы {mysql_table_name} добавленна")
                except:
                    form.label_7.setText(f"<font color=red>Содержание таблицы {mysql_table_name} не добавленно!</font></font>")
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
