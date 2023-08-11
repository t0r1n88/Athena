import tkinter
import sys
import pandas as pd
import openpyxl
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import time

pd.options.mode.chained_assignment = None  # default='warn'
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
# объединение файлов ОПК по ДФО
from union_opk_dfo import select_folder_data_opk_dfo
from union_opk_dfo import select_end_folder_opk_dfo
from union_opk_dfo import processing_data_opk_dfo

# Объедениение файлов по демоэкзамену
from union_demo_exam import select_folder_data_demo_exam
from union_demo_exam import select_end_folder_demo_exam
from union_demo_exam import processing_data_demo_exam


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
    Функция чтобы логотип отображался"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


if __name__ == '__main__':
    window = Tk()
    window.title('Афина ver 1.0')
    window.geometry('700x860')
    window.resizable(False, False)

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку обработки данных для Приложения 6
    tab_opk_dfo = ttk.Frame(tab_control)
    tab_control.add(tab_opk_dfo, text='Скрипт ОПК ДФО')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello_opk_dfo = Label(tab_opk_dfo,
                              text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                   'Обработка файлов трудоустройства ОПК от регионов ДФО ')
    lbl_hello_opk_dfo.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_opk_dfo = resource_path('logo.png')

    img_opk_dfo = PhotoImage(file=path_to_img_opk_dfo)
    Label(tab_opk_dfo,
          image=img_opk_dfo
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными
    btn_choose_data_opk_dfo = Button(tab_opk_dfo, text='1) Выберите папку с данными', font=('Arial Bold', 20),
                                     command=select_folder_data_opk_dfo
                                     )
    btn_choose_data_opk_dfo.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_opk_dfo = Button(tab_opk_dfo, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                           command=select_end_folder_opk_dfo
                                           )
    btn_choose_end_folder_opk_dfo.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку обработки данных

    btn_proccessing_data_opk_dfo = Button(tab_opk_dfo, text='3) Обработать данные', font=('Arial Bold', 20),
                                          command=processing_data_opk_dfo
                                          )
    btn_proccessing_data_opk_dfo.grid(column=0, row=4, padx=10, pady=10)

    window.mainloop()
