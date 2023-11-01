import pandas as pd
import os
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
# объединение файлов ОПК по ДФО
from union_opk_dfo import select_folder_data_opk_dfo
from union_opk_dfo import select_end_folder_opk_dfo
from union_opk_dfo import processing_data_opk_dfo

# Объедениение файлов по демоэкзамену
from union_demo_exam import select_folder_data_demo_exam
from union_demo_exam import select_end_folder_demo_exam
from union_demo_exam import processing_data_demo_exam

# Извлечение данных из отчета ФГИС по региону
from report_fgis_base import create_base_report # создание отчета

# Извлечение данных по школам муниципалитетов
from report_fgsi_mun import create_report_mun


import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
pd.options.mode.chained_assignment = None  # default='warn'


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
    Функция чтобы логотип отображался"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)



def select_file_data_fgis():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global file_data_fgis
    # Получаем путь к файлу
    file_data_fgis = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_end_folder_fgis():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_fgis
    path_to_end_folder_fgis = filedialog.askdirectory()


def proccessing_report_fgis():
    """
    Функция для подсчета данных из файлов
    :return:
    """
    try:
        create_base_report(file_data_fgis,path_to_end_folder_fgis)

    except NameError:
        messagebox.showerror('Афина',
                             f'Выберите параметры ,папку или файл с данными и папку куда будут генерироваться файлы')

"""
Функции для отчетов по муниципалитетам
"""

def select_data_folder_mun_fgis():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_data_folder_mun_fgis
    path_to_data_folder_mun_fgis= filedialog.askdirectory()

def select_end_folder_mun_fgis():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_mun_fgis
    path_to_end_folder_mun_fgis= filedialog.askdirectory()


def proccessing_report_mun_fgis():
    """
    Функция для подсчета данных из файлов
    :return:
    """
    try:
        create_report_mun(path_to_data_folder_mun_fgis,path_to_end_folder_mun_fgis)

    except NameError:
        messagebox.showerror('Афина',
                             f'Выберите параметры ,папку или файл с данными и папку куда будут генерироваться файлы')






if __name__ == '__main__':
    window = Tk()
    window.title('Афина ver 1.4')
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


    """
    Обработка отчетов по демоэкзамену
    """
    # Создаем вкладку обработки данных для Приложения 6
    tab_demo_exam = ttk.Frame(tab_control)
    tab_control.add(tab_demo_exam, text='Скрипт ДЭ')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello_demo_exam = Label(tab_demo_exam,
                                text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                     'Объединение отчетов по демоэкзамену')
    lbl_hello_demo_exam.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_demo_exam = resource_path('logo.png')

    img__demo_exam = PhotoImage(file=path_to_img_demo_exam)
    Label(tab_demo_exam,
          image=img__demo_exam
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными
    btn_choose_data_demo_exam = Button(tab_demo_exam, text='1) Выберите папку с данными', font=('Arial Bold', 20),
                                     command=select_folder_data_demo_exam
                                     )
    btn_choose_data_demo_exam.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_demo_exam = Button(tab_demo_exam, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                             command=select_end_folder_demo_exam
                                             )
    btn_choose_end_folder_demo_exam.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку обработки данных

    btn_proccessing_data_demo_exam = Button(tab_demo_exam, text='3) Обработать данные', font=('Arial Bold', 20),
                                            command=processing_data_demo_exam
                                            )
    btn_proccessing_data_demo_exam.grid(column=0, row=4, padx=10, pady=10)


    """Создаем вкладку для ФГИС по региону """
    # Создаем вкладку обработки данных для Приложения 6
    tab_fgis = ttk.Frame(tab_control)
    tab_control.add(tab_fgis, text='ФГИС по региону')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello_fgis = Label(tab_fgis,
                              text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                   'Данные по муниципалитетам')
    lbl_hello_fgis.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_fgis = resource_path('logo.png')

    img_fgis = PhotoImage(file=path_to_img_fgis)
    Label(tab_fgis,
          image=img_fgis
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными
    btn_choose_data_fgis= Button(tab_fgis, text='1) Выберите файл с  данными', font=('Arial Bold', 20),
                                     command=select_file_data_fgis
                                     )
    btn_choose_data_fgis.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_fgis = Button(tab_fgis, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                           command=select_end_folder_fgis
                                           )
    btn_choose_end_folder_fgis.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку обработки данных

    btn_proccessing_data_fgis = Button(tab_fgis, text='3) Обработать данные', font=('Arial Bold', 20),
                                          command=proccessing_report_fgis
                                          )
    btn_proccessing_data_fgis.grid(column=0, row=4, padx=10, pady=10)


    """Создаем вкладку для отчетов ФГИС по школам муниципалитета """
    tab_mun_fgis = ttk.Frame(tab_control)
    tab_control.add(tab_mun_fgis, text='ФГИС по школам')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello_mun_fgis = Label(tab_mun_fgis,
                               text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                    'Данные по школам муниципалитетов')
    lbl_hello_mun_fgis.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_mun_fgis = resource_path('logo.png')

    img_mun_fgis = PhotoImage(file=path_to_img_mun_fgis)
    Label(tab_mun_fgis,
          image=img_mun_fgis
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными
    btn_choose_data_mun_fgis = Button(tab_mun_fgis, text='1) Выберите папку с данными', font=('Arial Bold', 20),
                                      command=select_data_folder_mun_fgis
                                      )
    btn_choose_data_mun_fgis.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_mun_fgis = Button(tab_mun_fgis, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                            command=select_end_folder_mun_fgis
                                            )
    btn_choose_end_folder_mun_fgis.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку обработки данных

    btn_proccessing_data_mun_fgis = Button(tab_mun_fgis, text='3) Обработать данные', font=('Arial Bold', 20),
                                           command=proccessing_report_mun_fgis
                                           )
    btn_proccessing_data_mun_fgis.grid(column=0, row=4, padx=10, pady=10)




    window.mainloop()
