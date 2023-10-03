import pandas as pd
import os
from tkinter import *
from tkinter import ttk
# объединение файлов ОПК по ДФО
from union_opk_dfo import select_folder_data_opk_dfo
from union_opk_dfo import select_end_folder_opk_dfo
from union_opk_dfo import processing_data_opk_dfo

# Объедениение файлов по демоэкзамену
from union_demo_exam import select_folder_data_demo_exam
from union_demo_exam import select_end_folder_demo_exam
from union_demo_exam import processing_data_demo_exam
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


if __name__ == '__main__':
    window = Tk()
    window.title('Афина ver 1.2')
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





    window.mainloop()
