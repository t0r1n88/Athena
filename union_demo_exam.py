"""
Функции для объединения отчетов по демоэкзамену
"""
import pandas as pd
import os
from tkinter import filedialog
from tkinter import messagebox
import time
pd.options.mode.chained_assignment = None  # default='warn'
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)


def select_folder_data_demo_exam():
    """
    Функция для выбора папки c данными от цопп по опк
    :return:
    """
    global path_folder_data_demo_exam
    path_folder_data_demo_exam = filedialog.askdirectory()


def select_end_folder_demo_exam():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_demo_exam
    path_to_end_folder_demo_exam = filedialog.askdirectory()


def processing_data_demo_exam():
    """
    Фугкция для объединения данных из всех трех листов из файлов от каждого региона ДФО
    :return:
    """
    try:
        form1_df = pd.DataFrame(columns=range(22))
        form2_df = pd.DataFrame(columns=range(11))
        form3_df = pd.DataFrame(columns=range(8))

        for file in os.listdir(path_folder_data_demo_exam):
            if (file.endswith('.xlsx') and not file.startswith('~$')):
                print(file)
                # обрабатываем первый лист
                temp_df = pd.read_excel(f'{path_folder_data_demo_exam}/{file}', sheet_name=0, dtype={'гр.2': str}, skiprows=2,
                                        usecols='A:V')
                temp_df.columns = range(22)
                temp_df.dropna(thresh=4, inplace=True)
                form1_df = pd.concat([form1_df, temp_df], ignore_index=True)

                # обрабатываем второй лист
                temp_df = pd.read_excel(f'{path_folder_data_demo_exam}/{file}', sheet_name=1, dtype={'гр.2': str}, skiprows=2,
                                        usecols='A:K')
                temp_df.columns = range(11)
                temp_df.dropna(thresh=4, inplace=True)
                form2_df = pd.concat([form2_df, temp_df], ignore_index=True)

                # обрабатываем третий лист
                temp_df = pd.read_excel(f'{path_folder_data_demo_exam}/{file}', sheet_name=2, dtype={'гр.2': str}, skiprows=2,
                                        usecols='A:H')
                temp_df.columns = range(8)
                temp_df.dropna(thresh=3, inplace=True)
                form3_df = pd.concat([form3_df, temp_df], ignore_index=True)

        # генерируем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)

        with pd.ExcelWriter(f'{path_to_end_folder_demo_exam}/Общий файл от {current_time}.xlsx') as writer:
            form1_df.to_excel(writer, sheet_name='Форма по мониторингу', index=False)
            form2_df.to_excel(writer, sheet_name='Форма по принимаемым мерам', index=False)
            form3_df.to_excel(writer, sheet_name='Форма по социальной поддежке', index=False)
    except NameError:
            messagebox.showerror('Афина',
                                 f'Выберите файлы с данными и папку куда будет генерироваться файл')
    except KeyError as e:
        messagebox.showerror('Афина',
                             f'Не найдено значение {e.args}')
    except FileNotFoundError:
        messagebox.showerror('Афина',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except PermissionError as e:
        messagebox.showerror('Афина',
                             f'Закройте открытые файлы Excel {e.args}')
    else:
        messagebox.showinfo('Афина',
                            'Данные успешно обработаны.')