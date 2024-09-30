"""
Функции для объединения файлов по трудоустройству ОПК в ДФо
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


def select_folder_data_opk_dfo():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_folder_data_opk_dfo
    path_folder_data_opk_dfo = filedialog.askdirectory()



def select_end_folder_opk_dfo():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_opk_dfo
    path_to_end_folder_opk_dfo = filedialog.askdirectory()


def processing_data_opk_dfo():
    """
    Фугкция для объединения данных из всех трех листов из файлов от каждого региона ДФО
    :return:
    """
    try:
        form1_df = pd.DataFrame(columns=[f'гр.{i}' for i in range(1,21)])
        form2_df = pd.DataFrame(columns=[f'гр.{i}' for i in range(1,12)])
        form3_df = pd.DataFrame(columns=[f'гр.{i}' for i in range(1,9)])

        for file in os.listdir(path_folder_data_opk_dfo):
            if (file.endswith('.xlsx') and not file.startswith('~$')):
                print(file)
                # обрабатываем первый лист
                temp_df = pd.read_excel(f'{path_folder_data_opk_dfo}/{file}', sheet_name=0, dtype={'гр.2': str}, skiprows=2,
                                        usecols='A:T')
                cols_df = [f'гр.{i}' for i in range(1,21)]
                temp_df.columns = cols_df
                temp_df.dropna(thresh=4, inplace=True)
                form1_df = pd.concat([form1_df, temp_df], ignore_index=True)

                # обрабатываем второй лист
                temp_df = pd.read_excel(f'{path_folder_data_opk_dfo}/{file}', sheet_name=1, dtype={'гр.2': str}, skiprows=2,
                                        usecols='A:K')
                cols_df = [f'гр.{i}' for i in range(1, 12)]
                temp_df.columns = cols_df
                temp_df.dropna(thresh=4, inplace=True)
                form2_df = pd.concat([form2_df, temp_df], ignore_index=True)

                # обрабатываем третий лист
                temp_df = pd.read_excel(f'{path_folder_data_opk_dfo}/{file}', sheet_name=2, dtype={'гр.2': str}, skiprows=2,
                                        usecols='A:H')
                cols_df = [f'гр.{i}' for i in range(1, 9)]
                temp_df.columns = cols_df
                temp_df.dropna(thresh=3, inplace=True)
                form3_df = pd.concat([form3_df, temp_df], ignore_index=True)

        # генерируем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)

        with pd.ExcelWriter(f'{path_to_end_folder_opk_dfo}/Общий свод ОПК по ДФО от {current_time}.xlsx') as writer:
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


if __name__ == '__main__':
    global path_folder_data_opk_dfo
    path_folder_data_opk_dfo = 'data/DFO'
    global path_to_end_folder_opk_dfo
    path_to_end_folder_opk_dfo = 'data/Результат'
    processing_data_opk_dfo()


