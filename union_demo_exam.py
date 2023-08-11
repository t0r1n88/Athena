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
        error_df = pd.DataFrame(
            columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])  # датафрейм с ошибками
        # Создаем датафреймы в которые будем собирать данные
        df_1 = pd.DataFrame(columns=range(8))

        df_2 = pd.DataFrame(columns=range(17))

        df_3 = pd.DataFrame(columns=range(16))

        df_4 = pd.DataFrame(columns=range(22))

        df_5 = pd.DataFrame(columns=range(30))

        df_6 = pd.DataFrame(columns=range(12))

        df_7 = pd.DataFrame(columns=range(5))

        df_8 = pd.DataFrame(columns=range(9))

        df_9 = pd.DataFrame(columns=range(17))

        empty_list = 'Лист не заполнен'  # строка для незаполненных листов

        for file in os.listdir(path_folder_data_demo_exam):
            if not file.startswith('~$') and file.endswith('.xls'):
                name_file = file.split('.xls')[0]
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                    'Файл с расширением XLS (СТАРЫЙ ФОРМАТ EXCEL)!!! ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue

            if not file.startswith('~$') and file.endswith('.xlsx'):
                name_file = file.split('.xlsx')[0]
                print(name_file)

                # Лист 1
                temp_df_1 = pd.read_excel(f'{path_folder_data_demo_exam}/{file}', sheet_name='Форма1', skiprows=3,
                                          usecols='A:G', header=None)

                temp_df_1.dropna(inplace=True, thresh=2)  # удаляем пустые строки

                if temp_df_1.shape[
                    0] == 0:  # если лист пуст то добавляем строку чтобы можно было потом понять что лист не заполнен
                    temp_df_1[len(temp_df_1)] = [empty_list]
                temp_df_1[7] = name_file

                df_1 = pd.concat([df_1, temp_df_1], ignore_index=True)

                # Лист 2

                temp_df_2 = pd.read_excel(f'{path_folder_data_demo_exam}/{file}', sheet_name='Форма2', skiprows=3,
                                          usecols='A:P', header=None)
                temp_df_2.dropna(inplace=True, thresh=2)  # удаляем пустые строки
                if temp_df_2.shape[
                    0] == 0:  # если лист пуст то добавляем строку чтобы можно было потом понять что лист не заполнен
                    temp_df_2[len(temp_df_2)] = [empty_list]

                temp_df_2[16] = name_file

                df_2 = pd.concat([df_2, temp_df_2], ignore_index=True)

                # Лист 3
                temp_df_3 = pd.read_excel(f'{path_folder_data_demo_exam}/{file}', sheet_name='Форма3', skiprows=3,
                                          usecols='A:O', header=None)
                temp_df_3.dropna(inplace=True, thresh=2)  # удаляем пустые строки
                if temp_df_3.shape[
                    0] == 0:  # если лист пуст то добавляем строку чтобы можно было потом понять что лист не заполнен
                    temp_df_3[len(temp_df_3)] = [empty_list]
                temp_df_3[15] = name_file

                df_3 = pd.concat([df_3, temp_df_3], ignore_index=True)

                # Лист 4
                temp_df_4 = pd.read_excel(f'{path_folder_data_demo_exam}/{file}', sheet_name='Форма4', skiprows=3,
                                          usecols='A:U', header=None)
                temp_df_4.dropna(inplace=True, thresh=2)  # удаляем пустые строки
                if temp_df_4.shape[
                    0] == 0:  # если лист пуст то добавляем строку чтобы можно было потом понять что лист не заполнен
                    temp_df_4[len(temp_df_4)] = [empty_list]
                temp_df_4[22] = name_file

                df_4 = pd.concat([df_4, temp_df_4], ignore_index=True)

                # Лист 5
                temp_df_5 = pd.read_excel(f'{path_folder_data_demo_exam}/{file}', sheet_name='Форма5', skiprows=3,
                                          usecols='A:AC', header=None)
                temp_df_5.dropna(inplace=True, thresh=2)  # удаляем пустые строки
                if temp_df_5.shape[
                    0] == 0:  # если лист пуст то добавляем строку чтобы можно было потом понять что лист не заполнен
                    temp_df_5[len(temp_df_5)] = [empty_list]
                temp_df_5[29] = name_file

                df_5 = pd.concat([df_5, temp_df_5], ignore_index=True)

                # Лист 6
                temp_df_6 = pd.read_excel(f'{path_folder_data_demo_exam}/{file}', sheet_name='Форма6', skiprows=3,
                                          usecols='A:K', header=None)
                temp_df_6.dropna(inplace=True, thresh=2)  # удаляем пустые строки
                if temp_df_6.shape[
                    0] == 0:  # если лист пуст то добавляем строку чтобы можно было потом понять что лист не заполнен
                    temp_df_6[len(temp_df_6)] = [empty_list]
                temp_df_6[11] = name_file

                df_6 = pd.concat([df_6, temp_df_6], ignore_index=True)

                # Лист 7
                temp_df_7 = pd.read_excel(f'{path_folder_data_demo_exam}/{file}', sheet_name='Форма7', skiprows=3,
                                          usecols='A:D', header=None)
                temp_df_7.dropna(inplace=True, thresh=2)  # удаляем пустые строки
                if temp_df_7.shape[
                    0] == 0:  # если лист пуст то добавляем строку чтобы можно было потом понять что лист не заполнен
                    temp_df_7[len(temp_df_7)] = [empty_list]

                temp_df_7[5] = name_file

                df_7 = pd.concat([df_7, temp_df_7], ignore_index=True)

                # Лист 8
                temp_df_8 = pd.read_excel(f'{path_folder_data_demo_exam}/{file}', sheet_name='Форма8', skiprows=3,
                                          usecols='A:H', header=None)
                temp_df_8.dropna(inplace=True, thresh=2)  # удаляем пустые строки
                if temp_df_8.shape[
                    0] == 0:  # если лист пуст то добавляем строку чтобы можно было потом понять что лист не заполнен
                    temp_df_8[len(temp_df_8)] = [empty_list]

                temp_df_8[9] = name_file

                df_8 = pd.concat([df_8, temp_df_8], ignore_index=True)

                # Лист 9
                temp_df_9 = pd.read_excel(f'{path_folder_data_demo_exam}/{file}', sheet_name='Форма9', skiprows=4,
                                          usecols='A:P', header=None)
                temp_df_9.dropna(inplace=True, thresh=2)  # удаляем пустые строки
                if temp_df_9.shape[
                    0] == 0:  # если лист пуст то добавляем строку чтобы можно было потом понять что лист не заполнен
                    temp_df_9[len(temp_df_9)] = [empty_list]

                temp_df_9[17] = name_file

                df_9 = pd.concat([df_9, temp_df_9], ignore_index=True)

        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)

        with pd.ExcelWriter(f'{path_to_end_folder_demo_exam}/Общий свод по демоэкзаменам от {current_time}.xlsx') as writer:
            df_1.to_excel(writer, sheet_name='Форма1', index=False)
            df_2.to_excel(writer, sheet_name='Форма2', index=False)
            df_3.to_excel(writer, sheet_name='Форма3', index=False)
            df_4.to_excel(writer, sheet_name='Форма4', index=False)
            df_5.to_excel(writer, sheet_name='Форма5', index=False)
            df_6.to_excel(writer, sheet_name='Форма6', index=False)
            df_7.to_excel(writer, sheet_name='Форма7', index=False)
            df_8.to_excel(writer, sheet_name='Форма8', index=False)
            df_9.to_excel(writer, sheet_name='Форма9', index=False)
        # сохраняем датафрейм с ошибками
        error_df.to_excel(f'{path_to_end_folder_demo_exam}/Ошибки {current_time}.xlsx', index=False)


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
