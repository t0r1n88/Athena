from faker import Faker
import pandas as pd
import random
#
# fake = Faker('ru_RU')
#
# df = pd.DataFrame()
#
# fio = [fake.name() for i in range(17542)]
# date_birth = [fake.date() for i in range(17542)]
# mail = [fake.free_email() for i in range(17542)]
#
# df['ФИО'] = fio
# df['Дата рождения'] = date_birth
# df['email'] = mail
# df.to_excel('data/Список.xlsx',index=False)

df = pd.read_excel('data/Билет в будущее.xlsx')
print(df.shape[1])
fio_df = pd.read_excel('data/Список педагогов.xlsx')
fio_lst = fio_df['ФИО'].tolist()

df['ответственный педагог-навигатор'] = [fio_lst[random.randint(0,48)] for i in range(df.shape[0])]

df.to_excel('data/Билет в будущее сводный отчет по ученикам.xlsx',index=False)