import os
import re
import shutil
import pandas as pd
from dbfread import DBF
import sys

def clean_excel_string(s):
    if isinstance(s, str):
        return re.sub(r'[\x00-\x1F]', '', s)
    return s

# Рабочая папка (где лежит скрипт или exe)
if getattr(sys, 'frozen', False):
    work_folder = os.path.dirname(sys.executable)
else:
    work_folder = os.path.dirname(os.path.abspath(__file__))

# Сетевой путь
network_folder = r'\\192.168.1.160\Zapros - исправленная 06.10.22'
needed_files = ['zapros1.dbf', 'zapros2.dbf']

# Заголовки
new_headers = [
    'Регистрационный №', 'Адресант', 'Фамилия', 'Имя', 'Отчество',
    'Адрес заявителя', 'Дата запроса', 'Характер запроса', 'Содержание запроса',
    'Результат', 'Исполнитель', 'Дата исполнения', 'Используемые материалы',
    'Оплата', 'Поступление'
]

for filename in needed_files:
    source_path = os.path.join(network_folder, filename)
    dest_path = os.path.join(work_folder, filename)

    try:
        # Копируем файл в рабочую папку
        shutil.copy2(source_path, dest_path)
        print(f'📂 Скопирован {filename} в рабочую папку.')

        # Читаем DBF и обрабатываем
        table = DBF(dest_path, encoding='cp1251')
        df = pd.DataFrame(iter(table))

        for col in df.columns:
            df[col] = df[col].map(clean_excel_string)

        # Заголовки
        df.columns = new_headers[:len(df.columns)] + df.columns[len(new_headers):].tolist()

        # Удаляем существующие CSV файлы, если они есть
        csv_path = os.path.splitext(dest_path)[0] + '.csv'
        if os.path.exists(csv_path):
            os.remove(csv_path)
            print(f'🗑 Удалён существующий CSV файл: {csv_path}')

        # Экспортируем в CSV
        df.to_csv(csv_path, index=False, sep=';', encoding='utf-8-sig')
        print(f'✅ Экспортировано в CSV: {csv_path}')
        if '1' in filename:
            print('База социально-правовых запросов обновлена!\n')
        else:
            print('База тематических запросов обновлена!\n')
    except Exception as e:
        print(f'❌ Ошибка при обработке {filename}: {e}\nЗАКРОЙТЕ ЛИСИЧКУ!\n')

    finally:
        # Удаляем .dbf из рабочей папки
        if os.path.exists(dest_path):
            os.remove(dest_path)
            print(f'🗑 Удалён {filename} из рабочей папки.')
input('Нажмите Enter, чтобы завершить')
