import os
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# Путь к файлу учетных данных
CREDENTIALS_FILE = "credentials.json"

# ID Google Таблицы
SPREADSHEET_ID = "1ARwjDUy3KvxbZyDrkMsfoz4EZnbI_SqGG4ldjDW8dhc"

# Название нового листа
NEW_SHEET_NAME = "Summary2"

# Путь к папке с Excel-файлами
input_folder = "output"
output_file = "project_week_summary.xlsx"

# Список для хранения данных
data = []

# Авторизация
def authorize():
    creds = Credentials.from_service_account_file(
        CREDENTIALS_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds)

# Запись данных в Google Таблицу
def write_to_google_sheets(dataframe, spreadsheet_id, sheet_name):
    service = authorize()
    sheets = service.spreadsheets()
    # Добавление нового листа
    sheets.batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [
                {"addSheet": {"properties": {"title": sheet_name}}}
            ]
        }
    ).execute()
    # Замена NaN на пустые строки и приведение данных к строковому формату
    dataframe = dataframe.fillna("").astype(str)
    # Преобразование DataFrame в список
    data = [dataframe.columns.tolist()] + dataframe.values.tolist()
    # Запись данных
    sheets.values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A1",
        valueInputOption="RAW",
        body={"values": data}
    ).execute()
    print(f"Данные записаны в лист '{sheet_name}'")

# Путь к папке с Excel-файлами
input_folder = "output"
output_file = "project_week_summary.xlsx"

# Список для хранения данных
data = []

# Обход всех файлов в папке
for filename in os.listdir(input_folder):
    if filename.endswith(".xlsx") and "_" in filename:
        # Извлечение имени проекта и дат недели из названия файла
        project, week_dates = filename.split("__")
        week_dates = week_dates.replace(".xlsx", "")
        # Путь к файлу
        file_path = os.path.join(input_folder, filename)
        # Загрузка данных
        df = pd.read_excel(file_path)
        # Суммирование числовых значений
        numeric_sums = df.select_dtypes(include=["number"]).sum()
        # Добавление данных о проекте и неделе
        numeric_sums["Проект"] = project
        numeric_sums["Неделя"] = week_dates
        # Добавление в общий список
        data.append(numeric_sums)

# Создание итогового DataFrame
summary_df = pd.DataFrame(data)

# Упорядочивание столбцов: Проект, Неделя + остальные
columns_order = ["Проект", "Неделя"] + [
    col for col in summary_df.columns if col not in ["Проект", "Неделя"]
]
summary_df = summary_df[columns_order]

# Сохранение сводной таблицы в Excel
#summary_df.to_excel(output_file, index=False)

#print(f"Сводная таблица сохранена в файле: {output_file}")

# Запись DataFrame в Google Таблицу
write_to_google_sheets(summary_df, SPREADSHEET_ID, NEW_SHEET_NAME)
