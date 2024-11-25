import pandas as pd
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# Путь к файлу учетных данных
CREDENTIALS_FILE = "credentials.json"

# ID Google Таблицы
SPREADSHEET_ID = "1ARwjDUy3KvxbZyDrkMsfoz4EZnbI_SqGG4ldjDW8dhc"

# Название нового листа
NEW_SHEET_NAME = "Summary"

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

# Ваш DataFrame (пример, если он уже создан)
# summary_df = pd.DataFrame({...})
 
# assign data of lists. 
data = {'Name': ['Tom', 'Joseph', 'Krish', 'John'], 'Age': [20, 21, 19, 18]} 
 
# Create DataFrame 
summary_df = pd.DataFrame(data)

# Print the output. 
print(summary_df)
# Запись DataFrame в Google Таблицу
write_to_google_sheets(summary_df, SPREADSHEET_ID, NEW_SHEET_NAME)

