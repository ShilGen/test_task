from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.service_account import Credentials
import os
import zipfile
import pandas as pd
from io import BytesIO

# Укажите путь к вашему credentials.json
CREDENTIALS_FILE = 'credentials.json'
FOLDER_ID = '1H9_lQojUD1QDBqLV9iv9Tfza2PI2x6hV'

# Авторизация
def authorize():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=["https://www.googleapis.com/auth/drive"])
    return build('drive', 'v3', credentials=creds)

# Получение списка файлов в папке
def list_files(service, folder_id):
    query = f"'{folder_id}' in parents and trashed = false"
    results = service.files().list(q=query, fields="files(id, name, mimeType)").execute()
    return results.get('files', [])

def list_folders_in_folder(service, folder_id):
    """Получить список вложенных папок по ID родительской папки"""
    query = f"'{folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    folders = []
    page_token = None
    while True:
        response = service.files().list(
            q=query,
            spaces='drive',
            fields="nextPageToken, files(id, name)",
            pageToken=page_token
        ).execute()
        folders.extend(response.get('files', []))
        page_token = response.get('nextPageToken', None)
        if not page_token:
            break
    return folders




# Скачивание файла
def download_file(service, file_id, file_name):
    request = service.files().get_media(fileId=file_id)
    fh = BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)
    with open(file_name, 'wb') as f:
        f.write(fh.read())

# Обработка файлов
def process_files(service, folder_id, output_folder='output'):
    os.makedirs(output_folder, exist_ok=True)
    # Получить список вложенных папок
    folders = list_folders_in_folder(service, folder_id)

    # Вывод списка папок
    print("Вложенные папки:")
    for folder in folders:
        print(f"Название: {folder['name']}, ID: {folder['id']}")
        week_folders = list_folders_in_folder(service, folder['id'])
        for folder in week_folders:
            print(f"- Название: {folder['name']}, ID: {folder['id']}")
#    for file in files:
#        if file['mimeType'] == 'application/vnd.google-apps.folder':  # Это папка
#            subfolder_path = os.path.join(output_folder, file['name'])
#            os.makedirs(subfolder_path, exist_ok=True)
#            process_files(service, file['id'], subfolder_path)
#        elif file['mimeType'] in ['application/zip', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
#            file_path = os.path.join(output_folder, file['name'])
#            download_file(service, file['id'], file_path)
#            if file['name'].endswith('.zip'):  # Если это zip
#                with zipfile.ZipFile(file_path, 'r') as zip_ref:
#                    zip_ref.extractall(output_folder)
#                os.remove(file_path)  # Удалить исходный zip

# Преобразование данных в DataFrame
def create_dataframes(output_folder='output'):
    dataframes = {}
    for root, _, files in os.walk(output_folder):
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(root, file)
                df = pd.read_excel(file_path)
                project_name = os.path.basename(root)
                week = os.path.splitext(file)[0]
                key = f"{project_name}_{week}"
                dataframes[key] = df
    return dataframes

# Основной скрипт
if __name__ == "__main__":
    service = authorize()
    process_files(service, FOLDER_ID)
#    dataframes = create_dataframes()
#    for key, df in dataframes.items():
#        print(f"Данные для {key}:")
#        print(df.head())

