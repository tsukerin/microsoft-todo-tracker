import json
import pymstodo
import webbrowser
import requests
from msal import ConfidentialClientApplication
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
import time
import re
from dotenv import dotenv_values
import threading
from tabulate import tabulate
from colorama import Fore, Style
import pandas as pd
import datetime
import os

options = Options()
options.add_argument('--headless')
driver = webdriver.Firefox(options=options)

def loading_animation():
    print("Loading...", end="\r")
    time.sleep(0.5)

def find_authorization_code(app, scopes):

    loading_thread = threading.Thread(target=loading_animation)
    loading_thread.start()

    auth_url = app.get_authorization_request_url(scopes)
    driver.get(auth_url)

    email_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "i0116"))
    )
    email_input.send_keys(os.getenv('EMAIL'))
    email_input.send_keys(Keys.ENTER)
    time.sleep(2)

    password_input = driver.find_element(by='id', value='i0118')
    password_input.send_keys(os.getenv('EMAIL_PASSWORD'))
    password_input.send_keys(Keys.ENTER)
    time.sleep(3)

    driver.find_element(by='id', value='declineButton').click()
    time.sleep(3)

    current_url = driver.current_url
    match = re.search(r'code=([^&]+)', current_url)
    loading_thread.join()
    driver.quit()

    return match.group(1)

def create_confidential_client(client_id, client_secret, authority):
    return ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority
    )

def get_access_token(app, scopes, token):
    token_response = app.acquire_token_by_authorization_code(
        code=token,
        scopes=scopes
    )
    return token_response.get('access_token')

def get_todo_list(access_token):
    headers = {'Authorization': f'Bearer {access_token}'}
    todoTaskListId = 'AQMkADAwATNiZmYAZC0wODc1LWU2MDMtMDACLTAwCgAuAAADwsvZ7soLyEKfC3Ju9Ue4YAEANRCSvvhHdkO2d75WGfMsVQAAAgESAAAA'
    response = requests.get(f'https://graph.microsoft.com/v1.0/me/todo/lists/{todoTaskListId}/tasks', headers=headers)
    if response.status_code == 200:
        with open('todo_list.json', 'w', encoding='utf-8') as json_file:
            json.dump(response.json().get('value', []), json_file, indent=2, ensure_ascii=False)

        print(f"Задачи успешно получены!")
        return response.json().get('value', [])
    else:
        print(f"Ошибка при получении задач: {response.status_code}")

def sorted_data(headers, data):
    df = pd.DataFrame(data, columns=headers)

    df['Крайний срок выполенения'] = pd.to_datetime(df['Крайний срок выполенения'])
    df = df[(df['Крайний срок выполенения'].dt.date <= datetime.datetime.now().date()) & 
            (df['Время выполнения'].isin([str(datetime.datetime.now().date()), 'Отсутствует']))]
    
    return df.drop_duplicates()

def count_tasks(df):
    return {
        'completed': df[df['Статус выполнения'] == 'Выполнено']['Задача'].count(),
        'uncompleted': df[df['Статус выполнения'] == 'Не выполнено']['Задача'].count()
        }

def out(api):
    headers = ['Задача', 'Статус выполнения', 'Крайний срок выполенения', 'Время выполнения']

    mapping = {
        'completed': 'Выполнено',
        'notStarted': 'Не выполнено'
    }

    data = [
            [
                task["title"], 
                mapping.get(task["status"], task["status"]), 
                task['recurrence']['range']['startDate'] if 'recurrence' in task else task['dueDateTime']['dateTime'][:10],
                task['completedDateTime']['dateTime'][:10] if 'completedDateTime' in task else 'Отсутствует'
            ] 
        for task in api
        ]
    
    print(tabulate(sorted_data(headers, data), headers='keys', tablefmt='grid', showindex=False))
    print(f'{Fore.GREEN}Выполнено:{Style.RESET_ALL} {count_tasks(sorted_data(headers, data))['completed']} задачи')
    print(f'{Fore.RED}Не выполнено:{Style.RESET_ALL} {count_tasks(sorted_data(headers, data))['uncompleted']} задачи')


def main():
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')
    authority = 'https://login.microsoftonline.com/consumers/'
    scopes = ['Tasks.ReadWrite']

    app = create_confidential_client(client_id, client_secret, authority)
    token = find_authorization_code(app, scopes)
    access_token = get_access_token(app, scopes, token)

    if not access_token:
        print('access token не найден')
        return
    
    todo_client = get_todo_list(access_token)
    
    out(todo_client)

if __name__ == '__main__':
    main()