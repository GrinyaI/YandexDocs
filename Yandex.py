import requests
import os
from CONFIG import *


def download_database(DATABASE_NAME: str) -> bool:
    """
    Функция скачивает базу данных с Яндекс Диска
    :param DATABASE_NAME: Имя скачиваемой базы данных
    :return: True, при успешном скачивании файла с диска; False, при ошибке и выводит в консоль код ошибки
    """
    try:
        url = "https://cloud-api.yandex.net/v1/disk/resources/download"

        params = {'path': "/" + DATABASE_NAME}

        response = requests.get(url, params=params, headers=HEADERS)

        download_url = response.json()['href']
        download_response = requests.get(download_url, params=params, headers=HEADERS)
        print("Статус скачивания: " + str(download_response.status_code))
        with open(DATABASE_NAME, 'wb') as file:
            file.write(download_response.content)
        return True
    except requests.exceptions.HTTPError as error:
        print(f'Ошибка HTTP: {error}')
        return False
    except requests.exceptions.RequestException as error:
        print(f'Ошибка сети: {error}')
        return False


def upload_database(DATABASE_NAME: str) -> bool:
    """
    Функция загружает базу данных на Яндекс Диск
    :param DATABASE_NAME: Имя загружаемой базы данных (обязательно должна лежать в той же папке что и код)
    :return: True, при успешной загрузке файла на диск; False, при ошибке и выводит в консоль код ошибки, или если база
    данных редактируется преподавателем
    """
    try:
        url = "https://cloud-api.yandex.net/v1/disk/resources/upload"

        params = {'path': "/" + DATABASE_NAME}

        response = requests.get(url, params=params, headers=HEADERS)
        if response.status_code == 423:
            print("База данных редактируется преподавателем")
            return False
        upload_url = response.json()['href']

        with open(DATABASE_NAME, 'rb') as file:
            response = requests.put(upload_url, data=file, headers=HEADERS)
            print("Статус загрузки: " + str(response.status_code))
        return True
    except requests.exceptions.HTTPError as error:
        print(f'Ошибка HTTP: {error}')
        return False
    except requests.exceptions.RequestException as error:
        print(f'Ошибка сети: {error}')
        return False


def delete_database(DATABASE_NAME: str) -> bool:
    """
    Функция удаляет базу данных с Яндекс Диска
    :param DATABASE_NAME: Имя удаляемой базы данных
    :return: True, при успешном удалении файла с диска; False, при ошибке и выводит в консоль код ошибки, или если база
    данных редактируется преподавателем
    """
    try:
        url = "https://cloud-api.yandex.net/v1/disk/resources?force_async=true&permanently=true"

        params = {'path': "/" + DATABASE_NAME}

        response = requests.delete(url, params=params, headers=HEADERS)
        if response.status_code == 423:
            print("База данных редактируется преподавателем")
            return False
        print("Статус удаления: " + str(response.status_code))
        return True
    except requests.exceptions.HTTPError as error:
        print(f'Ошибка HTTP: {error}')
        return False
    except requests.exceptions.RequestException as error:
        print(f'Ошибка сети: {error}')
        return False


def delete_file(DATABASE_NAME: str) -> bool:
    """
    Функция удаляет временную базу данных необходимую для работы кода
    :param DATABASE_NAME: Имя удаляемой базы данных (обязательно должна лежать в той же папке что и код)
    :return: True, при успешном удалении локального файла; False, при ошибке и выводит в консоль код ошибки
    """
    try:
        os.remove(DATABASE_NAME)
        print("Файл успешно удален")
        return True
    except FileNotFoundError:
        print("Файл не найден")
        return False
    except PermissionError:
        print(f"Закройте локальный файл {DATABASE_NAME}")
        return False
