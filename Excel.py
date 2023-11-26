import pandas as pd
import pandas.core.frame
import warnings
from Yandex import *
from CONFIG import MyError

warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.simplefilter(action='ignore', category=UserWarning)


def _kolvo_lab(DF: pandas.core.frame.DataFrame) -> int:
    """
        :param DF: DataFrame с данными из эксель таблицы
        :return: Возвращает число, являющееся кол-вом лабораторных работ, нужно для того, чтобы не производить лишние вычисления
    """
    try:
        return abs(DF.columns.get_loc('Points') - DF.columns.get_loc('GitHub')) - 1
    except:
        raise MyError("Ошибка в расчёте кол-ва лабораторных работ")


def _set_formula(DF: pandas.core.frame.DataFrame):
    """
        :param DF: DataFrame с данными из эксель таблицы
        :return: Заносит формулы в эксель таблицу
    """
    _set_sum_formula(DF=DF)
    _set_if_formula(DF=DF)


def _set_if_formula(DF: pandas.core.frame.DataFrame):
    """
        :param DF: DataFrame с данными из эксель таблицы
        :return: Заносит формулы условий
    """
    list_of_let = ["D", "E", "F", "G", "H", "I", "J", "K"]
    list_of_let_new = list_of_let[:_kolvo_lab(DF=DF)]
    try:
        for col in range(1, _kolvo_lab(DF=DF) + 1):
            for row in range(0, DF.shape[0]):
                DF.loc[row, "Подсчёт " + str(col)] = '=IF(' + str(list_of_let_new[col - 1]) + str(
                    row + 2) + '="Принято",12,0)'
    except:
        raise MyError("Ошибка в занесении формул условий")


def _set_sum_formula(DF: pandas.core.frame.DataFrame):
    """
        :param DF: DataFrame с данными из эксель таблицы
        :return: Заносит сумирующие формулы
    """
    try:
        for i in range(0, DF.shape[0]):
            DF.loc[i, "Points"] = "=SUM(M" + str(i + 2) + ":" + "T" + str(i + 2) + ")"
    except:
        raise MyError("Ошибка в занесении сумирующих формул")


def _read_excel_bd(DATABASE_NAME: str, GROUP: str):
    """
        :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
        :param GROUP: имя группы в формате "ПИН-221"
        :return: Возвращает DataFrame с данными из эксель таблицы
    """
    try:
        return pd.read_excel(DATABASE_NAME, sheet_name=GROUP.upper(), engine="openpyxl")
    except FileNotFoundError:
        raise MyError("Файл не найден")
    except:
        raise MyError("Ошибка при чтении файла")


def _save_excel_bd(DF: pandas.core.frame.DataFrame, DATABASE_NAME: str, GROUP: str):
    """
        :param DF: DataFrame с данными из эксель таблицы
        :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
        :param GROUP: имя группы в формате "ПИН-221"
        :return: Сохраняет данные в эксель таблицу
    """
    excel_header = []
    for col in range(0, len(list(DF.columns))):
        if list(DF.columns)[col].split()[0] == "Unnamed:":
            excel_header.append("")
        else:
            excel_header.append(DF.columns[col])
    try:
        _set_formula(DF=DF)
    except:
        raise MyError("Не удалось занести формулы")
    try:
        with pd.ExcelWriter(DATABASE_NAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            DF.to_excel(writer, index=False, sheet_name=GROUP.upper(), header=excel_header)
    except PermissionError:
        raise MyError(f"Закройте локальный файл {DATABASE_NAME}")
    except FileNotFoundError:
        raise MyError(f"Файл {DATABASE_NAME} не найден")
    except:
        raise MyError("Ошибка при сохранении")


def _find_student(DATABASE_NAME: str, GROUP: str, NAME: str) -> bool:
    """
    :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
    :param GROUP: имя группы в формате "ПИН-221"
    :param NAME: имя студента в формате "Фролов Григорий"
    :return: True, если студент найден; False, если студент не найден
    """
    df = _read_excel_bd(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
    filtered_df = df.loc[df["Name"] == NAME.title()]
    try:
        if filtered_df.empty:
            print(f"Студент {NAME.title()} не найден")
            return False
        else:
            print(f"Студент {NAME.title()} найден")
            return True
    except:
        raise MyError("Ошибка при поиске студента")


def authorization_student(DATABASE_NAME: str, GROUP: str, NAME: str) -> bool:
    """
        :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
        :param GROUP: имя группы в формате "ПИН-221"
        :param NAME: имя студента в формате "Фролов Григорий"
        :return: True, если студент найден; False, если студент не найден
    """
    download_database(DATABASE_NAME=DATABASE_NAME)
    if _find_student(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME):
        delete_file(DATABASE_NAME=DATABASE_NAME)
        return True
    else:
        delete_file(DATABASE_NAME=DATABASE_NAME)
        return False


def change_github(DATABASE_NAME: str, GROUP: str, NAME: str, NEW_LINK: str) -> bool:
    """
    :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
    :param GROUP: имя группы в формате "ПИН-221"
    :param NAME: имя студента в формате "Фролов Григорий"
    :param NEW_LINK: новая ссылка на GitHub студента в формате "https://github.com/"
    :return: True, если GitHub студента изменён, или не нуждается в
    изменении; False, если в ссылке на GitHub есть ошибки/опечатки, или студент не найден
    """
    download_database(DATABASE_NAME=DATABASE_NAME)
    df = _read_excel_bd(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
    if not _find_student(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME):
        delete_file(DATABASE_NAME=DATABASE_NAME)
        return False
    else:
        if NEW_LINK.split("/")[0] == "https:" and NEW_LINK.split("/")[2] == "github.com":
            try:
                OLD_LINK = df.loc[(df["Name"] == NAME.title()), "GitHub"].values
                if OLD_LINK != NEW_LINK:
                    df.loc[(df["Name"] == NAME.title()), "GitHub"] = NEW_LINK
                    _save_excel_bd(DF=df, DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
                    delete_database(DATABASE_NAME=DATABASE_NAME)
                    upload_database(DATABASE_NAME=DATABASE_NAME)
                    delete_file(DATABASE_NAME=DATABASE_NAME)
                    print(f"GitHub студента {NAME.title()} изменён")
                    return True
                else:
                    delete_file(DATABASE_NAME=DATABASE_NAME)
                    print(f"GitHub студента {NAME.title()} не нуждается в изменении")
                    return True
            except:
                delete_file(DATABASE_NAME=DATABASE_NAME)
                raise MyError("Ошибка при замене GitHub")
        else:
            delete_file(DATABASE_NAME=DATABASE_NAME)
            print(f"Ссылка на GitHub студента {NAME.title()} указана неверно")
            return False


def _show_me_my_points(DATABASE_NAME: str, GROUP: str, NAME: str):
    """
    :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
    :param GROUP: имя группы в формате "ПИН-221"
    :param NAME: имя студента в формате "Фролов Григорий"
    :return: Возвращает кол-во баллов студента; False, если студент не найден
    """
    df = _read_excel_bd(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
    if not _find_student(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME):
        return False
    else:
        try:
            student = df[df["Name"] == NAME.title()]
            points = student["Points"].values[0]
            print(f"Баллы студента {NAME.title()}: {points}")
            return points
        except:
            raise MyError("Ошибка при отображении баллов")


def set_status_ready_for_inspection(DATABASE_NAME: str, GROUP: str, NAME: str, LAB_WORK: str) -> bool:
    """
    :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
    :param GROUP: имя группы в формате "ПИН-221"
    :param NAME: имя студента в формате "Фролов Григорий"
    :param LAB_WORK: название лабораторной работы в формате "ЛР1"
    :return: True, если для работы {LAB_WORK} установлен статус "Готово к проверке", или работа уже принята; False, если студент не найден
    """
    download_database(DATABASE_NAME=DATABASE_NAME)
    df = _read_excel_bd(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
    if not _find_student(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME):
        delete_file(DATABASE_NAME=DATABASE_NAME)
        return False
    else:
        new_status = "Готово к проверке"
        try:
            student = df[df["Name"] == NAME.title()]
            if student[LAB_WORK].values[0] != "Принято" and student[LAB_WORK].values[0] != "принято" and student[LAB_WORK].values[0] != "прин":
                df.loc[(df["Name"] == NAME.title()), LAB_WORK] = new_status
                _save_excel_bd(DF=df, DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
                delete_database(DATABASE_NAME=DATABASE_NAME)
                upload_database(DATABASE_NAME=DATABASE_NAME)
                delete_file(DATABASE_NAME=DATABASE_NAME)
                print(f"Для работы {LAB_WORK}, студента {NAME.title()}, установлен статус {new_status}")
                return True
            else:
                delete_file(DATABASE_NAME=DATABASE_NAME)
                print("Работа уже принята")
                return True
        except:
            delete_file(DATABASE_NAME=DATABASE_NAME)
            raise MyError(f"Ошибка при смене статуса на {new_status}")


def set_telegram_id(DATABASE_NAME: str, GROUP: str, NAME: str, NEW_TELEGRAM_ID: str) -> bool:
    """
    :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
    :param GROUP: имя группы в формате "ПИН-221"
    :param NAME: имя студента в формате "Фролов Григорий"
    :param NEW_TELEGRAM_ID: новый Telegram ID студента
    :return: True, если Telegram ID изменён, или не нуждается в изменении; False, если студент не найден
    """
    download_database(DATABASE_NAME=DATABASE_NAME)
    df = _read_excel_bd(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
    if not _find_student(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME):
        delete_file(DATABASE_NAME=DATABASE_NAME)
        return False
    else:
        try:
            OLD_TELEGRAM_ID = df.loc[(df["Name"] == NAME.title()), "Telegram ID"].values[0]
            if OLD_TELEGRAM_ID != NEW_TELEGRAM_ID:
                df.loc[(df["Name"] == NAME.title()), "Telegram ID"] = str(NEW_TELEGRAM_ID)
                _save_excel_bd(DF=df, DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
                delete_database(DATABASE_NAME=DATABASE_NAME)
                upload_database(DATABASE_NAME=DATABASE_NAME)
                delete_file(DATABASE_NAME=DATABASE_NAME)
                print(f"Telegram ID студента {NAME.title()} изменён")
                return True
            else:
                delete_file(DATABASE_NAME=DATABASE_NAME)
                print(f"Telegram ID студента {NAME.title()} не нуждается в изменении")
                return True
        except:
            delete_file(DATABASE_NAME=DATABASE_NAME)
            raise MyError("Ошибка при усатновке/смене Telegram ID")


def check_status(DATABASE_NAME: str, GROUP: str, NAME: str):
    """
    :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
    :param GROUP: имя группы в формате "ПИН-221"
    :param NAME: имя студента в формате "Фролов Григорий"
    :return: Возвращает словарь со статусами лабораторных работ и общим количеством баллов; False, если студент не найден
    """
    download_database(DATABASE_NAME=DATABASE_NAME)
    df = _read_excel_bd(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
    if not _find_student(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME):
        delete_file(DATABASE_NAME=DATABASE_NAME)
        return False
    else:
        try:
            student = df[df["Name"] == NAME.title()]
            status = {}
            for i in range(0, _kolvo_lab(DF=df)):
                status[f"ЛР{i + 1}"] = student[f"ЛР{i + 1}"].values[0]
            count = _show_me_my_points(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME)
            if count != "nan":
                status["Баллы"] = count
            else:
                status["Баллы"] = "-"
            delete_file(DATABASE_NAME=DATABASE_NAME)
            print(f"Статус работ студента {NAME.title()}: {status}")
            return status
        except:
            delete_file(DATABASE_NAME=DATABASE_NAME)
            raise MyError("Ошибка при отображении статуса работы")


def find_by_telegram_id(DATABASE_NAME: str, TELEGRAM_ID: int):
    """
    :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
    :param TELEGRAM_ID: Telegram ID студента
    :return: Возвращает список с именем студента(0) и его грппой(1); False, если студент не найден
    """
    download_database(DATABASE_NAME=DATABASE_NAME)
    df = pd.read_excel(DATABASE_NAME, sheet_name=None)
    results = []
    for sheet_name, data in df.items():
        if "Name" in data.columns and "Telegram ID" in data.columns:
            filtered_df = data[data["Telegram ID"] == TELEGRAM_ID]
            if not filtered_df.empty:
                student_name = filtered_df.iloc[0]["Name"]
                results.append(student_name)
                results.append(sheet_name)
    if results:
        print(f"Студент {results[0]} найден в группе {results[1]}")
        delete_file(DATABASE_NAME=DATABASE_NAME)
        return results
    else:
        print("Студент не найден по заданному Telegram ID")
        delete_file(DATABASE_NAME=DATABASE_NAME)
        return False
