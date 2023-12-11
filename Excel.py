import pandas as pd
import pandas.core.frame
import warnings
from Yandex import *
from CONFIG import MyError

nest_asyncio.apply()

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
        set_points = 60 / _kolvo_lab(DF=DF)
        for col in range(1, _kolvo_lab(DF=DF) + 1):
            for row in range(0, DF.shape[0]):
                DF.loc[row, "Подсчёт " + str(col)] = '=IF(OR(' + str(list_of_let_new[col - 1]) + str(
                    row + 2) + '="Принято",' + str(list_of_let_new[col - 1]) + str(
                    row + 2) + '="принято",' + str(list_of_let_new[col - 1]) + str(
                    row + 2) + '="прин"' + f'),{int(set_points)},0)'
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
    except ValueError:
        raise MyError("Группа не найдена")
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
    try:
        df = _read_excel_bd(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
    except:
        return False
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


async def kolvo_lab(DATABASE_NAME: str, GROUP: str) -> int:
    var = await download_database(DATABASE_NAME=DATABASE_NAME)
    if not var:
        print("База данных не найдена")
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return False
    df = _read_excel_bd(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
    kolvo = _kolvo_lab(DF=df)
    await delete_file(DATABASE_NAME=DATABASE_NAME)
    return kolvo


async def authorization_student(DATABASE_NAME: str, GROUP: str, NAME: str) -> bool:
    """
        :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
        :param GROUP: имя группы в формате "ПИН-221"
        :param NAME: имя студента в формате "Фролов Григорий"
        :return: True, если студент найден; False, если студент не найден
    """
    var = await download_database(DATABASE_NAME=DATABASE_NAME)
    if not var:
        print("База данных не найдена")
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return False
    if _find_student(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME):
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return True
    else:
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return False


async def change_github(DATABASE_NAME: str, GROUP: str, NAME: str, NEW_LINK: str) -> str:
    """
    :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
    :param GROUP: имя группы в формате "ПИН-221"
    :param NAME: имя студента в формате "Фролов Григорий"
    :param NEW_LINK: новая ссылка на GitHub студента в формате "https://github.com/"
    :return: True, если GitHub студента изменён, или не нуждается в
    изменении; False, если в ссылке на GitHub есть ошибки/опечатки, или студент не найден
    """
    await download_database(DATABASE_NAME=DATABASE_NAME)
    try:
        df = _read_excel_bd(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
    except:
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return "Возможен ввод некорректных данных, попробуйте снова"
    if not _find_student(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME):
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return "Студент не найден"
    else:
        if NEW_LINK.split("/")[0] == "https:" and NEW_LINK.split("/")[2] == "github.com":
            try:
                OLD_LINK = df.loc[(df["Name"] == NAME.title()), "GitHub"].values
                if OLD_LINK != NEW_LINK:
                    df.loc[(df["Name"] == NAME.title()), "GitHub"] = NEW_LINK
                    _save_excel_bd(DF=df, DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
                    var = await delete_database(DATABASE_NAME=DATABASE_NAME)
                    if not var:
                        print("Не удалось удалить базу данных")
                        await delete_file(DATABASE_NAME=DATABASE_NAME)
                        return "Данные редактируются преподавателем, попробуйте позже"
                    await upload_database(DATABASE_NAME=DATABASE_NAME)
                    await delete_file(DATABASE_NAME=DATABASE_NAME)
                    return f"GitHub студента {NAME.title()} изменён"
                else:
                    await delete_file(DATABASE_NAME=DATABASE_NAME)
                    return f"GitHub студента {NAME.title()} не нуждается в изменении"
            except:
                await delete_file(DATABASE_NAME=DATABASE_NAME)
                print("Ошибка при замене github")
                return "Возможен ввод некорректных данных или данные редактируются преподавателем, попробуйте позже"
        else:
            await delete_file(DATABASE_NAME=DATABASE_NAME)
            return f"Ссылка на GitHub студента {NAME.title()} указана неверно"


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
            points = 0
            student = df[df["Name"] == NAME.title()]
            set_points = 60 / _kolvo_lab(DF=df)
            for i in range(1, _kolvo_lab(DF=df) + 1):
                if student[f"ЛР{i}"].values[0] == "Принято" or student[f"ЛР{i}"].values[0] == "принято" or \
                        student[f"ЛР{i}"].values[0] == "прин":
                    points += int(set_points)
            return points
        except:
            raise MyError("Ошибка при отображении баллов")


async def set_status_ready_for_inspection(DATABASE_NAME: str, GROUP: str, NAME: str, LAB_WORK: str) -> str:
    """
    :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
    :param GROUP: имя группы в формате "ПИН-221"
    :param NAME: имя студента в формате "Фролов Григорий"
    :param LAB_WORK: название лабораторной работы в формате "ЛР1"
    :return: True, если для работы {LAB_WORK} установлен статус "Готово к проверке", или работа уже принята; False, если студент не найден
    """
    var = await download_database(DATABASE_NAME=DATABASE_NAME)
    if not var:
        print("База данных не найдена")
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return "Данные редактируются, попробуйте позже"
    df = _read_excel_bd(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
    if not _find_student(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME):
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return "Студент не найден"
    else:
        new_status = "Готово к проверке"
        student = df[df["Name"] == NAME.title()]
        if student[LAB_WORK.upper()].values[0] != "Принято" and student[LAB_WORK.upper()].values[0] != "принято" and \
                student[LAB_WORK.upper()].values[0] != "прин":
            df.loc[(df["Name"] == NAME.title()), LAB_WORK.upper()] = new_status
            _save_excel_bd(DF=df, DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
            var = await delete_database(DATABASE_NAME=DATABASE_NAME)
            if not var:
                print("Не удалось удалить базу данных")
                await delete_file(DATABASE_NAME=DATABASE_NAME)
                return "Данные редактируются преподавателем, попробуйте позже"
            await upload_database(DATABASE_NAME=DATABASE_NAME)
            await delete_file(DATABASE_NAME=DATABASE_NAME)
            return f"Для работы {LAB_WORK.upper()}, студента {NAME.title()}, установлен статус {new_status}"
        else:
            await delete_file(DATABASE_NAME=DATABASE_NAME)
            return "Работа уже принята"


async def set_telegram_id(DATABASE_NAME: str, GROUP: str, NAME: str, NEW_TELEGRAM_ID: int) -> bool:
    """
    :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
    :param GROUP: имя группы в формате "ПИН-221"
    :param NAME: имя студента в формате "Фролов Григорий"
    :param NEW_TELEGRAM_ID: новый Telegram ID студента
    :return: True, если Telegram ID изменён, или не нуждается в изменении; False, если студент не найден
    """
    var = await download_database(DATABASE_NAME=DATABASE_NAME)
    if not var:
        print("База данных не найдена")
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return False
    try:
        df = _read_excel_bd(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
    except:
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return False
    if not _find_student(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME):
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return False
    else:
        try:
            try:
                OLD_TELEGRAM_ID = df.loc[(df["Name"] == NAME.title()), "Telegram ID"].values[0]
            except:
                print("Проблема со старым Telegram ID")
                await delete_file(DATABASE_NAME=DATABASE_NAME)
                return False
            if OLD_TELEGRAM_ID != NEW_TELEGRAM_ID:
                df.loc[(df["Name"] == NAME.title()), "Telegram ID"] = NEW_TELEGRAM_ID
                _save_excel_bd(DF=df, DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
                var = await delete_database(DATABASE_NAME=DATABASE_NAME)
                if not var:
                    print("Не удалось удалить базу данных")
                    await delete_file(DATABASE_NAME=DATABASE_NAME)
                    return False
                await upload_database(DATABASE_NAME=DATABASE_NAME)
                await delete_file(DATABASE_NAME=DATABASE_NAME)
                # print(f"Telegram ID студента {NAME.title()} изменён")
                return True
            else:
                await delete_file(DATABASE_NAME=DATABASE_NAME)
                # print(f"Telegram ID студента {NAME.title()} не нуждается в изменении")
                return True
        except:
            await delete_file(DATABASE_NAME=DATABASE_NAME)
            raise MyError("Ошибка при усатновке/смене Telegram ID")


async def check_status(DATABASE_NAME: str, GROUP: str, NAME: str):
    """
    :param DATABASE_NAME: имя базы данных в формате "ОПД.xlsx"
    :param GROUP: имя группы в формате "ПИН-221"
    :param NAME: имя студента в формате "Фролов Григорий"
    :return: Возвращает словарь со статусами лабораторных работ и общим количеством баллов; False, если студент не найден
    """
    await download_database(DATABASE_NAME=DATABASE_NAME)
    try:
        df = _read_excel_bd(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP)
    except:
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return "Возможен ввод некорректных данных, попробуйте снова"
    if not _find_student(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME):
        await delete_file(DATABASE_NAME=DATABASE_NAME)
        return "Студент не найден"
    else:
        try:
            student = df[df["Name"] == NAME.title()]
            status = {}
            text = "Всего работ:\n"
            for i in range(1, _kolvo_lab(DF=df) + 1):
                status[f"ЛР{i}"] = student[f"ЛР{i}"].values[0]
                text += "ЛР" + str(i) + " " + str(student[f"ЛР{i}"].values[0]) + "\n"
            text += "Всего баллов: " + str(_show_me_my_points(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME))
            status["Баллы"] = _show_me_my_points(DATABASE_NAME=DATABASE_NAME, GROUP=GROUP, NAME=NAME)
            await delete_file(DATABASE_NAME=DATABASE_NAME)
            return text
        except:
            await delete_file(DATABASE_NAME=DATABASE_NAME)
            raise MyError("Ошибка при отображении статуса работы")
