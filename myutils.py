import datetime
import os
import sys
import pickle
import shutil
import threading
import platform
import hashlib
import subprocess
import json
import requests
import socket
import struct
import base64
from cryptography.fernet import Fernet

import myconstants


def thread(my_func):
    # Запускает функцию в отдельным процессом.
    def wrapper(*args, **kwargs):
        my_thread = threading.Thread(target=my_func, args=args, kwargs=kwargs)
        my_thread.start()
    return wrapper


def iif(if_condition, true_ret_value, false_ret_value):
    if if_condition:
        return true_ret_value
    else:
        return false_ret_value


def save_param(param_name, param_value, filename="saved.pkl"):
    params_dict = get_all_params(filename="saved.pkl")
    params_dict[param_name] = param_value

    with open(filename, 'wb') as file_handle:
        pickle.dump(params_dict, file_handle)


def get_all_params(filename="saved.pkl"):
    if not os.path.isfile(filename):
        with open(filename, 'wb') as file_handle:
            pickle.dump({"<ДатаВремя создания>": datetime.datetime.now()}, file_handle)

    with open(filename, "rb") as file_handle:
        params_dict = pickle.load(file_handle)

    return params_dict


def load_param(param_name, default="", filename="saved.pkl"):
    params_dict = get_all_params(filename=filename)

    return params_dict.get(param_name, default)


def get_files_list(path2files="", files_starts="", files_ends=".xlsx", reverse=True):
    path2files = os.path.join(os.getcwd(), path2files)
    files_list = \
        [one_file[len(files_starts):][:-len(files_ends)]
            for one_file in os.listdir(path2files)
                if (one_file.lower().startswith(files_starts.lower()) and
                    one_file.lower().endswith(files_ends.lower()) and
                    one_file[0] != "~"
                    )
         ]

    files_list = sorted(files_list, reverse=reverse)
    return files_list


def rel_path(path):
    home_dir = get_home_dir()
    ret_value = os.path.relpath(path, home_dir)

    return ret_value


def get_home_dir():
    return os.path.abspath(os.curdir)


def get_download_dir():
    if os.name == 'nt':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
        return location
    else:
        return os.path.join(os.path.expanduser('~'), 'downloads')


def get_later_raw_file():
    path = get_download_dir()
    # Получим список имен всего содержимого папки
    # и превратим их в абсолютные пути. Но обработаем
    # только те, которые являются файлами Excel.
    dir_list = [os.path.join(path, x) for x in os.listdir(path) if x[-len(myconstants.EXCEL_FILES_ENDS):].lower() == myconstants.EXCEL_FILES_ENDS]

    if dir_list:
        # Если список не пустой, то создадим список из путей к файлам и дат их создания.
        date_list = [[x, os.path.getctime(x)] for x in dir_list]

        # Отсортируем список по дате создания в обратном порядке
        sort_date_list = sorted(date_list, key=lambda x: x[1], reverse=True)

        # Выведем первый элемент списка. Он и будет самым последним по дате
        return sort_date_list[0][0]
    else:
        return None


def open_download_dir():
    open_dir_in_explore(get_download_dir())


def open_user_files_dir():
    from mytablefuncs import get_parameter_value
    user_files_path = get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME)
    user_files_path = user_files_path.replace("/", "\\")
    open_dir_in_explore(user_files_path)


def open_raw_files_dir():
    from mytablefuncs import get_parameter_value
    user_files_path = get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)
    user_files_path = user_files_path.replace("/", "\\")
    open_dir_in_explore(user_files_path)


def open_dir_in_explore(dir_path):
    os.system(f"explorer.exe {dir_path}")


def test_create_dir(dir_path):
    # Проверка на существование и создание директории.
    if not os.path.isdir(dir_path):
        # Если её не существуем, то попробуем создать
        os.mkdir(dir_path)
        # Повторно проверим наличие директории
        if not os.path.isdir(dir_path):
            return False

    return True


def get_internet_time():
    address = ('pool.ntp.org', 123)
    msg = '\x1b' + '\0' * 47

    client = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    client.sendto(bytes(msg, encoding='utf-8'), address)
    msg, _ = client.recvfrom(1024)

    secs = struct.unpack("!12I", msg)[10] - 2208988800
    return datetime.datetime.fromtimestamp(secs)

@thread
def copy_file_as_drop_process(mainwindow, xls_files, create_report=False):
    from mytablefuncs import open_and_test_raw_struct, get_parameter_value
    # Установим флаг, который используется при проверке изменений на диске (FileSystemEventHandler),
    # а так же на основании него определяется какие элементы доступны в интерфейсе.
    mainwindow.parent.drag_and_prop_in_process = True
    mainwindow.ui.lock_unlock_interface_items()

    drug_and_drop_type = (
            mainwindow.ui.radioButtonDD1.isChecked() * 1 +
            mainwindow.ui.radioButtonDD2.isChecked() * 2 +
            mainwindow.ui.radioButtonDD3.isChecked() * 3 +
            mainwindow.ui.radioButtonDD4.isChecked() * 4
    )

    raw_section_path = get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)
    counter = 0

    mainwindow.add_text_to_log_box("Обрабатываем Excel файл" + ("ы:" if len(xls_files) > 1 else ":"))
    for file_num, one_file_path in enumerate(xls_files):
        if file_num > 0:
            mainwindow.add_text_to_log_box("")

        if len(xls_files) == 1:
            mainwindow.add_text_to_log_box(f"   {one_file_path}")
        else:
            mainwindow.add_text_to_log_box(f"   {file_num + 1}. {one_file_path}")
        this_file_name = os.path.basename(one_file_path)

        if drug_and_drop_type >= 2:
            # Проверим структуру файла:
            ret_value = open_and_test_raw_struct(one_file_path, short_text=True)
            if type(ret_value) == str:
                mainwindow.add_text_to_log_box("   Структура файла не соответствует требованиям.")
                mainwindow.add_text_to_log_box("   Копирование отклонено.")
                continue

        if drug_and_drop_type >= 3:
            # Определим новое имя файла исходя из его данных.
            # Сначала определим дату файла:
            file_dt = datetime.datetime.fromtimestamp(os.path.getmtime(one_file_path))
            creation_str = f"{file_dt:%Y-%m-%d %H-%M}"
            # Определим данные за какой период присутствуют:
            month_column = list(myconstants.RAW_DATA_COLUMNS.keys())[0]

            start_month = ret_value[month_column].min()
            report_year = int(start_month * 10000 - int(start_month) * 10000)
            start_month = int(start_month)
            end_month = int(ret_value[month_column].max())

            if start_month == end_month:
                data_in_file_period = f"{myconstants.MONTHS[end_month]} {report_year}"
            else:
                data_in_file_period = f"{myconstants.MONTHS[start_month]}-{myconstants.MONTHS[end_month]} {report_year}"
            new_filename = f"{creation_str}  ({data_in_file_period}).xlsx"
            if new_filename != this_file_name:
                mainwindow.add_text_to_log_box(f"   Имя файла меняется на {new_filename}.")

        else:
            new_filename = this_file_name

        raw_file_path = raw_section_path + "/" + new_filename
        try:
            shutil.copy(one_file_path, raw_file_path)
            select_filename = os.path.splitext(os.path.basename(raw_file_path))[0]
            counter += 1
            mainwindow.add_text_to_log_box("   Файл скопирован.")
        except (OSError, shutil.Error):
            mainwindow.add_text_to_log_box("   Копирование не удалось - возникли ошибки.")
            continue

        if drug_and_drop_type == 4:
            if this_file_name == new_filename:
                # Не надо переименовывать файл сам в себя.
                mainwindow.add_text_to_log_box(
                    f"   Исходный файл переименовывать не надо, так как он уже имеет нужное имя: {new_filename}.")
            else:
                new_src_file_path = os.path.join(os.path.dirname(one_file_path), new_filename)
                try:
                    os.rename(one_file_path, new_src_file_path)
                    mainwindow.add_text_to_log_box("   Исходный файл так же переименован.")
                except (OSError, shutil.Error):
                    mainwindow.add_text_to_log_box("   Переименование исходного файла не удалось.")

    if counter == 1:
        mainwindow.refresh_raw_files_list(select_filename)
    else:
        mainwindow.refresh_raw_files_list()

    if not mainwindow.parent.report_automation_in_process:
        mainwindow.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)

    mainwindow.parent.drag_and_prop_in_process = False

    and_create_report(mainwindow=mainwindow, create_report=create_report)
    return


def and_create_report(mainwindow, create_report=False):
    if create_report:
        mainwindow.set_status_bar_text("... начинаем формировать отчёт на основании скопированного файла", 3)
        mainwindow.parent.reporter.create_report(p_dont_clear_log_box=True)
    else:
        mainwindow.ui.lock_unlock_interface_items()


def is_admin():
    from mytablefuncs import get_parameter_value
    user_files_path = get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME)
    hash_string = "#" + platform.node() + "#" + os.environ.get('USERNAME') + "#"
    hash_string = hashlib.blake2s(hash_string.encode(encoding = "utf-8"), digest_size=5).hexdigest()
    hash_file = os.path.join(user_files_path, hash_string)

    return os.path.isfile(hash_file)


def open_file_in_application(file_name):
    subprocess.Popen(file_name, shell=True)


def get_resource_path(relative):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative)
    else:
        if os.path.isdir(myconstants.RES_FOLDER):
            return os.path.join(os.path.abspath("."), myconstants.RES_FOLDER, relative)
        else:
            return os.path.join(os.path.abspath("."), relative)


def get_des_lm_url_parameters(year=None, month1=1, month2=None):
    if year is None:
        year = datetime.datetime.now().year

    if month2 is None:
        month2 = month1
    month2 = max(month1, month2)

    last_day_of_month2 = datetime.date(year, month2, 1) + datetime.timedelta(days=32)
    last_day_of_month2 = datetime.date(last_day_of_month2.year, last_day_of_month2.month, 1) - datetime.timedelta(days=1)
    last_day_of_month2 = f"{last_day_of_month2.day}"

    month1 = f"{month1:02}"
    month2 = f"{month2:02}"

    year = f"{year:04}"

    replace_data = [
        (myconstants.PARAMETER_STR_YEAR, year),
        (myconstants.PARAMETER_STR_MONTH1, month1),
        (myconstants.PARAMETER_STR_MONTH2, month2),
        (myconstants.PARAMETER_STR_FIRST_REPORT_DAY, "01"),
        (myconstants.PARAMETER_STR_LAST_REPORT_DAY, last_day_of_month2),
    ]

    parameter_data = [myconstants.PARAMETERS_FOR_GETTING_DATA_FOR_URL.copy()]
    parameter_str = parameter_data[0][myconstants.PARAMETER_STR_KEY_WITH_PERIOD]
    for one_parameter in replace_data:
        parameter_str = parameter_str.replace(one_parameter[0], one_parameter[1])

    parameter_data[0][myconstants.PARAMETER_STR_KEY_WITH_PERIOD] = parameter_str
    return parameter_data


@thread
def get_data_using_url(mainwindow=None, year=None, month1=1, month2=None, create_report=None):
    if not mainwindow is None:
        mainwindow.parent.internet_downloading_in_process = True
        mainwindow.ui.lock_unlock_interface_items()

    from iCodes import get_des_lm_url
    data = get_des_lm_url()
    if data["ret_code"] != 1:
        # Либо код устарел, либо что-то с Интернетом:
        if not mainwindow is None:
            # Надо остановить процессы и вывести сообщение на экран
            mainwindow.parent.internet_downloading_in_process = False
            # mainwindow.parent.report_automation_in_process
            mainwindow.add_text_to_log_box(f"Не удалось получить доступ к DES.LM.")
            mainwindow.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
            mainwindow.ui.lock_unlock_interface_items()
    else:
        url = data["url"]

        data = get_des_lm_url_parameters(year=year, month1=month1, month2=month2)

        rs = requests.post(url, data=json.dumps(data), headers={"Content-Type": "application/json"})

        cdt = datetime.datetime.now()

        if month1 == month2:
            data_in_file_period = f"{myconstants.MONTHS[month2]} {year}"
        else:
            data_in_file_period = f"{myconstants.MONTHS[month1]}-{myconstants.MONTHS[month2]} {year}"

        only_filename = f"{cdt.year:04}-{cdt.month:02}-{cdt.day:02} {cdt.hour:02}-{cdt.minute:02}-{cdt.second:02}  DES.LM.Reports ({data_in_file_period})" + myconstants.EXCEL_FILES_ENDS
        filename = os.path.join(get_download_dir(), only_filename)

        with open(filename, "wb") as file:
            file.write(rs.content)

        if not mainwindow is None:
            mainwindow.add_text_to_log_box(f"Завершена загрузка данных из DES.LM через Интернет.")
            mainwindow.add_text_to_log_box(f"Файл: {only_filename} размещён в папке 'Загрузки'.")

            if not mainwindow.parent.report_automation_in_process:
                mainwindow.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)

            mainwindow.parent.internet_downloading_in_process = False
            mainwindow.ui.lock_unlock_interface_items()

            if mainwindow.parent.report_automation_in_process:
                copy_file_as_drop_process(mainwindow, [filename], create_report=create_report)


@thread
def test_access_key(mainwindow):
    mainwindow.ui.LoadDataFromDESLM.setVisible(False)
    mainwindow.ui.LoadFromDELMAndCreateReport.setVisible(False)
    from iCodes import get_des_lm_url, get_user_code
    data = get_des_lm_url()
    menu = True
    toolbar = False
    action = mainwindow.ui.GetUserCode
    if data["ret_code"] == 1:
        # Всё нормально.
        mainwindow.ui.LoadDataFromDESLM.setVisible(True)
        mainwindow.ui.LoadFromDELMAndCreateReport.setVisible(True)
        action.setVisible(False)
        return

    mainwindow.ui.GetUserCode.setText(f"Пользовательский код: [{get_user_code()}]")

    if data["ret_code"] == -1:
        # Истёк срок действия ключа
        action.setText(f"Истёк срок валидности. Пользовательский код : [{get_user_code()}]")
        action.setToolTip("Истёк срок валидности ключа доступа.")
        pict = "locked"
        mainwindow.ui.setup_one_action(action=action, pict=pict, menu=menu, toolbar=toolbar)
        return

    if data["ret_code"] == -2:
        # Получена не правильная ссылка
        action.setText(f"Проблема с доступом! Пользовательский код : [{get_user_code()}]")
        action.setToolTip("Проблема с доступом к данным\nв DES.LM через Интернет!")
        pict = "cancel"
        mainwindow.ui.setup_one_action(action=action, pict=pict, menu=menu, toolbar=toolbar)
        return

    if data["ret_code"] == -3:
        # Нет локального файла ключа
        action.setText(f"Пользовательский код: [{get_user_code()}]")
        action.setToolTip("Отсутствует настройка\nдля доступа к данным из DES.LM\nчерез Интернет.")
        pict = "key"
        mainwindow.ui.setup_one_action(action=action, pict=pict, menu=menu, toolbar=toolbar)
        return

    if data["ret_code"] == -4:
        # Нет доступа в Интернет
        action.setText(f"Нет доступа в Интернет. Пользовательский код : [{get_user_code()}]")
        action.setToolTip("Отсутствует доступ в Интернет,\nнеобходимый для получения\nданных из DES.LM.")
        pict = "cancel"
        mainwindow.ui.setup_one_action(action=action, pict=pict, menu=menu, toolbar=toolbar)
        return

    if data["ret_code"] == -5:
        # Не понятная проблема
        action.setToolTip("Возникли неопределённые проблемы\nс доступом в Интернет.")
        # action.setVisible(False)
        return

    if data["ret_code"] == -6:
        # Нет локального файла ключа
        action.setText(f"Пользовательский код: [{get_user_code()}]")
        action.setToolTip("Отсутствует настройка\nдля доступа к данным из DES.LM\nчерез Интернет\n(InvalidToken)")
        pict = "key"
        mainwindow.ui.setup_one_action(action=action, pict=pict, menu=menu, toolbar=toolbar)
        return

    if data["ret_code"] == -7:
        # Нет серверного файла ключа
        action.setText(f"Пользовательский код: [{get_user_code()}]")
        action.setToolTip("Отсутствует настройка\nдля доступа к данным из DES.LM\nчерез Интернет\n(на сервере нет ключа)")
        pict = "key"
        mainwindow.ui.setup_one_action(action=action, pict=pict, menu=menu, toolbar=toolbar)
        return

    if data["ret_code"] == -8:
        # Другая ошибка считывания серверного файла ключа
        action.setText(f"Пользовательский код: [{get_user_code()}]")
        action.setToolTip("Не удалось прочитать\nключ доступа для подключения\nк DES.LM через Интернет")
        pict = "key"
        mainwindow.ui.setup_one_action(action=action, pict=pict, menu=menu, toolbar=toolbar)
        return

    if data["ret_code"] == -9:
        # Не удалось прочитать время в Интернет, чтобы определить "срок годности" ключа
        action.setText(f"Пользовательский код: [{get_user_code()}]")
        action.setToolTip("Не удалось прочитать\nвремя с сервера Интернет")
        pict = "key"
        mainwindow.ui.setup_one_action(action=action, pict=pict, menu=menu, toolbar=toolbar)
        return


def get_common_crypter(ui):
    # Прочитаем ключ
    key_path = get_resource_path("common.key")
    if not os.path.isfile(key_path):
        if ui is not None:
            ui.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
            ui.add_text_to_log_box("Не удалось найти общий ключ шифрования.")
            ui.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
        return ""
    else:
        with open(get_resource_path("common.key"), 'rb') as file:
            key = file.read()
    return Fernet(key)


@thread
def test_internet_data_version(ui):
    ui.UpdateParametersFromInternet.setVisible(False)
    ui.UpdateReportsFromInternet.setVisible(False)
    ui.UpdateParameterEMails.setVisible(False)

    # Прочитаем ключ, если он не доступен, то и скачивать ничего нельзя
    crypter = get_common_crypter(ui)
    if not crypter:
        return

    url = "https://raw.githubusercontent.com/iCodes/AccessCodes/main/inform"
    rs = requests.get(url)

    if rs.ok:
        versions_info = json.loads(rs.content)
        # Проверим необходимость обновления программы
        if versions_info.get("common version", {"version": float("-inf")})["version"] > myconstants.COMMON_VERSION:
            ui.parent.setWindowTitle(f"DES.LM.Reporter ({myconstants.APP_VERSION}) - НЕОБХОДИМО ОБНОВИТЬ ПРОГРАММУ")

        # Проверим версию параметров
        params_on_internet_last_ver = load_param(myconstants.LAST_INTERNET_PARAMS_NAME, myconstants.LAST_INTERNET_PARAMS_VERSION)
        params_on_internet_curr_ver = versions_info.get("params", {"version": float("-inf")})["version"]

        if params_on_internet_curr_ver > params_on_internet_last_ver:
            ui.UpdateParametersFromInternet.setVisible(True)

        # Проверим версию отчётных форм
        reports_on_internet_last_ver = load_param(myconstants.LAST_INTERNET_REPORTS_NAME, myconstants.LAST_INTERNET_REPORTS_VERSION)
        reports_on_internet_curr_ver = versions_info.get("reports", {"version": float("-inf")})["version"]

        if reports_on_internet_curr_ver > reports_on_internet_last_ver:
            ui.UpdateReportsFromInternet.setVisible(True)

        # Проверим возможность обновлять почтовые адреса для данного пользователя
        if test_user_access_2_download_emails(ui):
            # Проверим версию списка почтовых адресов
            emails_on_internet_last_ver = load_param(myconstants.LAST_INTERNET_EMAILS_NAME, myconstants.LAST_INTERNET_EMAILS_VERSION)
            emails_on_internet_curr_ver = versions_info.get("emails", {"version": float("-inf")})["version"]

            if emails_on_internet_curr_ver > emails_on_internet_last_ver:
                ui.UpdateParameterEMails.setVisible(True)


def test_user_access_2_download_emails(ui):
    url_user_4_emails = "https://raw.githubusercontent.com/iCodes/AccessCodes/main/data4"
    rs = requests.get(url_user_4_emails)
    if rs.ok:
        # Прочитаем ключ
        crypter = get_common_crypter(ui)

        try:
            users_list = json.loads(crypter.decrypt(rs.content).decode())
        except:
            users_list = {}
        users_list = users_list.get("users", [])

        from iCodes import get_user_code
        return get_user_code() in users_list

    return False


@thread
def get_internet_data(ui, stype):
    url = "https://raw.githubusercontent.com/iCodes/AccessCodes/main/inform"
    rs = requests.get(url)

    if rs.ok:
        import pandas as pd
        from mytablefuncs import get_parameter_value
        # Прочитаем ключ
        crypter = get_common_crypter(ui)

        versions_info = json.loads(rs.content)
        if stype == "params":
            params_on_internet_last_ver = load_param(myconstants.LAST_INTERNET_PARAMS_NAME, myconstants.LAST_INTERNET_PARAMS_VERSION)
            params_on_internet_curr_ver = versions_info.get("params", {"version": float("inf")})["version"]

            if params_on_internet_curr_ver > params_on_internet_last_ver:
                url_data = "https://raw.githubusercontent.com/iCodes/AccessCodes/main/data1"
                ui.UpdateParametersFromInternet.setEnabled(False)
                ui.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
                ui.add_text_to_log_box("Начата загрузка параметров...")
                rs = requests.get(url_data)
                if rs.ok:
                    # Обработаем строку полученную из Интернета
                    from_internet = json.loads(crypter.decrypt(rs.content).decode())
                    files_location_save_to = get_parameter_value(myconstants.PARAMETERS_SECTION_NAME)

                    ui.change_last_log_box_text("Обновляем файлы параметров:")
                    for filename in from_internet.keys():
                        from_internet[filename] = base64.b64decode(from_internet[filename].encode())
                        if not os.path.isfile("_tmp_DUP.txt"):
                            if myconstants.VIRTUAL_FTE_FILE_NAME.lower().split("(")[0] in filename.lower():
                                # Если это файл с виртуальными FTE, то его надо скопировать в пользовательскую папку
                                with open(os.path.join(get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME), filename), 'wb') as hfile:
                                    hfile.write(from_internet[filename])
                                vff_year = filename.split("(")[1].split(")")[0]
                                ui.add_text_to_log_box(f"   Таблица с искусственными FTE за {vff_year} год")
                            else:
                                with open(os.path.join(files_location_save_to, filename), 'wb') as hfile:
                                    hfile.write(from_internet[filename])
                                ui.add_text_to_log_box(f"   {myconstants.PARAMETERS_ALL_TABLES[filename][0]}")
                        else:
                            if myconstants.VIRTUAL_FTE_FILE_NAME.lower().split("(")[0] in filename.lower():
                                # Если это файл с виртуальными FTE, то его надо скопировать в пользовательскую папку
                                vff_year = filename.split("(")[1].split(")")[0]
                                ui.add_text_to_log_box(f"   заблокировано: Таблица с искусственными FTE за {vff_year} год")
                            else:
                                ui.add_text_to_log_box(f"   заблокировано: {myconstants.PARAMETERS_ALL_TABLES[filename][0]}")
                    ui.UpdateParametersFromInternet.setVisible(False)
                    save_param(myconstants.LAST_INTERNET_PARAMS_NAME, params_on_internet_curr_ver)
                else:
                    ui.UpdateParametersFromInternet.setEnabled(True)
                    ui.change_last_log_box_text("... загрузка не удалась.")
                ui.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)

        elif stype == "reports":
            reports_on_internet_last_ver = load_param(myconstants.LAST_INTERNET_REPORTS_NAME, myconstants.LAST_INTERNET_REPORTS_VERSION)
            reports_on_internet_curr_ver = versions_info.get("reports", {"version": float("inf")})["version"]

            if reports_on_internet_curr_ver > reports_on_internet_last_ver:
                url_data = "https://raw.githubusercontent.com/iCodes/AccessCodes/main/data2"
                ui.UpdateReportsFromInternet.setEnabled(False)
                ui.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
                ui.add_text_to_log_box("Начата загрузка шаблонов отчётов...")
                rs = requests.get(url_data)
                if rs.ok:
                    # Обработаем строку полученную из Интернета
                    from_internet = json.loads(crypter.decrypt(rs.content).decode())
                    files_location_save_to = get_parameter_value(myconstants.REPORTS_SECTION_NAME)

                    ui.change_last_log_box_text("Обновляем шаблоны отчётов:")
                    for filename in from_internet.keys():
                        from_internet[filename] = base64.b64decode(from_internet[filename].encode())
                        if not os.path.isfile("_tmp_DUP.txt"):
                            # Обновляем только в случае если такой отчёт уже был
                            if os.path.isfile(os.path.join(files_location_save_to, filename)):
                                with open(os.path.join(files_location_save_to, filename), 'wb') as hfile:
                                    hfile.write(from_internet[filename])
                                ui.add_text_to_log_box(f"   {filename}")
                        else:
                            ui.add_text_to_log_box(f"   заблокировано: {filename}")
                    ui.UpdateReportsFromInternet.setVisible(False)
                    save_param(myconstants.LAST_INTERNET_REPORTS_NAME, reports_on_internet_curr_ver)
                else:
                    ui.UpdateReportsFromInternet.setEnabled(True)
                    ui.change_last_log_box_text("... загрузка не удалась.")
                ui.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)

        elif stype == "emails":
            if not test_user_access_2_download_emails(ui):
                ui.UpdateParameterEMails.setVisible(False)
                return

            emails_on_internet_last_ver = load_param(myconstants.LAST_INTERNET_EMAILS_NAME, myconstants.LAST_INTERNET_EMAILS_VERSION)
            emails_on_internet_curr_ver = versions_info.get("emails", {"version": float("inf")})["version"]

            if emails_on_internet_curr_ver > emails_on_internet_last_ver:
                url_data = "https://raw.githubusercontent.com/iCodes/AccessCodes/main/data3"
                ui.UpdateParameterEMails.setEnabled(False)
                ui.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
                ui.add_text_to_log_box("Начата загрузка электронных адресов...")
                rs = requests.get(url_data)
                if rs.ok:
                    # Обработаем строку полученную из Интернета
                    from_internet = json.loads(crypter.decrypt(rs.content).decode())
                    files_location_save_to = get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME)
                    filename = myconstants.EMAILS_TABLE
                    from_internet[filename] = base64.b64decode(from_internet[filename].encode())
                    ui.change_last_log_box_text("Обновлён файл:")

                    if not os.path.isfile(os.path.join(files_location_save_to, filename)):
                        filename = "excluded__" + myconstants.EMAILS_TABLE

                    if not os.path.isfile("_tmp_DUP.txt"):
                        with open(os.path.join(files_location_save_to, filename), 'wb') as hfile:
                            hfile.write(from_internet[myconstants.EMAILS_TABLE])
                        ui.add_text_to_log_box(f"   {myconstants.PARAMETERS_ALL_TABLES[myconstants.EMAILS_TABLE][0]}")
                    else:
                        ui.add_text_to_log_box(
                            f"   заблокировано: {myconstants.PARAMETERS_ALL_TABLES[myconstants.EMAILS_TABLE][0]}")
                    ui.UpdateParameterEMails.setVisible(False)
                    save_param(myconstants.LAST_INTERNET_EMAILS_NAME, emails_on_internet_curr_ver)
                else:
                    ui.UpdateReportsFromInternet.setEnabled(True)
                    ui.change_last_log_box_text("... загрузка не удалась.")
                ui.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
    else:
        ui.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
        if stype == "params":
            ui.add_text_to_log_box("Не удалось обновить параметры.")
        if stype == "reports":
            ui.add_text_to_log_box("Не удалось обновить шаблоны отчётов.")
        if stype == "emails":
            ui.add_text_to_log_box("Не удалось обновить адреса электронной почты.")
        ui.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)


if __name__ == "__main__":
    print(get_data_using_url(month2=datetime.datetime.now().month))
