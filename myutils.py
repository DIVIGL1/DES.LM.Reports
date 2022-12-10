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
        if os.path.isdir("Picts"):
            return os.path.join(os.path.abspath("."), "Picts", relative)
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
        (myconstants.PARAMETER_STR_LASTDAYOFMONHT, last_day_of_month2),
    ]

    parameter_data = [myconstants.PARAMETERS_FOR_GETTING_DATA_FOR_URL.copy()]
    parameter_str = parameter_data[0][myconstants.PARAMETER_STR_KEY_WITH_PERIOD]
    for one_parameter in replace_data:
        parameter_str = parameter_str.replace(one_parameter[0], one_parameter[1])

    parameter_data[0][myconstants.PARAMETER_STR_KEY_WITH_PERIOD] = parameter_str
    return(parameter_data)


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
            mainwindow.parent.report_automation_in_process
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
    from iCodes import get_des_lm_url, get_secret_code
    data = get_des_lm_url()
    if data["ret_code"] == -1:
        mainwindow.ui.LoadDataFromDESLM.setVisible(False)
        mainwindow.ui.LoadFromDELMAndCreateReport.setVisible(False)
        mainwindow.ui.GetUserCode.setText(f"Истёк срок валидности. Пользовательский код : [{get_secret_code()}]")

if __name__ == "__main__":
    print(get_data_using_url(month2=datetime.datetime.now().month))
