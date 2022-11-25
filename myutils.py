import datetime
import os
import pickle
import shutil
import threading
import platform
import hashlib
import subprocess

from PyQt5 import QtWidgets

import myconstants


def thread(my_func):
    """
    Запускает функцию в отдельным процессом.
    """
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
    # и превратим их в абсолютные пути
    dir_list = [os.path.join(path, x) for x in os.listdir(path)]

    if dir_list:
        # Создадим список из путей к файлам и дат их создания.
        date_list = [[x, os.path.getctime(x)] for x in dir_list]

        # Отсортируем список по дате создания в обратном порядке
        sort_date_list = sorted(date_list, key=lambda x: x[1], reverse=True)

        # Выведем первый элемент списка. Он и будет самым последним по дате
        return sort_date_list[0][0]
    else:
        return None


def open_download_dir():
    open_dir_in_explore(get_download_dir())


def open_dir_in_explore(dir_path):
    os.system(f"explorer.exe {dir_path}")


@thread
def copy_file_as_drop_process(mainwindow, xls_files):
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

    mainwindow.ui.set_status("Обрабатываем Excel файл" + ("ы:" if len(xls_files) > 1 else ":"))
    for file_num, one_file_path in enumerate(xls_files):
        if file_num > 0:
            mainwindow.ui.set_status("")

        if len(xls_files) == 1:
            mainwindow.ui.set_status(f"   {one_file_path}")
        else:
            mainwindow.ui.set_status(f"   {file_num + 1}. {one_file_path}")
        this_file_name = os.path.basename(one_file_path)

        if drug_and_drop_type >= 2:
            # Проверим структуру файла:
            ret_value = open_and_test_raw_struct(one_file_path, short_text=True)
            if type(ret_value) == str:
                QtWidgets.QMessageBox.question(mainwindow, f"Файл: {this_file_name}",
                                               ret_value,
                                               QtWidgets.QMessageBox.Yes)

                mainwindow.ui.set_status("   Структура файла не соответствует требованиям.")
                mainwindow.ui.set_status("   Копирование не отклонено.")
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
                mainwindow.ui.set_status(f"   Имя файла меняется на {new_filename}.")

        else:
            new_filename = this_file_name

        raw_file_path = raw_section_path + "/" + new_filename
        if os.path.isfile(raw_file_path):
            result = QtWidgets.QMessageBox.question(mainwindow, "Заменить файл?",
                                                    "В папке, где находятся данные, выгруженные из DES.LM" +
                                                    f"Файл с таким названием уже есть {new_filename}\n\n" +
                                                    "Вы действительно хотите переписать его новым файлом?",
                                                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                                    QtWidgets.QMessageBox.No)

            if result == QtWidgets.QMessageBox.No:
                mainwindow.ui.set_status("   Пользователь отказался от копирования.")
                continue

        try:
            shutil.copy(one_file_path, raw_file_path)
            select_filename = os.path.splitext(os.path.basename(raw_file_path))[0]
            counter += 1
            mainwindow.ui.set_status("   Файл скопирован.")
        except (OSError, shutil.Error):
            QtWidgets.QMessageBox.question(mainwindow, "Ошибка копирования.",
                                           "Не удалось скопировать файл с данными выгруженными из DES.LM.",
                                           QtWidgets.QMessageBox.Yes)

            mainwindow.ui.set_status("   Копирование не удалось - возникли ошибки.")
            continue

        if drug_and_drop_type == 4:
            if this_file_name == new_filename:
                # Не надо переименовывать файл сам в себя.
                mainwindow.ui.set_status(
                    f"   Исходный файл переименовывать не надо, так как он уже имеет нужное имя: {new_filename}.")
            else:
                new_src_file_path = os.path.join(os.path.dirname(one_file_path), new_filename)
                try:
                    os.rename(one_file_path, new_src_file_path)
                    mainwindow.ui.set_status("   Исходный файл так же переименован.")
                except (OSError, shutil.Error):
                    QtWidgets.QMessageBox.question(mainwindow, "Ошибка копирования.",
                                                   "Не удалось скопировать файл с данными выгруженными из DES.LM.",
                                                   QtWidgets.QMessageBox.Yes)
                    mainwindow.ui.set_status("   Переименование исходного файла не удалось.")

    if counter == 1:
        mainwindow.refresh_raw_files_list(select_filename)
    else:
        mainwindow.refresh_raw_files_list()

    mainwindow.ui.set_status(myconstants.TEXT_LINES_SEPARATOR)

    mainwindow.parent.drag_and_prop_in_process = False
    mainwindow.ui.lock_unlock_interface_items()


def is_admin():
    from mytablefuncs import get_parameter_value
    user_files_path = get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME)
    hash_string = "#" + platform.node() + "#" + os.environ.get('USERNAME') + "#"
    hash_string = hashlib.blake2s(hash_string.encode(encoding = "utf-8"), digest_size=5).hexdigest()
    hash_file = os.path.join(user_files_path, hash_string)

    return(os.path.isfile(hash_file))

def open_file_in_application(file_name):
    subprocess.Popen(file_name, shell=True)

if __name__ == "__main__":
    print(get_files_list("RawData", "", ".xlsx"))
