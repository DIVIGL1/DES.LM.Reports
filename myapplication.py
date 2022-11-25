import sys
import os
import shutil

import myconstants
import mainform
from reportcreator import reportCreator
import mytablefuncs
import myutils

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


class handlerRawFolder(FileSystemEventHandler):
    def __init__(self, parent):
        self.parent = parent
    
    def on_created(self, event):
        if self.parent.drag_and_prop_in_process:
            return
        if event.is_directory:
            return
        self.parent.mainwindow.refresh_raw_files_list()

    def on_deleted(self, event):
        if self.parent.drag_and_prop_in_process:
            return
        if event.is_directory:
            return
        self.parent.mainwindow.refresh_raw_files_list()

    def on_moved(self, event):
        if self.parent.drag_and_prop_in_process:
            return
        if event.is_directory:
            return
        new_filename = os.path.splitext(os.path.basename(event.dest_path))[0]
        self.parent.mainwindow.refresh_raw_files_list(new_filename)


class handlerUserFolder(FileSystemEventHandler):
    def __init__(self, parent):
        self.parent = parent

    def on_created(self, event):
        if event.is_directory:
            return
        self.call_action()

    def on_deleted(self, event):
        if event.is_directory:
            return
        self.call_action()

    def on_moved(self, event):
        if event.is_directory:
            return
        self.call_action()

    def call_action(self):
        self.parent.mainwindow.ui.update_user_files_menus()


class MyApplication:
    drag_and_prop_in_process = False
    def __init__(self):
        myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        self.report_parameters = MyReportParameters(self)
        self.reporter = reportCreator(self)
        self.mainwindow = mainform.MyWindow(self)
        self.mainwindow.ui.setup_form(self.reporter.get_reports_list())

        self.mainwindow.show()
        if self.report_parameters.is_all_parameters_exist():
            self.event_handler_raw_folder = handlerRawFolder(self)
            self.observer_raw_folder = Observer()
            control_path = mytablefuncs.get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)
            self.observer_raw_folder.schedule(self.event_handler_raw_folder, path=control_path, recursive=False)
            self.observer_raw_folder.start()

        # Проверяем наличие пользовательского каталога.
        # Если нет, то стараемся создать.
        user_files_path = mytablefuncs.get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME)
        if not os.path.exists(user_files_path):
            os.mkdir(user_files_path)

        # Проверим, удалось ли создать папку:
        if os.path.exists(user_files_path):
            # Проверим, существуют ли нужные файлы:
            section = mytablefuncs.get_parameter_value(myconstants.PARAMETERS_SECTION_NAME)

            for one_file in myconstants.USER_FILES_LIST:
                system_file_path = os.path.join(os.path.join(os.getcwd(), section, one_file))
                user_file_path = os.path.join(os.path.join(user_files_path, one_file))
                excluded_file = myconstants.USER_FILES_EXCLUDE_PREFFIX + one_file
                user_excluded_file_path = os.path.join(os.path.join(user_files_path, excluded_file))
                if not (os.path.isfile(user_file_path) or os.path.isfile(user_excluded_file_path)):
                    try:
                        shutil.copyfile(system_file_path, user_file_path)
                    except:
                        self.mainwindow.ui.set_status(myconstants.TEXT_LINES_SEPARATOR)
                        self.mainwindow.ui.set_status(
                            f"Не удалось создать файл пользовательских настроек: {one_file}")
                        self.mainwindow.ui.set_status(myconstants.TEXT_LINES_SEPARATOR)

            self.event_handler_use_folder = handlerUserFolder(self)
            self.observer_user_folder = Observer()
            self.observer_user_folder.schedule(self.event_handler_use_folder, path=user_files_path, recursive=False)
            self.observer_user_folder.start()

        else:
            self.mainwindow.ui.set_status(myconstants.TEXT_LINES_SEPARATOR)
            self.mainwindow.ui.set_status("Не удалось создать пользовательскую папку для хранения пользовательских настроек.")
            self.mainwindow.ui.set_status(myconstants.TEXT_LINES_SEPARATOR)

        self.mainwindow.ui.lock_unlock_interface_items()
        sys.exit(self.mainwindow.app.exec_())
    

class MyReportParameters:
    def __init__(self, parent):
        self.parent = parent
        self.slasterror = ""
        self.p_delete_vip = None
        self.p_delete_not_prod_units = None
        self.p_projects_with_add_info = None
        self.p_delete_without_fact = None
        self.p_delete_without_fact = None
        self.p_curr_month_half = None
        self.p_delete_pers_data = None
        self.p_delete_vacation = None
        self.p_virtual_FTE = None
        self.p_save_without_formulas = None
        self.p_delete_rawdata_sheet = None
        self.p_open_in_excel = None
        self.raw_file_name = None
        self.report_file_name = None
        self.report_prepared_name = None

    def update(self, raw_file_name, report_file_name):
        self.slasterror = ""
        if report_file_name is None:
            self.report_prepared_name = ""
            self.parent.mainwindow.ui.set_status("Необходимо выбрать отчётную форму.")
            return False
        if raw_file_name is None:
            self.report_prepared_name = ""
            self.parent.mainwindow.ui.set_status("Необходимо выбрать файл, выгруженный из DES.LM для формирования отчёта.")
            return False

        # Сохраним параметры для данного отчёта - требуются во время исполнение.
        # Их необходимо сохранить, потому что во время исполнения в основном окне
        # может быть выбран другой отчёт и параметры на главном окне изменятся.
        myconstants.ROUND_FTE_VALUE = mytablefuncs.get_parameter_value(myconstants.ROUND_FTE_SECTION_NAME, myconstants.ROUND_FTE_DEFVALUE)
        myconstants.MEANOURSPERMONTH_VALUE = mytablefuncs.get_parameter_value(myconstants.MEANHOURSPERMONTH_SECTION_NAME, myconstants.MEANOURSPERMONTH_DEFVALUE)

        # Параметры без префиксов будем использовать для получения
        self.p_delete_vip = self.parent.mainwindow.ui.checkBoxDeleteVIP.isChecked()
        self.p_delete_not_prod_units = self.parent.mainwindow.ui.checkBoxDeleteNotProduct.isChecked()
        self.p_projects_with_add_info = self.parent.mainwindow.ui.checkBoxOnlyProjectsWithAdd.isChecked()
        self.p_delete_without_fact = self.parent.mainwindow.ui.checkBoxDeleteWithoutFact.isChecked()
        self.p_curr_month_half = self.parent.mainwindow.ui.checkBoxCurrMonthAHalf.isChecked()
        self.p_delete_pers_data = self.parent.mainwindow.ui.checkBoxDelPDn.isChecked()
        self.p_delete_vacation = self.parent.mainwindow.ui.checkBoxDeleteVac.isChecked()
        self.p_virtual_FTE = self.parent.mainwindow.ui.checkBoxAddVFTE.isChecked()
        self.p_save_without_formulas = self.parent.mainwindow.ui.checkBoxSaveWithOutFotmulas.isChecked()
        self.p_delete_rawdata_sheet = self.parent.mainwindow.ui.checkBoxDelRawData.isChecked()
        self.p_open_in_excel = self.parent.mainwindow.ui.checkBoxOpenExcel.isChecked()

        # Получим полные пути до файлов.
        report_prepared_name = \
            os.path.join( 
                os.path.join(os.getcwd(), mytablefuncs.get_parameter_value(myconstants.REPORTS_PREPARED_SECTION_NAME)),
                raw_file_name + "__" + report_file_name + myconstants.EXCEL_FILES_ENDS
            )
        
        report_file_name = \
            os.path.join( 
                os.path.join(os.getcwd(), mytablefuncs.get_parameter_value(myconstants.REPORTS_SECTION_NAME)),
                myconstants.REPORT_FILE_PREFFIX + report_file_name + myconstants.EXCEL_FILES_ENDS
            )
        
        raw_file_name = \
            os.path.join( 
                os.path.join(os.getcwd(), mytablefuncs.get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)),
                raw_file_name + myconstants.EXCEL_FILES_ENDS
            )

        self.raw_file_name = raw_file_name.replace("\\", "/")
        self.report_file_name = report_file_name.replace("\\", "/")
        self.report_prepared_name = report_prepared_name.replace("\\", "/")
    
    def is_all_parameters_exist(self):
        self.slasterror = ""
        ret_value = True
        
        # Для начала проверим наличие всех необходимых папок,
        # в которых должны храниться данные программы:
        folders_list = [
            (myconstants.PARAMETERS_SECTION_NAME, "Параметры"),
            (myconstants.RAW_DATA_SECTION_NAME, "Исходные (сырые) данные"),
            (myconstants.REPORTS_SECTION_NAME, "Отчёты"),
            (myconstants.REPORTS_PREPARED_SECTION_NAME, "Сформированные отчёты"),
        ]
        for one_folder_info in folders_list:
            if not os.path.isdir(os.path.join(os.getcwd(), mytablefuncs.get_parameter_value(one_folder_info[0]))):
                spath = one_folder_info[0].replace('\\', '/')
                self.slasterror = f"(!) Отсутствует необходимая директория: {one_folder_info[1]}.\nИмя директории: {spath}\n\nВыполнение программы невозможно."
                self.parent.mainwindow.ui.plainTextEdit.setPlainText(self.slasterror)
                ret_value = False
                break
        
        if ret_value:
            # Проверим параметры:
            s_section_path = mytablefuncs.get_parameter_value(myconstants.PARAMETERS_SECTION_NAME) + "/"
            files_list = myconstants.PARAMETERS_ALL_TABLES

            for one_file in files_list.keys():
                if not os.path.isfile(s_section_path + "/" + one_file):
                    self.slasterror = f"(!) Отсутствует файл настройки: {files_list[one_file][0]}.\nИмя файла: {one_file}\n\nВыполнение программы невозможно."
                    self.parent.mainwindow.ui.plainTextEdit.setPlainText(self.slasterror)
                    ret_value = False
                    break
            
        return ret_value
