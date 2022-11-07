import sys
import os

import myconstants
import mainform
import reportcreater
import mytablefuncs
import myutils

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from PyQt5 import QtGui


class MyHandler(FileSystemEventHandler):
    def __init__(self, parent):
        self.parent = parent
    
    def on_created(self, event):
        if event.is_directory:
            return
        self.parent.mainwindow.refresh_raw_files_list()

    def on_deleted(self, event):
        if event.is_directory:
            return
        self.parent.mainwindow.refresh_raw_files_list()

    def on_moved(self, event):
        if event.is_directory:
            return
        old_filename = os.path.splitext(os.path.basename(event.src_path))[0]
        new_filename = os.path.splitext(os.path.basename(event.dest_path))[0]
        self.parent.mainwindow.refresh_raw_files_list(new_filename)


class MyApplication:
    def __init__(self):
        myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        self.report_parameters = MyReportParameters(self)
        self.reporter = reportcreater.ReportCreater(self)
        self.mainwindow = mainform.MyWindow(self)
        self.mainwindow.ui.setup_form(self.reporter.get_reports_list())

        self.mainwindow.show()
        if self.report_parameters.is_all_parameters_exist():
            self.event_handler = MyHandler(self)
            self.observer = Observer()
            control_path = mytablefuncs.get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)
            self.observer.schedule(self.event_handler, path=control_path, recursive=False)
            self.observer.start()
        
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
            self.parent.mainwindow.ui.enable_buttons()
            return False
        if raw_file_name is None:
            self.report_prepared_name = ""
            self.parent.mainwindow.ui.set_status("Необходимо выбрать файл, выгруженный из DES.LM для формирования отчёта.")
            self.parent.mainwindow.ui.enable_buttons()
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
        if os.path.isfile(myconstants.SECRET_COSTS_LOCATION + "/" + myconstants.COSTS_TABLE):
            self.p_save_without_formulas = True
            self.p_delete_rawdata_sheet = True

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
        foldefs_list = [
            (myconstants.PARAMETERS_SECTION_NAME, "Параметры"),
            (myconstants.RAW_DATA_SECTION_NAME, "Исходные (сырые) данные"),
            (myconstants.REPORTS_SECTION_NAME, "Отчёты"),
            (myconstants.REPORTS_PREPARED_SECTION_NAME, "Сформированные отчёты"),
        ]
        for one_folder_info in foldefs_list:
            if not os.path.isdir(os.path.join(os.getcwd(), mytablefuncs.get_parameter_value(one_folder_info[0]))):
                spath = one_folder_info[0].replace('\\', '/')
                self.slasterror = f"(!) Отсутствует необходимая директория: {one_folder_info[1]}.\nИмя директории: {spath}\n\nВыполнение программы невозможно."
                self.parent.mainwindow.ui.plainTextEdit.setPlainText(self.slasterror)
                ret_value = False
                break
        
        if ret_value:
            # Проверим параметры:
            s_section_path = mytablefuncs.get_parameter_value(myconstants.PARAMETERS_SECTION_NAME) + "/"
            files_list = [
                (myconstants.MONTH_WORKING_HOURS_TABLE, "Таблица с количеством рабочих часов в месяцах"), 
                (myconstants.DIVISIONS_NAMES_TABLE, "Таблица с наименованиями подразделений"),
                (myconstants.FNS_NAMES_TABLE, "Таблица с наименованиями функциональных направлений"), 
                (myconstants.P_FN_SUBST_TABLE, "Таблица подстановок названий функциональных направлений"), 
                (myconstants.PROJECTS_SUB_TYPES_TABLE, "Таблица с наименованиями подтипов проектов"), 
                (myconstants.PROJECTS_TYPES_DESCR, "Таблица с расшифровкой типов (букв) проектов"), 
                (myconstants.PROJECTS_SUB_TYPES_DESCR, "Таблица с расшифровок подтипов проектов"), 
                (myconstants.COSTS_TABLE, "Таблица часовых ставок"), 
                (myconstants.EMAILS_TABLE, "Таблица адресов электронной почты"),
                (myconstants.VIP_TABLE, "Таблица списка VIP"), 
                (myconstants.PORTFEL_TABLE, "Таблица списка портфелей проектов"), 
                (myconstants.ISDOGNAME_TABLE, "Таблица наименований ИС из контракта"),
                (myconstants.PROJECTS_LIST_ADD_INFO, "Таблица наименований ИС из контракта"),
            ]

            for one_file_info in files_list:
                if not os.path.isfile(s_section_path + "/" + one_file_info[0]):
                    self.slasterror = f"(!) Отсутствует файл настройки: {one_file_info[1]}.\nИмя файла: {one_file_info[0]}\n\nВыполнение программы невозможно."
                    self.parent.mainwindow.ui.plainTextEdit.setPlainText(self.slasterror)
                    ret_value = False
                    break
            
        return ret_value
