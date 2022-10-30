import sys
import os

import myconstants
import mainform
import reportcreater
import mytablefuncs
import myutils

class MyApplication:
    def __init__(self):
        myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        self.report_parameters = MyReportParameters(self)
        self.reporter = reportcreater.ReportCreater(self)
        self._mainwindow = mainform.MyWindow(self)
        self._mainwindow.ui.setup_form(self.reporter.get_reports_list(),
                                        myutils.get_files_list(reportcreater.get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)))

        self._mainwindow.show()
        self.report_parameters.is_all_parametars_exist()
        
        sys.exit(self._mainwindow._app.exec_())
    
class MyReportParameters:
    def __init__(self, parent):
        self.parent = parent
        self.slasterror = ""
    
    def update(self, raw_file_name, report_file_name):
        self.slasterror = ""
        if report_file_name is None:
            self.report_prepared_name = ""
            self.parent._mainwindow.ui.set_status("Необходимо выбрать отчётную форму.")
            self.parent._mainwindow.ui.enable_buttons()
            return False
        if raw_file_name is None:
            self.report_prepared_name = ""
            self.parent._mainwindow.ui.set_status("Необходимо выбрать файл, выгруженный из DES.LM для формирования отчёта.")
            self.parent._mainwindow.ui.enable_buttons()
            return False

        # Сохраним параметры для данного отчёта - требуются во время исполнение.
        # Их необходимо сохранить, потому что во время исполнения в основном окне
        # может быть выбран другой отчёт и параметры на главном окне изменятся.
        myconstants.ROUND_FTE_VALUE = mytablefuncs.get_parameter_value(myconstants.ROUND_FTE_SECTION_NAME, myconstants.ROUND_FTE_DEFVALUE)
        myconstants.MEANOURSPERMONTH_VALUE = mytablefuncs.get_parameter_value(myconstants.MEANOURSPERMONTH_SECTION_NAME, myconstants.MEANOURSPERMONTH_DEFVALUE)

        # Парамертры без префиксов будем использовать для получения
        self.p_delete_vip = self.parent._mainwindow.ui.checkBoxDeleteVIP.isChecked()
        self.p_delete_not_prod_units = self.parent._mainwindow.ui.checkBoxDeleteNotProduct.isChecked()
        self.p_delete_without_fact = self.parent._mainwindow.ui.checkBoxDeleteWithoutFact.isChecked()
        self.p_curr_month_half = self.parent._mainwindow.ui.checkBoxCurrMonthAHalf.isChecked()
        self.p_delete_pers_data = self.parent._mainwindow.ui.checkBoxDelPDn.isChecked()
        self.p_delete_vacation = self.parent._mainwindow.ui.checkBoxDeleteVac.isChecked()
        self.p_virtual_FTE = self.parent._mainwindow.ui.checkBoxAddVFTE.isChecked()
        self.p_save_without_formulas = self.parent._mainwindow.ui.checkBoxSaveWithOutFotmulas.isChecked()
        self.p_delete_rawdata_sheet = self.parent._mainwindow.ui.checkBoxDelRawData.isChecked()
        self.p_open_in_excel = self.parent._mainwindow.ui.checkBoxOpenExcel.isChecked()
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
    
    def is_all_parametars_exist(self):
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
                self.slasterror = f"(!) Отсутствует необходимая дирректория: {one_folder_info[1]}.\nИмя дирректории: {spath}\n\nВыполнение программы невозможно."
                self.parent._mainwindow.ui.plainTextEdit.setPlainText(self.slasterror)
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
                (myconstants.EMAILS_TABLE, "Таблица адресов электроной почты"), 
                (myconstants.VIP_TABLE, "Таблица списка VIP"), 
                (myconstants.PORTFEL_TABLE, "Таблица списка портфелей проектов"), 
                (myconstants.ISDOGNAME_TABLE, "Таблица наименований ИС из контракта"), 
            ]

            for one_file_info in files_list:
                if not os.path.isfile(s_section_path + "/" + one_file_info[0]):
                    self.slasterror = f"(!) Отсутствует файл настройки: {one_file_info[1]}.\nИмя файла: {one_file_info[0]}\n\nВыполнение программы невозможно."
                    self.parent._mainwindow.ui.plainTextEdit.setPlainText(self.slasterror)
                    ret_value = False
                    break
            
        return ret_value
