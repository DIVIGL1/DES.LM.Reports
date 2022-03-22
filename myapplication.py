import sys
import time
import os

import myconstants
import mainform
import reportcreater
import mytablefuncs
import myutils
import win32com.client

class MyExcel:
    def __init__(self, excel_file_name):
        self._oExcel = win32com.client.Dispatch("Excel.Application")
        self._oExcel.Visible = self._oExcel.WorkBooks.Count > 0
        self._oExcel.DisplayAlerts = False
            
        self._wb = self._oExcel.Workbooks.Open(excel_file_name)
        self._currwindow = self._oExcel.ActiveWindow
        self._currwindow.WindowState = myconstants.EXCELWINDOWSTATE_MIN
        self._n_save_excel_calc_status = self._oExcel.Calculation
        self._oExcel.Calculation = myconstants.EXCEL_MANUAL_CALC
        
        self.quit_on_destroy()
    
    def dont_quit_on_destroy(self):
        self.p_quit_excel_on_destroy = False

    def quit_on_destroy(self):
        self.p_quit_excel_on_destroy = False

    def get_sheets_list(self):
        return([one_sheet.Name for one_sheet in self._wb.Sheets])

    def save_report(self, wb, p_save_without_formulas, p_delete_rawdata_sheet):
        wb.Save()
        if p_save_without_formulas: 
            for curr_sheet_name in self.get_sheets_list():
                if curr_sheet_name not in myconstants.SHEETS_DONT_DELETE_FORMULAS:
                    self._wb.Sheets[curr_sheet_name].UsedRange.Value = self._wb.Sheets[curr_sheet_name].UsedRange.Value
            
            if p_delete_rawdata_sheet:
                for one_sheet_name in myconstants.DELETE_SHEETS_LIST_IF_NO_FORMULAS:
                    if one_sheet_name in self.get_sheets_list():
                        self._wb.Sheets[one_sheet_name].Delete()
            
            self._wb.Save()
        
class MyApplication:
    def __init__(self):
        self.report_parameters = MyReportParameters(self)
        self._curr_reporter = reportcreater.ReportCreater(self)
        self._mainwindow = mainform.MyWindow(self)
        self._mainwindow.ui.setup_form(self._curr_reporter.get_reports_list(),
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

        self.start_timer()
        # Сохраним параметры для данного отчёта - требуются во время исполнение.
        # Их необходимо сохранить, потому что во время исполнения в основном окне
        # может быть выбран другой отчёт и параметры на главном окне изменятся.
        myconstants.ROUND_FTE_VALUE = mytablefuncs.get_parameter_value(myconstants.ROUND_FTE_SECTION_NAME, myconstants.ROUND_FTE_DEFVALUE)
        myconstants.MEANOURSPERMONTH_VALUE = mytablefuncs.get_parameter_value(myconstants.MEANOURSPERMONTH_SECTION_NAME, myconstants.MEANOURSPERMONTH_DEFVALUE)

        # Парамертры без префиксов будем использовать для получения
        self.p_delete_not_prod_units = self.parent._mainwindow.ui.checkBoxDeleteNotProduct.isChecked()
        self.p_delete_pers_data = self.parent._mainwindow.ui.checkBoxDelPDn.isChecked()
        self.p_delete_vacation = self.parent._mainwindow.ui.checkBoxDeleteVac.isChecked()
        self.p_add_vfte = self.parent._mainwindow.ui.checkBoxAddVFTE.isChecked()
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
        s_section_path = mytablefuncs.get_parameter_value(myconstants.PARAMETERS_SECTION_NAME) + "/"
        files_list = [
            (myconstants.MONTH_WORKING_HOURS_TABLE, "Таблица с количеством рабочих часов в месяцах"), 
            (myconstants.DIVISIONS_NAMES_TABLE, "Таблица с наименованиями подразделений"),
            (myconstants.FNS_NAMES_TABLE, "Таблица с наименованиями функциональных направлений"), 
            (myconstants.P_FN_SUBST_TABLE, "Таблица подстановок названий функциональных направлений"), 
            (myconstants.PROJECTS_SUB_TYPES_TABLE, "Таблица с наименованиями подтипов проектов"), 
            (myconstants.PROJECTS_TYPES_DESCR, "Таблица расшифровок типов проектов"), 
            (myconstants.PROJECTS_SUB_TYPES_DESCR, "Таблица с расшифровок подтипов проектов"), 
            (myconstants.COSTS_TABLE, "Таблица часовых ставок"), 
        ]
        ret_value = True
        for one_file_info in files_list:
            if not os.path.isfile(s_section_path + "/" + one_file_info[0]):
                self.slasterror = f"(!) Отсутствует файл настройки: {one_file_info[1]}.\nИмя файла: {one_file_info[0]}\n\nВыполнение программы не возможно."
                self.parent._mainwindow.ui.plainTextEdit.setPlainText(self.slasterror)
                ret_value = False
                break
            
        return ret_value

    def start_timer(self):
        self.start_prog_time = time.time()
