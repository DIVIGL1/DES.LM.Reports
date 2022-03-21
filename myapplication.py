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
        self._curr_reporter = reportcreater.ReportCreater()
        
        self._mainwindow = mainform.MyWindow()
        
        self._mainwindow.ui.setup_form(self._curr_reporter.get_reports_list(),
                                        myutils.get_files_list(reportcreater.get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)))

        self._mainwindow.show()

        sys.exit(self._mainwindow._app.exec_())
        

class MyReportParameters:
    def __init__(self, raw_file_name, report_file_name):

        self.start_timer()
        # Получим параметры для данного отчёта - требуются во время исполнение.
        # Их необходимо сохранить, потому что во время исполнения в основном окне
        # может быть выбран другой отчёт и параметры на главном окне изменятся.
        myconstants.ROUND_FTE_VALUE = mytablefuncs.get_parameter_value(myconstants.ROUND_FTE_SECTION_NAME, myconstants.ROUND_FTE_DEFVALUE)
        myconstants.MEANOURSPERMONTH_VALUE = mytablefuncs.get_parameter_value(myconstants.MEANOURSPERMONTH_SECTION_NAME, myconstants.MEANOURSPERMONTH_DEFVALUE)
        
        s_preff = myconstants.DO_IT_PREFFIX
        self.p_delete_not_prod_units =\
            myutils.load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_NONPROD, myconstants.PARAMETER_SAVED_VALUE_DELETE_NONPROD_DEFVALUE)
        self.p_delete_pers_data =\
            myutils.load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_PERSDATA, myconstants.PARAMETER_SAVED_VALUE_DELETE_PERSDATA_DEFVALUE)
        self.p_delete_vacation =\
            myutils.load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_VAC, myconstants.PARAMETER_SAVED_VALUE_DELETE_VAC_DEFVALUE)
        self.p_add_vfte =\
            myutils.load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_ADD_VFTE, myconstants.PARAMETER_SAVED_VALUE_ADD_VFTE_DEFVALUE)
        self.p_save_without_formulas =\
            myutils.load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS, myconstants.PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS_DEFVALUE)
        self.p_delete_rawdata_sheet =\
            myutils.load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DEL_RAWSHEET, myconstants.PARAMETER_SAVED_VALUE_DEL_RAWSHEET_DEFVALUE)
        self.p_open_in_excel =\
            myutils.load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL, myconstants.PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL_DEFVALUE)

        if os.path.isfile(myconstants.SECRET_COSTS_LOCATION + "/" + myconstants.COSTS_TABLE):
            self.p_save_without_formulas = True
            self.p_delete_rawdata_sheet = True
        
        # Получим полные пути до файлов.
        raw_file_name, report_file_name, report_prepared_name = myutils.get_full_files_names(raw_file_name, report_file_name)
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
        
    def start_timer(self):
        self.start_prog_time = time.time()
