import sys

import myconstants
import mainform
import reportcreater
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
        

