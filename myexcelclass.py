import win32com.client
import myconstants
import myutils


class MyExcel:
    def __init__(self, report_parameters):
        self.oexcel = win32com.client.Dispatch("Excel.Application")
        self.oexcel.Visible = self.oexcel.WorkBooks.Count > 0
        self.save_DisplayAlerts = self.oexcel.DisplayAlerts
        self.oexcel.DisplayAlerts = False
        
        self.report_parameters = report_parameters
        
        self.work_book = self.oexcel.Workbooks.Open(report_parameters.report_prepared_name)
        self.currwindow = self.oexcel.ActiveWindow
        self.currwindow.WindowState = myconstants.EXCELWINDOWSTATE_MIN
        self.save_excel_calc_status = self.oexcel.Calculation
        self.oexcel.Calculation = myconstants.EXCEL_MANUAL_CALC
        self.not_ready = not self.test_structure()
        
        self.report_prepared = False
        
    def test_structure(self):
        ui_handle = self.report_parameters.parent.mainwindow.ui
        if myconstants.RAW_DATA_SHEET_NAME not in self.get_sheets_list():
            ui_handle.add_text_to_log_box("")
            ui_handle.add_text_to_log_box("")
            ui_handle.add_text_to_log_box("[Ошибка в структуре отчета]")
            ui_handle.add_text_to_log_box("")
            ui_handle.add_text_to_log_box("В файле для выбранной формы отчёта отсутствует необходимый лист для данных.")
            ui_handle.add_text_to_log_box("Формирование отчёта не возможно.")

            return False
        elif myconstants.UNIQE_LISTS_SHEET_NAME not in self.get_sheets_list():
            ui_handle.add_text_to_log_box("")
            ui_handle.add_text_to_log_box("")
            ui_handle.add_text_to_log_box("[Ошибка в структуре отчета]")
            ui_handle.add_text_to_log_box("")
            ui_handle.add_text_to_log_box("В файле для выбранной формы отчёта отсутствует необходимый лист для уникальных списков.")
            ui_handle.add_text_to_log_box("Формирование отчёта не возможно.")

            return False
        elif myconstants.SETTINGS_SHEET_NAME not in self.get_sheets_list():
            ui_handle.add_text_to_log_box("")
            ui_handle.add_text_to_log_box("")
            ui_handle.add_text_to_log_box("[Ошибка в структуре отчета]")
            ui_handle.add_text_to_log_box("")
            ui_handle.add_text_to_log_box("В файле для выбранной формы отчёта отсутствует необходимый лист c настройками.")
            ui_handle.add_text_to_log_box("Формирование отчёта не возможно.")

            return False
        else:
            ui_handle.change_last_log_box_text("Пройдена проверка структуры файла, содержащего форму отчёта.")

        ui_handle.add_text_to_log_box("Файл Excel с формой отчёта подгружен.")
        return True
    
    def get_sheets_list(self):
        return [one_sheet.Name for one_sheet in self.work_book.Sheets]

    def save_report(self):
        self.oexcel.Calculation = self.save_excel_calc_status
        for curr_sheet_name in self.get_sheets_list():
            if curr_sheet_name[-1] == myconstants.REPLACE_EQ_SHEET_MARKER:
                self.work_book.Sheets[curr_sheet_name].Cells.Replace(What="=", Replacement="=", FormulaVersion=1)
        if self.report_parameters.p_save_without_formulas:
            for curr_sheet_name in self.get_sheets_list():
                if curr_sheet_name not in myconstants.SHEETS_DONT_DELETE_FORMULAS:
                    column1 = self.work_book.Sheets[curr_sheet_name].UsedRange.Column
                    column2 = self.work_book.Sheets[curr_sheet_name].UsedRange.Columns(self.work_book.Sheets[curr_sheet_name].UsedRange.Columns.Count).Column
                    
                    row1 = self.work_book.Sheets[curr_sheet_name].UsedRange.Row
                    row2 = self.work_book.Sheets[curr_sheet_name].UsedRange.Rows(self.work_book.Sheets[curr_sheet_name].UsedRange.Rows.Count).Row
                    
                    cell1 = self.work_book.Sheets[curr_sheet_name].Cells(row1 + 3, column1).Address
                    cell2 = self.work_book.Sheets[curr_sheet_name].Cells(row2, column2).Address
                    
                    self.work_book.Sheets[curr_sheet_name].Range(cell1, cell2).Value = self.work_book.Sheets[curr_sheet_name].Range(cell1, cell2).Value
            
            if self.report_parameters.p_delete_rawdata_sheet:
                for one_sheet_name in myconstants.DELETE_SHEETS_LIST_IF_NO_FORMULAS:
                    if one_sheet_name in self.get_sheets_list():
                        self.work_book.Sheets[one_sheet_name].Delete()

        self.work_book.Save()

    def __del__(self):
        if self.not_ready:
            pass
        else:
            if not self.report_prepared:
                self.oexcel.Calculation = self.save_excel_calc_status
                self.work_book.Close()
                myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
            else:
                # Отчёт был подготовлен. Закончим его обработку.
                report_prepared_name = self.report_parameters.report_prepared_name
                self.report_parameters.parent.mainwindow.set_status_bar_text("Отчёт сформирован. Идёт сохранение...")
                self.report_parameters.parent.mainwindow.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
                self.report_parameters.parent.mainwindow.add_text_to_log_box(f"Сохраняем в файл: {myutils.rel_path(report_prepared_name)}")

                # -----------------------------------
                # Произведём пересчёт ячеек иначе, если не сработают формулы
                # используемые для проставления признаков скрываемых/удаляемых строк/столбцов.
                self.oexcel.Calculate()
                self.hide_and_delete_rows_and_columns()

                # Скроем вспомогательные листы:
                if myconstants.UNIQE_LISTS_SHEET_NAME in self.get_sheets_list():
                    self.work_book.Sheets[myconstants.UNIQE_LISTS_SHEET_NAME].Visible = False
                
                if myconstants.SETTINGS_SHEET_NAME in self.get_sheets_list():
                    self.work_book.Sheets[myconstants.SETTINGS_SHEET_NAME].Visible = False

                self.save_report()

                self.report_parameters.parent.mainwindow.add_text_to_log_box("Отчёт сохранён.")
                myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, report_prepared_name)

                if self.report_parameters.p_open_in_excel:
                    self.oexcel.Visible = True
                    self.currwindow.WindowState = myconstants.EXCELWINDOWSTATE_MAX
                else:
                    self.work_book.Close()

                self.report_parameters.parent.reporter.start_prog_time.display()
                self.report_parameters.parent.mainwindow.set_status_bar_text("Отчёт сформирован и сохранён")

        self.oexcel.DisplayAlerts = self.save_DisplayAlerts

    def hide_and_delete_rows_and_columns(self):
        for curr_sheet_name in self.get_sheets_list():
            if curr_sheet_name not in myconstants.SHEETS_DONT_DELETE_FORMULAS:
                row_counter = 0
                p_found_first_row = False
                last_row_4_test = myconstants.PARAMETER_MAX_ROWS_TEST_IN_REPORT
                range_from_excel = \
                    self.work_book.Sheets[curr_sheet_name].Range(self.work_book.Sheets[curr_sheet_name].Cells(1, 1), self.work_book.Sheets[curr_sheet_name].Cells(last_row_4_test, 1)).Value

                # Ищем первый признак 'delete'
                for row_counter in range(len(range_from_excel)):
                    row_del_flag_value = range_from_excel[row_counter][0]
                    if row_del_flag_value is None:
                        p_found_first_row = False
                        break
                    
                    if type(row_del_flag_value) == str:
                        row_del_flag_value = row_del_flag_value.replace(" ", "")
                        if row_del_flag_value == myconstants.DELETE_ROW_MARKER:
                            p_found_first_row = True
                            break

                if p_found_first_row:
                    first_row_with_del = row_counter + 1
                    last_row_with_del = row_counter
                    while last_row_with_del < len(range_from_excel):
                        row_del_flag_value = range_from_excel[last_row_with_del][0]
                        if (type(row_del_flag_value) != str) or row_del_flag_value.replace(" ", "") != myconstants.DELETE_ROW_MARKER:
                            break
                        last_row_with_del += 1

                    self.work_book.Sheets[curr_sheet_name].Range(self.work_book.Sheets[curr_sheet_name].Cells(
                        first_row_with_del, 1), self.work_book.Sheets[curr_sheet_name].Cells(last_row_with_del, 1)).Rows.EntireRow.Delete()
        # -----------------------------------
                if curr_sheet_name not in [myconstants.RAW_DATA_SHEET_NAME, myconstants.UNIQE_LISTS_SHEET_NAME, myconstants.SETTINGS_SHEET_NAME]:
                    # Скрываем строки с признаком 'hide'
                    for curr_row in range(1, myconstants.NUM_ROWS_FOR_HIDE + 1):
                        cell_value = self.work_book.Sheets[curr_sheet_name].Cells(curr_row, 1).Value
                        if type(cell_value) == str and cell_value is not None and cell_value.replace(" ", "") == myconstants.HIDE_MARKER:
                            pass
                            self.work_book.Sheets[curr_sheet_name].Rows(curr_row).Hidden = True
                    # Скрываем столбцы с признаком 'hide'
                    for curr_col in range(1, myconstants.NUM_COLUMNS_FOR_HIDE + 1):
                        cell_value = self.work_book.Sheets[curr_sheet_name].Cells(1, curr_col).Value
                        if type(cell_value) == str and cell_value is not None and cell_value.replace(" ", "") == myconstants.HIDE_MARKER:
                            self.work_book.Sheets[curr_sheet_name].Columns(curr_col).Hidden = True
                        else:
                            pass
        # -----------------------------------
