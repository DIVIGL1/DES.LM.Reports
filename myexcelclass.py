import win32com.client
import logging
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
        logging.debug(' MyExcel.save_report: Let`s recalc.')
        self.oexcel.Calculation = self.save_excel_calc_status
        for curr_sheet_name in self.get_sheets_list():
            if curr_sheet_name[-1] == myconstants.REPLACE_EQ_SHEET_MARKER:
                self.work_book.Sheets[curr_sheet_name].Cells.Replace(What="=", Replacement="=", FormulaVersion=1)
        logging.debug(f' MyExcel.save_report: Save without formulas == {self.report_parameters.p_save_without_formulas}.')
        if self.report_parameters.p_save_without_formulas:
            for curr_sheet_name in self.get_sheets_list():
                logging.debug(f' MyExcel.save_report: Replace formulas with values on sheet {curr_sheet_name}.')
                # Если в конце имени листа стоит признак того, что на этом листе не надо стирать формулы,
                # то необходимо сохранить формулы, а имя листа слегка поправить, убрав из имени признак:
                if curr_sheet_name[-len(myconstants.FLAG_DONT_DELETE_FORMULAS_ON_THE_SHEET):] == myconstants.FLAG_DONT_DELETE_FORMULAS_ON_THE_SHEET:
                    self.work_book.Sheets[curr_sheet_name].Name = curr_sheet_name[:-len(myconstants.FLAG_DONT_DELETE_FORMULAS_ON_THE_SHEET)]
                    logging.debug(f' MyExcel.save_report: No need to replace formulas for especial sheet {curr_sheet_name}. Name of sheet changed')
                else:
                    # Если название листа не входит в список тех на которых необходимо
                    # сохранить формулы, то нужно заменить формулы значениями:
                    if curr_sheet_name not in myconstants.SHEETS_DONT_DELETE_FORMULAS:
                        logging.debug(f' MyExcel.save_report: stage 1.')
                        column1 = self.work_book.Sheets[curr_sheet_name].UsedRange.Column
                        logging.debug(f' MyExcel.save_report: stage 2.')
                        logging.debug(f' MyExcel.save_report: stage 7. self.work_book.Sheets[curr_sheet_name].UsedRange.Columns.Count = {self.work_book.Sheets[curr_sheet_name].UsedRange.Columns.Count}')
                        column2 = self.work_book.Sheets[curr_sheet_name].UsedRange.Columns(self.work_book.Sheets[curr_sheet_name].UsedRange.Columns.Count).Column
                        logging.debug(f' MyExcel.save_report: stage 7. column2 = {column2}')

                        logging.debug(f' MyExcel.save_report: stage 3.')
                        row1 = self.work_book.Sheets[curr_sheet_name].UsedRange.Row
                        logging.debug(f' MyExcel.save_report: stage 4.')
                        row2 = self.work_book.Sheets[curr_sheet_name].UsedRange.Rows(self.work_book.Sheets[curr_sheet_name].UsedRange.Rows.Count).Row

                        logging.debug(f' MyExcel.save_report: stage 5.')
                        cell1 = self.work_book.Sheets[curr_sheet_name].Cells(row1 + 3, column1).Address
                        logging.debug(f' MyExcel.save_report: stage 6.')
                        cell2 = self.work_book.Sheets[curr_sheet_name].Cells(row2, column2).Address

                        logging.debug(f' MyExcel.save_report: stage 7. cell1 = {cell1}  cell2 = {cell2}')
                        self.work_book.Sheets[curr_sheet_name].Range(cell1, cell2).Value = self.work_book.Sheets[curr_sheet_name].Range(cell1, cell2).Value
                        logging.debug(f' MyExcel.save_report: End processing sheet {curr_sheet_name}.')
                    else:
                        logging.debug(f' MyExcel.save_report: No need to replace formulas for sheet {curr_sheet_name}.')

            logging.debug(f' MyExcel.save_report: Delete raw data sheet == {self.report_parameters.p_delete_rawdata_sheet}.')
            if self.report_parameters.p_delete_rawdata_sheet:
                for one_sheet_name in myconstants.DELETE_SHEETS_LIST_IF_NO_FORMULAS:
                    logging.debug(f' MyExcel.save_report: Sheet {one_sheet_name}.')
                    if one_sheet_name in self.get_sheets_list():
                        logging.debug(f' MyExcel.save_report: Deleting sheet {one_sheet_name}.')
                        self.work_book.Sheets[one_sheet_name].Delete()
                    else:
                        logging.debug(f' MyExcel.save_report: Sheet {one_sheet_name} no need to delete.')
        else:
            for curr_sheet_name in self.get_sheets_list():
                # Для листов в конце имени листа стоит признак того,
                # что на этом листе не надо стирать формулы нужно поправить имя,
                # даже если формулы заменять на значения не надо:
                if curr_sheet_name[-len(myconstants.FLAG_DONT_DELETE_FORMULAS_ON_THE_SHEET):] == myconstants.FLAG_DONT_DELETE_FORMULAS_ON_THE_SHEET:
                    self.work_book.Sheets[curr_sheet_name].Name = curr_sheet_name[:-len(myconstants.FLAG_DONT_DELETE_FORMULAS_ON_THE_SHEET)]
                    logging.debug(f' MyExcel.save_report: For especial sheet {curr_sheet_name} name was changed.')

        for curr_sheet_name in self.get_sheets_list():
            logging.debug(f' MyExcel.save_report: Before Replace marker "{myconstants.MAKE_FORMULAS_MARKER}" with sign "=" on sheet {curr_sheet_name}.')
            self.work_book.Sheets[curr_sheet_name].Cells.Replace(
                What=myconstants.MAKE_FORMULAS_MARKER,
                Replacement="=",
                LookAt=2,
                SearchOrder=1,
                MatchCase=False,
                SearchFormat=False,
                ReplaceFormat=False,
                FormulaVersion=1
            )
            logging.debug(f' MyExcel.save_report: After Replace marker "{myconstants.MAKE_FORMULAS_MARKER}" with sign "=" on sheet {curr_sheet_name}.')

        logging.debug(' MyExcel.save_report: Before call Excel.save().')
        self.work_book.Save()
        logging.debug(' MyExcel.save_report: After call Excel.save().')

    def __del__(self):
        if self.not_ready:
            logging.debug(' MyExcel.__del__: do nothing.')
            pass
        else:
            if not self.report_prepared:
                logging.debug(' MyExcel.__del__: not self.report_prepared.')
                self.oexcel.Calculation = self.save_excel_calc_status
                self.work_book.Close()
                myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
            else:
                # Отчёт был подготовлен. Закончим его обработку.
                logging.debug(' MyExcel.__del__: Report was prepared.')
                report_prepared_name = self.report_parameters.report_prepared_name
                self.report_parameters.parent.mainwindow.set_status_bar_text("Отчёт сформирован. Идёт сохранение...")
                self.report_parameters.parent.mainwindow.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
                self.report_parameters.parent.mainwindow.add_text_to_log_box(f"Сохраняем в файл: {myutils.rel_path(report_prepared_name)}")

                # -----------------------------------
                # Произведём пересчёт ячеек иначе, если не сработают формулы
                # используемые для проставления признаков скрываемых/удаляемых строк/столбцов.
                logging.debug(' MyExcel.__del__: Recalculating.')
                self.oexcel.Calculate()
                logging.debug(' MyExcel.__del__: Ready to hide and delete rows and columns.')
                self.hide_and_delete_rows_and_columns()
                logging.debug(' MyExcel.__del__: Hide_and_delete_rows_and_columns was processed.')

                # Скроем вспомогательные листы:
                logging.debug(' MyExcel.__del__: Hiding some sheets.')
                if myconstants.UNIQE_LISTS_SHEET_NAME in self.get_sheets_list():
                    self.work_book.Sheets[myconstants.UNIQE_LISTS_SHEET_NAME].Visible = False
                
                if myconstants.SETTINGS_SHEET_NAME in self.get_sheets_list():
                    self.work_book.Sheets[myconstants.SETTINGS_SHEET_NAME].Visible = False

                logging.debug(' MyExcel.__del__: Let`s save report.')
                self.save_report()
                logging.debug(' MyExcel.__del__: Report was saved.')

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
        logging.debug(' MyExcel.__del__.hide_and_delete_rows_and_columns. Entered the procedure.')
        for curr_sheet_name in self.get_sheets_list():
            logging.debug(' MyExcel.__del__.hide_and_delete_rows_and_columns. Iterating over sheets:')
            if curr_sheet_name not in myconstants.SHEETS_DONT_DELETE_FORMULAS and curr_sheet_name[-1] != "-":
                logging.debug(f' MyExcel.__del__.hide_and_delete_rows_and_columns. Sheet: {curr_sheet_name}')
                row_counter = 0
                p_found_first_row = False
                last_row_4_test = myconstants.PARAMETER_MAX_ROWS_TEST_IN_REPORT
                range_from_excel = \
                    self.work_book.Sheets[curr_sheet_name].Range(self.work_book.Sheets[curr_sheet_name].Cells(1, 1), self.work_book.Sheets[curr_sheet_name].Cells(last_row_4_test, 1)).Value

                # Ищем первый признак 'delete'
                logging.debug(f' MyExcel.__del__.hide_and_delete_rows_and_columns. Looking for DELETE_ROW_MARKER')
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

                if not p_found_first_row:
                    logging.debug(f' MyExcel.__del__.hide_and_delete_rows_and_columns. p_found_first_row == False')
                else:
                    logging.debug(f' MyExcel.__del__.hide_and_delete_rows_and_columns. p_found_first_row == True')
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
                if curr_sheet_name in [myconstants.RAW_DATA_SHEET_NAME, myconstants.UNIQE_LISTS_SHEET_NAME, myconstants.SETTINGS_SHEET_NAME]:
                    logging.debug(f' MyExcel.__del__.hide_and_delete_rows_and_columns. UnProcessing sheet.')
                else:
                    logging.debug(f' MyExcel.__del__.hide_and_delete_rows_and_columns. Hiding rows.')
                    # Скрываем строки с признаком 'hide'
                    for curr_row in range(1, myconstants.NUM_ROWS_FOR_HIDE + 1):
                        cell_value = self.work_book.Sheets[curr_sheet_name].Cells(curr_row, 1).Value
                        if type(cell_value) == str and cell_value is not None and cell_value.replace(" ", "") == myconstants.HIDE_MARKER:
                            pass
                            self.work_book.Sheets[curr_sheet_name].Rows(curr_row).Hidden = True

                    logging.debug(f' MyExcel.__del__.hide_and_delete_rows_and_columns. Hiding columns.')
                    # Скрываем столбцы с признаком 'hide'
                    for curr_col in range(1, myconstants.NUM_COLUMNS_FOR_HIDE + 1):
                        cell_value = self.work_book.Sheets[curr_sheet_name].Cells(1, curr_col).Value
                        if type(cell_value) == str and cell_value is not None and cell_value.replace(" ", "") == myconstants.HIDE_MARKER:
                            self.work_book.Sheets[curr_sheet_name].Columns(curr_col).Hidden = True
                        else:
                            pass
        # -----------------------------------
        logging.debug(f' MyExcel.__del__.hide_and_delete_rows_and_columns. End of function.')
