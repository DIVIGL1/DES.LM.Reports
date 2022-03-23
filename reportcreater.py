import shutil
import threading
import time
import warnings

import pythoncom

import myconstants
from mytablefuncs import get_parameter_value, prepare_data
import myutils

def thread(my_func):
    """
    Запускает функцию в отдельном потоке
    """
    def wrapper(*args, **kwargs):
        my_thread = threading.Thread(target=my_func, args=args, kwargs=kwargs)
        my_thread.start()
    return wrapper

class ReportCreater(object):
    def __init__(self, parent, *args):
        super(ReportCreater, self).__init__(*args)
        self.parent = parent
        self.reports_list = \
            myutils.get_files_list(
                get_parameter_value(myconstants.REPORTS_SECTION_NAME), 
                myconstants.REPORT_FILE_PREFFIX, ".xlsx", reverse=False
            )
        warnings.filterwarnings("ignore")
    
    def get_reports_list(self):
        return(self.reports_list)
    
    def get_report_file_name_by_num(self, num):
        return(self.reports_list[num])
    
    def create_report(self):
        if self.parent.report_parameters.report_file_name is None:
            myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
            self.parent._mainwindow.ui.set_status("Необходимо выбрать отчётную форму.")
            self.parent._mainwindow.ui.enable_buttons()
            return False
        if self.parent.report_parameters.raw_file_name is None:
            myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
            self.parent._mainwindow.ui.set_status("Необходимо выбрать файл, выгруженный из DES.LM для формирования отчёта.")
            self.parent._mainwindow.ui.enable_buttons()
            return False
        
        self.start_timer()
        self.parent._mainwindow.ui.clear_status()
        send_df_2_xls(self.parent.report_parameters)
        self.show_timer()
        
        return True

    def start_timer(self):
        self.start_prog_time = time.time()

    def show_timer(self):
        end_prog_time = time.time()
        duration_in_seconds = int(end_prog_time - self.start_prog_time)
        self.parent._mainwindow.ui.set_status("Время выполнения: {0:0>2}:{1:0>2}".format(duration_in_seconds // 60, duration_in_seconds % 60))

@thread
def send_df_2_xls(report_parameters):
    # ----------------------------------------------------
    raw_file_name = report_parameters.raw_file_name
    report_file_name = report_parameters.report_file_name
    report_prepared_name = report_parameters.report_prepared_name

    p_delete_not_prod_units = report_parameters.p_delete_not_prod_units
    p_delete_pers_data = report_parameters.p_delete_pers_data
    p_delete_vacation = report_parameters.p_delete_vacation
    p_virtual_FTE = report_parameters.p_virtual_FTE
    p_save_without_formulas = report_parameters.p_save_without_formulas
    p_delete_rawdata_sheet = report_parameters.p_delete_rawdata_sheet
    p_open_in_excel = report_parameters.p_open_in_excel

    ui_handle = report_parameters.parent._mainwindow.ui
    # ----------------------------------------------------
    
    ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
    ui_handle.set_status(f"1. Выбран отчет:\n>>   {report_file_name}")
    ui_handle.set_status(f"2. Выбран файл с данными:\n>>   {raw_file_name}")
    ui_handle.set_status(f"3. {'В отчете оставитьтолько производство.' if p_delete_not_prod_units else 'В отчет включить и производство и управленцев.'}")
    ui_handle.set_status(f"4. Персональные данные (больничные листы): {'исключить из отчета.' if p_delete_pers_data else 'оставить в отчете.'}")
    ui_handle.set_status(f"5. Вакансии: {'удалить из отчета.' if p_delete_vacation else 'оставить в отчете.'}")
    ui_handle.set_status(f"6. Искусственно добавить FTE: {'да (если есть).' if p_virtual_FTE else 'нет.'}")
    if p_save_without_formulas:
        ui_handle.set_status("7. Все формулы заменить их значениями (быстрее открывается и меньше размер файла).")
        if p_delete_rawdata_sheet:
            ui_handle.set_status("8. Удалить закладки с данными (дополнительное уменьшение размера файла).")
    else:
        ui_handle.set_status("7. Сохранть формулы (возможно медленное открытие и больше размер файла).")
            
    if p_save_without_formulas & p_delete_rawdata_sheet:
        ui_handle.set_status(f"9. Округление до: {myconstants.ROUND_FTE_VALUE}-го знака после запятой")
    else:
        ui_handle.set_status(f"8. Округление до: {myconstants.ROUND_FTE_VALUE}-го знака после запятой")
            
    ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
    
    ui_handle.set_status("Проверим структуру файла, содержащего форму отчёта.")

    try:
        shutil.copyfile(report_file_name, report_prepared_name)
    except (OSError, shutil.Error):
        ui_handle.set_status("Не удалось скопировать файл с формой отчёта.")
        ui_handle.set_status(f"проверьте, пожалуйста, не открыт ли у Вас файл: \n>>   {report_prepared_name}")
        ui_handle.set_status("Формирование отчёта не возможно.")
        myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.enable_buttons()
        return

    pythoncom.CoInitializeEx(0)
    
    # необходимо проверить запуск Excel - здесь может подвисать #
    oExcel, currwindow, wb, n_save_excel_calc_status = myutils.get_excel_and_wb(report_prepared_name)
    
    if (myconstants.RAW_DATA_SHEET_NAME not in myutils.get_sheets_list(wb)):
        ui_handle.set_status("")
        ui_handle.set_status("")
        ui_handle.set_status("[Ошибка в структуре отчета]")
        ui_handle.set_status("")
        ui_handle.set_status("В файле для выбранной формы отчёта отсутствует необходимый лист для данных.")
        ui_handle.set_status("Формирование отчёта не возможно.")
        myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.enable_buttons()
        oExcel.Calculation = n_save_excel_calc_status
        oExcel.DisplayAlerts = True
        wb.Close()
        return
    elif (myconstants.UNIQE_LISTS_SHEET_NAME not in myutils.get_sheets_list(wb)):
        ui_handle.set_status("")
        ui_handle.set_status("")
        ui_handle.set_status("[Ошибка в структуре отчета]")
        ui_handle.set_status("")
        ui_handle.set_status("В файле для выбранной формы отчёта отсутствует необходимый лист для уникальных списков.")
        ui_handle.set_status("Формирование отчёта не возможно.")
        myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.enable_buttons()
        oExcel.Calculation = n_save_excel_calc_status
        oExcel.DisplayAlerts = True
        wb.Close()
        return
    elif (myconstants.SETTINGS_SHEET_NAME not in myutils.get_sheets_list(wb)):
        ui_handle.set_status("")
        ui_handle.set_status("")
        ui_handle.set_status("[Ошибка в структуре отчета]")
        ui_handle.set_status("")
        ui_handle.set_status("В файле для выбранной формы отчёта отсутствует необходимый лист c настройками.")
        ui_handle.set_status("Формирование отчёта не возможно.")
        myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.enable_buttons()
        oExcel.Calculation = n_save_excel_calc_status
        oExcel.DisplayAlerts = True
        wb.Close()
        return
    else:
        ui_handle.change_last_status_line("Пройдена проверка структуры файла, содержащего форму отчёта.")

    ui_handle.set_status("Файл Excel с формой отчёта подгружен.")
    
    report_df = prepare_data(raw_file_name, p_delete_not_prod_units, p_delete_pers_data, p_delete_vacation, ui_handle)
    ui_handle.set_status(f"Таблица для загрузки полностью подготовлена (всего строк данных: {report_df.shape[0]})")

    ui_handle.set_status("Начинаем перенос строк в Excel:")
    data_sheet = wb.Sheets[myconstants.RAW_DATA_SHEET_NAME]
    ulist_sheet = wb.Sheets[myconstants.UNIQE_LISTS_SHEET_NAME]

    data_array = report_df.to_numpy()
    data_sheet.Range(data_sheet.Cells(2, 1), data_sheet.Cells(len(data_array) + 1, len(data_array[0]))).Value = data_array    
    
    ui_handle.set_status("Строки в Excel скопированы.")

    ui_handle.set_status("Формируем списки с уникальными значениями.")
    # Запоним списки уникальными значениями
    column = 1
    values_dict = dict()
    while True:
        uniq_col_name = ulist_sheet.Cells(1, column).value
        if (type(uniq_col_name) != str) and (uniq_col_name is not None):
            ui_handle.set_status("")
            ui_handle.set_status("")
            ui_handle.set_status("[Ошибка в структуре отчета]")
            ui_handle.set_status("")
            ui_handle.set_status("В файле для выбранной формы на листе для уникальных списков в стороке 1:1 " + \
                                    "в качестве наименований списков должны быть символьные значения. Формирование отчёта остановлено.")
            myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
            ui_handle.enable_buttons()
            oExcel.DisplayAlerts = True
            wb.Close()
            return
            
        if uniq_col_name is not None and uniq_col_name.replace(" ", "") != "":
            values_dict[uniq_col_name] = column
            column += 1
        else:
            break

    full_column_list = report_df.columns.tolist()
    columns_4_unique_list = [column for column in values_dict.keys() if column in full_column_list]
    if len(columns_4_unique_list) == 0:
        ui_handle.set_status("")
        ui_handle.set_status("")
        ui_handle.set_status("[Ошибка в структуре отчета]")
        ui_handle.set_status("")
        ui_handle.set_status("В файле для выбранной формы на листе для уникальных списков не указан " + \
                                "уникальный список из возможного перечня. Формирование отчёта остановлено.")
        myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.enable_buttons()
        oExcel.DisplayAlerts = True
        wb.Close()
        return

    ui_handle.set_status(f"Всего списков c уникальными данными: {len(columns_4_unique_list)} шт.")
    ui_handle.set_status("")
    for index, one_column in enumerate(columns_4_unique_list):
        unique_elements_list = sorted(report_df[one_column].unique())
        ui_handle.change_last_status_line(f"{index + 1} из {len(columns_4_unique_list)} ({one_column}): {len(unique_elements_list)}")

        data_array = [[one_uelement] for one_uelement in unique_elements_list]
        ulist_sheet.Range(
            ulist_sheet.Cells(2, values_dict[one_column]), ulist_sheet.Cells(len(data_array) + 1, values_dict[one_column])\
        ).Value = data_array
        
    if len(columns_4_unique_list) == 1:
        ui_handle.change_last_status_line(f"Значений в списке {len(unique_elements_list)} шт.")
    else:
        ui_handle.change_last_status_line("Собраны и сохранениы списки с уникальными значениями.")
                
    ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
    ui_handle.set_status(f"Сохраняем в файл: {report_prepared_name}")
    
    myutils.hide_and_delete_rows_and_columns(oExcel, wb)

    oExcel.Calculation = n_save_excel_calc_status
    myutils.save_report(wb, p_save_without_formulas, p_delete_rawdata_sheet)
    ui_handle.set_status("Отчёт сохранён.")
    
    myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, report_prepared_name)
    
    if p_open_in_excel:
        # Скроем вспомогательные листы
        if myconstants.UNIQE_LISTS_SHEET_NAME in myutils.get_sheets_list(wb):
            wb.Sheets[myconstants.UNIQE_LISTS_SHEET_NAME].Visible = False
        
        if myconstants.SETTINGS_SHEET_NAME in myutils.get_sheets_list(wb):
            wb.Sheets[myconstants.SETTINGS_SHEET_NAME].Visible = False
            
        oExcel.Visible = True
        currwindow.WindowState = myconstants.EXCELWINDOWSTATE_MAX
    else:
        oExcel.DisplayAlerts = True
        wb.Close()

    ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
    ui_handle.enable_buttons()


if __name__ == "__main__":
    print(myutils.get_files_list(get_parameter_value(myconstants.REPORTS_SECTION_NAME)))
