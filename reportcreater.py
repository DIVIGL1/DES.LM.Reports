import os
import shutil
import threading
import time
import warnings

import pythoncom

import myconstants
#from myutils import get_files_list, save_param, get_report_parameters, get_full_files_names
#from myutils import get_excel_and_wb, save_report, hide_and_delete_rows_and_columns
from myutils import *
from mytablefuncs import get_parameter_value, prepare_data


def thread(my_func):
    """
    Запускает функцию в отдельном потоке
    """
    def wrapper(*args, **kwargs):
        my_thread = threading.Thread(target=my_func, args=args, kwargs=kwargs)
        my_thread.start()
    return wrapper

class ReportCreater(object):
    def __init__(self, *args):
        super(ReportCreater, self).__init__(*args)
        self.reports_list = get_files_list(\
                                get_parameter_value(myconstants.REPORTS_SECTION_NAME),\
                                myconstants.REPORT_FILE_PREFFIX,\
                                ".xlsx",\
                                reverse=False\
                                        )
        warnings.filterwarnings("ignore")
    
    def get_reports_list(self):
        return(self.reports_list)
    
    def get_report_file_name_by_num(self, num):
        return(self.reports_list[num])


@thread
def send_df_2_xls(report_file_name, raw_file_name, ui_handle):
    save_param(myconstants.PARAMETER_SAVED_SELECTED_REPORT,ui_handle.reports_list.index(report_file_name) + 1)
    if report_file_name == None:
        save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.set_status("Необходимо выбрать отчётную форму.")
        ui_handle.enable_buttons()
        return
    if raw_file_name == None:
        save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.set_status("Необходимо выбрать файл, выгруженный из DES.LM для формирования отчёта.")
        ui_handle.enable_buttons()
        return
    
    start_prog_time = time.time()
    p_delete_vacation, _, p_save_without_formulas, p_open_in_excel = get_report_parameters()
    raw_file_name, report_file_name, report_prepared_name = get_full_files_names(raw_file_name, report_file_name)
    
    ui_handle.clear_status()
    ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
    ui_handle.set_status(f"1. Выбран отчет:\n>>   {report_file_name}")
    ui_handle.set_status(f"2. Выбран файл с данными:\n>>   {raw_file_name}")
    ui_handle.set_status(f"3. Вакансии: {'удалить из отчета.' if p_delete_vacation else 'оставить в отчете.'}")
    ui_handle.set_status(f"4. Округление до: {myconstants.ROUND_FTE_VALUE}-го знака после запятой")
    ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
    ui_handle.set_status("Проверим структуру файла, содержащего форму отчёта.")

    try:
        shutil.copyfile(report_file_name, report_prepared_name)
    except:
        ui_handle.set_status("Не удалось скопировать файл с формой отчёта.")
        ui_handle.set_status(f"проверьте, пожалуйста. не открыт ли у Вас файл: \n>>   {report_prepared_name}")
        ui_handle.set_status("Формирование отчёта не возможно.")
        save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.enable_buttons()
        return

    pythoncom.CoInitializeEx(0)
    
    oExcel, currwindow, wb, n_save_excel_calc_status = get_excel_and_wb(report_prepared_name)
    
    if (myconstants.RAW_DATA_SHEET_NAME not in [one_sheet.Name for one_sheet in wb.Sheets]):
        ui_handle.set_status("")
        ui_handle.set_status("")
        ui_handle.set_status("[Ошибка в структуре отчета]")
        ui_handle.set_status("")
        ui_handle.set_status("В файле для выбранной формы отчёта отсутствует необходимый лист для данных.")
        ui_handle.set_status("Формирование отчёта не возможно.")
        save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.enable_buttons()
        oExcel.Calculation = n_save_excel_calc_status
        oExcel.Quit()
        return
    elif (myconstants.UNIQE_LISTS_SHEET_NAME not in [one_sheet.Name for one_sheet in wb.Sheets]):
        ui_handle.set_status("")
        ui_handle.set_status("")
        ui_handle.set_status("[Ошибка в структуре отчета]")
        ui_handle.set_status("")
        ui_handle.set_status("В файле для выбранной формы отчёта отсутствует необходимый лист для уникальных списков.")
        ui_handle.set_status("Формирование отчёта не возможно.")
        save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.enable_buttons()
        oExcel.Calculation = n_save_excel_calc_status
        oExcel.Quit()
        return
    elif (myconstants.SETTINGS_SHEET_NAME not in [one_sheet.Name for one_sheet in wb.Sheets]):
        ui_handle.set_status("")
        ui_handle.set_status("")
        ui_handle.set_status("[Ошибка в структуре отчета]")
        ui_handle.set_status("")
        ui_handle.set_status("В файле для выбранной формы отчёта отсутствует необходимый лист c настройками.")
        ui_handle.set_status("Формирование отчёта не возможно.")
        save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.enable_buttons()
        oExcel.Calculation = n_save_excel_calc_status
        oExcel.Quit()
        return
    else:
        ui_handle.set_status("Ошибок не найдено.")

    ui_handle.set_status("Файл Excel с формой отчёта подгружен.")
    
    report_df = prepare_data(raw_file_name, p_delete_vacation, ui_handle)
    ui_handle.set_status(f"Таблица для загрузки полностью подготовлена (всего строк данных: {report_df.shape[0]})")

    ui_handle.set_status("Начинаем перенос строк в Excel:")
    data_sheet = wb.Sheets[myconstants.RAW_DATA_SHEET_NAME]
    ulist_sheet = wb.Sheets[myconstants.UNIQE_LISTS_SHEET_NAME]

    data_array = report_df.to_numpy()
    data_sheet.Range(data_sheet.Cells(2,1), data_sheet.Cells(len(data_array)+1, len(data_array[0]))).Value = data_array    
    
    ui_handle.set_status("Строки в Excel скопированы.")

    ui_handle.set_status("Формируем списки с уникальными значениями.")
    # Запоним списки уникальными значениями
    column = 1
    values_dict = dict()
    while True:
        uniq_col_name = ulist_sheet.Cells(1, column).value
        if uniq_col_name != None and uniq_col_name.replace(" ","") != "":
            values_dict[uniq_col_name] = column
            column += 1
        else:
            break

    full_column_list = report_df.columns.tolist()
    columns_4_unique_list = [column for column in values_dict.keys() if column in full_column_list]
    if len(columns_4_unique_list)==0:
        ui_handle.set_status("")
        ui_handle.set_status("")
        ui_handle.set_status("[Ошибка в структуре отчета]")
        ui_handle.set_status("")
        ui_handle.set_status("В файле для выбранной формы на листе для уникальных списков не указано ничего.")
        ui_handle.set_status("Формирование отчёта остановлено.")
        save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.enable_buttons()
        oExcel.Quit()
        return

    ui_handle.set_status(f"Всего списков c уникальными данными: {len(columns_4_unique_list)} шт.")
    ui_handle.set_status(f"")
    for index, one_column in enumerate(columns_4_unique_list):
        unique_elements_list = sorted(report_df[one_column].unique())
        ui_handle.change_last_status_line(f"{index + 1} из {len(columns_4_unique_list)} ({one_column}): {len(unique_elements_list)}")

        data_array = [[one_uelement] for one_uelement in unique_elements_list]
        ulist_sheet.Range(\
            ulist_sheet.Cells(2, values_dict[one_column]), ulist_sheet.Cells(len(data_array) + 1, values_dict[one_column])\
                        ).Value = data_array
        
    if len(columns_4_unique_list)==1:
        ui_handle.change_last_status_line(f"Значений в списке {len(unique_elements_list)} шт.")
    else:
        ui_handle.change_last_status_line("Собраны и сохранениы списки с уникальными значениями.")
                
    ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
    ui_handle.set_status(f"Сохраняем в файл: {report_prepared_name}")
    
    hide_and_delete_rows_and_columns(oExcel, wb)

    oExcel.Calculation = n_save_excel_calc_status
    save_report(wb, p_save_without_formulas)
    ui_handle.set_status(f"Отчёт сохранён.")
    
    save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, report_prepared_name)
    
    end_prog_time = time.time()
    duration_in_seconds = int(end_prog_time - start_prog_time)
    ui_handle.set_status("Время выполнения: {0:0>2}:{1:0>2}".format(duration_in_seconds//60,duration_in_seconds%60))

    if p_open_in_excel:
        # Скроем вспомогательные листы
        wb.Sheets[myconstants.UNIQE_LISTS_SHEET_NAME]. Visible = False
        wb.Sheets[myconstants.SETTINGS_SHEET_NAME]. Visible = False
            
        oExcel.Visible = True
        currwindow.WindowState = myconstants.EXCELWINDOWSTATE_MAX
    else:
        pass
        oExcel.Quit()

    ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
    ui_handle.enable_buttons()







if __name__ == "__main__":
    print(get_files_list(get_parameter_value(myconstants.REPORTS_SECTION_NAME)))
