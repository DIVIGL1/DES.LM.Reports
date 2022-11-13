import shutil
import threading
import time
import warnings

import pythoncom

import myconstants
from mytablefuncs import get_parameter_value, prepare_data, test_secret_files_list
import myutils
from myexcelclass import MyExcel


def thread(my_func):
    """
    Запускает функцию в отдельном потоке
    """
    def wrapper(*args, **kwargs):
        my_thread = threading.Thread(target=my_func, args=args, kwargs=kwargs)
        my_thread.start()
    return wrapper


class reportCreator(object):
    def __init__(self, parent, *args):
        super(reportCreator, self).__init__(*args)
        self.parent = parent
        self.reports_list = \
            myutils.get_files_list(
                get_parameter_value(myconstants.REPORTS_SECTION_NAME), 
                myconstants.REPORT_FILE_PREFFIX, ".xlsx", reverse=False
            )
        warnings.filterwarnings("ignore")
        self.start_prog_time = None
    
    def get_reports_list(self):
        return self.reports_list
    
    def get_report_file_name_by_num(self, num):
        return self.reports_list[num]
    
    def create_report(self):
        if self.parent.report_parameters.report_file_name is None:
            myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
            self.parent.mainwindow.ui.set_status("Необходимо выбрать отчётную форму.")
            self.parent.mainwindow.ui.enable_buttons()
            return False
        if self.parent.report_parameters.raw_file_name is None:
            myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
            self.parent.mainwindow.ui.set_status("Необходимо выбрать файл, выгруженный из DES.LM для формирования отчёта.")
            self.parent.mainwindow.ui.enable_buttons()
            return False
        
        self.start_timer()
        self.parent.mainwindow.ui.clear_status()

        send_df_2_xls(self.parent.report_parameters)

        return True

    def start_timer(self):
        self.start_prog_time = time.time()

    def show_timer(self):
        end_prog_time = time.time()
        duration_in_seconds = int(end_prog_time - self.start_prog_time)
        self.parent.mainwindow.ui.set_status("Время выполнения: {0:0>2}:{1:0>2}".format(duration_in_seconds // 60, duration_in_seconds % 60))


@thread
def send_df_2_xls(report_parameters):
    # Определим ключевые параметры в переменные,
    # чтобы не указывать полную ссылку каждый раз
    raw_file_name = report_parameters.raw_file_name
    report_file_name = report_parameters.report_file_name
    report_prepared_name = report_parameters.report_prepared_name

    p_delete_vip = report_parameters.p_delete_vip
    p_delete_not_prod_units = report_parameters.p_delete_not_prod_units
    p_projects_with_add_info = report_parameters.p_projects_with_add_info
    p_delete_without_fact = report_parameters.p_delete_without_fact
    p_curr_month_half = report_parameters.p_curr_month_half
    p_delete_pers_data = report_parameters.p_delete_pers_data
    p_delete_vacation = report_parameters.p_delete_vacation
    p_virtual_FTE = report_parameters.p_virtual_FTE
    p_save_without_formulas = report_parameters.p_save_without_formulas
    p_delete_rawdata_sheet = report_parameters.p_delete_rawdata_sheet

    ui_handle = report_parameters.parent.mainwindow.ui
    # ----------------------------------------------------
    num_poz = 1
    ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
    ui_handle.set_status(f"{num_poz}. Рабочий каталог: {myutils.get_home_dir()}")
    num_poz += 1
    ui_handle.set_status(f"{num_poz}. Выбран отчет: {myutils.rel_path(report_file_name)}")
    num_poz += 1
    ui_handle.set_status(f"{num_poz}. Выбран файл с данными: {myutils.rel_path(raw_file_name)}")
    num_poz += 1
    if p_delete_vip:
        ui_handle.set_status(f"{num_poz}. Данные о VIP удалены.")
        num_poz += 1
    if p_curr_month_half:
        ui_handle.set_status(f"{num_poz}. В текущем месяце для расчета FTE использована половина нормы.")
        num_poz += 1
    if p_delete_without_fact:
        ui_handle.set_status(f"{num_poz}. Исключены строки с фактом равным нулю.")
        num_poz += 1
    ui_handle.set_status(f"{num_poz}. {'В отчете включено только производство.' if p_delete_not_prod_units else 'В отчет включены производство и управленцы.'}")
    num_poz += 1
    ui_handle.set_status(
        f"{num_poz}. Обрабатываются {'только проекты из списка с доп данными.' if p_projects_with_add_info else 'все проекты.'}")
    num_poz += 1
    ui_handle.set_status(f"{num_poz}. Персональные данные (больничные листы): {'исключены из отчета.' if p_delete_pers_data else 'оставлены в отчете.'}")
    num_poz += 1
    ui_handle.set_status(f"{num_poz}. Вакансии: {'удалены из отчета.' if p_delete_vacation else 'оставлены в отчете.'}")
    num_poz += 1
    ui_handle.set_status(f"{num_poz}. Искусственно добавлены FTE: {'да (если есть).' if p_virtual_FTE else 'нет.'}")
    num_poz += 1
    if p_save_without_formulas:
        ui_handle.set_status(f"{num_poz}. Все формулы заменены их значениями (быстрее открывается и меньше размер файла).")
        num_poz += 1
        if p_delete_rawdata_sheet:
            ui_handle.set_status(f"{num_poz}. Удалены закладки с данными (дополнительное уменьшение размера файла).")
            num_poz += 1
    else:
        ui_handle.set_status(f"{num_poz}. Сохранены формулы (возможно медленное открытие и больше размер файла).")
        num_poz += 1
            
    ui_handle.set_status(f"{num_poz}. Округление до: {myconstants.ROUND_FTE_VALUE}-го знака после запятой")
    num_poz += 1
            
    # В случае наличия в специальной парке файлов с секретными данными и если
    # в них есть какие-то не нулевые значения, то нужно удалять лист с данными:
    test_result = test_secret_files_list()
    if test_result:
        ui_handle.set_status(test_result)
        p_save_without_formulas = True
        report_parameters.p_save_without_formulas = p_save_without_formulas
        p_delete_rawdata_sheet = True
        report_parameters.p_delete_rawdata_sheet = p_delete_rawdata_sheet
    else:
        ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)

    ui_handle.set_status("Проверим структуру файла, содержащего форму отчёта.")

    try:
        shutil.copyfile(report_file_name, report_prepared_name)
    except (OSError, shutil.Error):
        ui_handle.set_status("Не удалось скопировать файл с формой отчёта.")
        ui_handle.set_status(f"проверьте, пожалуйста, не открыт ли у Вас файл:")
        ui_handle.set_status(f"{myutils.rel_path(report_prepared_name)}")
        ui_handle.set_status("Формирование отчёта невозможно.")
        myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")
        ui_handle.enable_buttons()
        return False

    pythoncom.CoInitializeEx(0)

    report_df = prepare_data(
        raw_file_name,
        p_delete_vip,
        p_delete_not_prod_units,
        p_projects_with_add_info,
        p_delete_without_fact,
        p_curr_month_half,
        p_delete_pers_data,
        p_delete_vacation,
        p_virtual_FTE,
        ui_handle
    )
    if report_df is None:
        return False

    ui_handle.set_status(f"Таблица для загрузки полностью подготовлена (всего строк данных: {report_df.shape[0]})")

    oexcel = MyExcel(report_parameters)
    ret_value = None
    if oexcel.not_ready:
        # Что-то пошло не так.
        ret_value = False
    else:
        ui_handle.set_status("Начинаем перенос строк в Excel:")
        data_sheet = oexcel.work_book.Sheets[myconstants.RAW_DATA_SHEET_NAME]
        ulist_sheet = oexcel.work_book.Sheets[myconstants.UNIQE_LISTS_SHEET_NAME]

        data_array = report_df.to_numpy()
        data_sheet.Range(data_sheet.Cells(2, 1), data_sheet.Cells(len(data_array) + 1, len(data_array[0]))).Value = data_array

        ui_handle.set_status("Строки в Excel скопированы.")

        ui_handle.set_status("Формируем списки с уникальными значениями.")

        # Заполним списки уникальными значениями
        column = 1
        values_dict = dict()
        while True:
            uniq_col_name = ulist_sheet.Cells(1, column).value
            if (type(uniq_col_name) != str) and (uniq_col_name is not None):
                ui_handle.set_status("")
                ui_handle.set_status("")
                ui_handle.set_status("[Ошибка в структуре отчета]")
                ui_handle.set_status("")
                ui_handle.set_status("В файле для выбранной формы на листе для уникальных списков в строке 1:1 " +
                                        "в качестве наименований списков должны быть символьные значения. Формирование отчёта остановлено.")

                ret_value = False
                break

            if uniq_col_name is not None and uniq_col_name.replace(" ", "") != "":
                values_dict[uniq_col_name] = column
                column += 1
            else:
                break

        if ret_value is None:
            full_column_list = report_df.columns.tolist()
            columns_4_unique_list = [column for column in values_dict.keys() if column in full_column_list]
            if len(columns_4_unique_list) == 0:
                ui_handle.set_status("")
                ui_handle.set_status("")
                ui_handle.set_status("[Ошибка в структуре отчета]")
                ui_handle.set_status("")
                ui_handle.set_status("В файле для выбранной формы на листе для уникальных списков не указан " +
                                        "уникальный список из возможного перечня. Формирование отчёта остановлено.")
                myutils.save_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "")

                ret_value = False
            else:
                ui_handle.set_status(f"Всего списков c уникальными данными: {len(columns_4_unique_list)} шт.")
                ui_handle.set_status("")
                unique_elements_list = None
                for index, one_column in enumerate(columns_4_unique_list):
                    unique_elements_list = sorted(report_df[one_column].unique())
                    ui_handle.change_last_status_line(f"{index + 1} из {len(columns_4_unique_list)} ({one_column}): {len(unique_elements_list)}")

                    data_array = [[one_uelement] for one_uelement in unique_elements_list]
                    ulist_sheet.Range(
                        ulist_sheet.Cells(2, values_dict[one_column]), ulist_sheet.Cells(len(data_array) + 1, values_dict[one_column])
                    ).Value = data_array

                if len(columns_4_unique_list) == 1:
                    ui_handle.change_last_status_line(f"Значений в списке {len(unique_elements_list)} шт.")
                else:
                    ui_handle.change_last_status_line("Собраны и сохранены списки с уникальными значениями.")

                oexcel.report_prepared = True

    del oexcel

    return ret_value


if __name__ == "__main__":
    print(myutils.get_files_list(get_parameter_value(myconstants.REPORTS_SECTION_NAME)))
