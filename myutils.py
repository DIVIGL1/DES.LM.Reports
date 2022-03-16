import datetime
import os
import pickle

import win32com.client

import myconstants
import mytablefuncs


def save_param(param_name, param_value, filename="saved.pkl"):
    params_dict = get_all_params(filename="saved.pkl")
    params_dict[param_name] = param_value
    
    with open(filename, 'wb') as file_handle:
        pickle.dump(params_dict, file_handle)

def get_all_params(filename="saved.pkl"):
    if not os.path.isfile(filename):
        with open(filename, 'wb') as file_handle:
            pickle.dump({"<ДатаВремя создания>": datetime.datetime.now()}, file_handle)

    with open(filename, "rb") as file_handle:
        params_dict = pickle.load(file_handle)

    return(params_dict)

def load_param(param_name, default="", filename="saved.pkl"):
    params_dict = get_all_params(filename="saved.pkl")
    
    return(params_dict.get(param_name, default))

def get_files_list(path2files="", files_starts="", files_ends=".xlsx", reverse=True):
    path2files = os.path.join(os.getcwd(), path2files)
    files_list = \
        [one_file[len(files_starts):][:-len(files_ends)] \
            for one_file \
                in os.listdir(path2files) \
                    if (one_file.lower().startswith(files_starts.lower()) and \
                        one_file.lower().endswith(files_ends.lower()))]
    
    files_list = sorted(files_list, reverse=reverse)
    return(files_list)

def is_all_parametars_exist():
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
    ret_value = ""
    for one_file_info in files_list:
        if not os.path.isfile(s_section_path + "/" + one_file_info[0]):
            ret_value = one_file_info
            break
        
    return ret_value

def get_report_parameters():
    myconstants.ROUND_FTE_VALUE = mytablefuncs.get_parameter_value(myconstants.ROUND_FTE_SECTION_NAME, myconstants.ROUND_FTE_DEFVALUE)
    myconstants.MEANOURSPERMONTH_VALUE = mytablefuncs.get_parameter_value(myconstants.MEANOURSPERMONTH_SECTION_NAME, myconstants.MEANOURSPERMONTH_DEFVALUE)
    s_preff = myconstants.DO_IT_PREFFIX

    p_delete_not_prod_units =\
        load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_NONPROD, myconstants.PARAMETER_SAVED_VALUE_DELETE_NONPROD_DEFVALUE)
    p_delete_pers_data =\
        load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_PERSDATA, myconstants.PARAMETER_SAVED_VALUE_DELETE_PERSDATA_DEFVALUE)
    p_delete_vacation =\
        load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_VAC, myconstants.PARAMETER_SAVED_VALUE_DELETE_VAC_DEFVALUE)
    p_add_vfte =\
        load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_ADD_VFTE, myconstants.PARAMETER_SAVED_VALUE_ADD_VFTE_DEFVALUE)
    p_save_without_formulas =\
        load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS, myconstants.PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS_DEFVALUE)
    p_delete_rawdata_sheet =\
        load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DEL_RAWSHEET, myconstants.PARAMETER_SAVED_VALUE_DEL_RAWSHEET_DEFVALUE)
    p_open_in_excel =\
        load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL, myconstants.PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL_DEFVALUE)

    if os.path.isfile(myconstants.SECRET_COSTS_LOCATION + "/" + myconstants.COSTS_TABLE):
        p_save_without_formulas = True
        p_delete_rawdata_sheet = True
    
    return p_delete_not_prod_units, p_delete_pers_data, p_delete_vacation, p_add_vfte, p_save_without_formulas, p_delete_rawdata_sheet, p_open_in_excel

def get_full_files_names(raw_file_name, report_file_name):
    report_prepared_name = \
        os.path.join( 
            os.path.join(os.getcwd(), mytablefuncs.get_parameter_value(myconstants.REPORTS_PREPARED_SECTION_NAME)),
            raw_file_name + "__" + report_file_name + myconstants.EXCEL_FILES_ENDS
        )
    report_prepared_name = report_prepared_name.replace("\\", "/")
    
    report_file_name = \
        os.path.join( 
            os.path.join(os.getcwd(), mytablefuncs.get_parameter_value(myconstants.REPORTS_SECTION_NAME)),
            myconstants.REPORT_FILE_PREFFIX + report_file_name + myconstants.EXCEL_FILES_ENDS
        )
    report_file_name = report_file_name.replace("\\", "/")
    
    raw_file_name = \
        os.path.join( 
            os.path.join(os.getcwd(), mytablefuncs.get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)),
            raw_file_name + myconstants.EXCEL_FILES_ENDS
        )
    raw_file_name = raw_file_name.replace("\\", "/")
    
    return raw_file_name, report_file_name, report_prepared_name
    
def get_excel_and_wb(excel_file_name):
    oExcel = win32com.client.Dispatch("Excel.Application")
    oExcel.Visible = oExcel.WorkBooks.Count > 0
    oExcel.DisplayAlerts = False
        
    wb = oExcel.Workbooks.Open(excel_file_name)
    currwindow = oExcel.ActiveWindow
    currwindow.WindowState = myconstants.EXCELWINDOWSTATE_MIN
    n_save_excel_calc_status = oExcel.Calculation
    oExcel.Calculation = myconstants.EXCEL_MANUAL_CALC
    
    return oExcel, currwindow, wb, n_save_excel_calc_status

def get_sheets_list(wb):
    return([one_sheet.Name for one_sheet in wb.Sheets])

def save_report(wb, p_save_without_formulas, p_delete_rawdata_sheet):
    wb.Save()
    if p_save_without_formulas: 
        for curr_sheet_name in get_sheets_list(wb):
            if curr_sheet_name not in myconstants.SHEETS_DONT_DELETE_FORMULAS:
                wb.Sheets[curr_sheet_name].UsedRange.Value = wb.Sheets[curr_sheet_name].UsedRange.Value
        
        if p_delete_rawdata_sheet:
            for one_sheet_name in myconstants.DELETE_SHEETS_LIST_IF_NO_FORMULAS:
                if one_sheet_name in get_sheets_list(wb):
                    wb.Sheets[one_sheet_name].Delete()
        
        wb.Save()


def hide_and_delete_rows_and_columns(oExcel, wb):
    # -----------------------------------
    oExcel.Calculation = myconstants.EXCEL_AUTOMATIC_CALC
    oExcel.Calculation = myconstants.EXCEL_MANUAL_CALC
    for curr_sheet_name in get_sheets_list(wb):
        if curr_sheet_name not in myconstants.SHEETS_DONT_DELETE_FORMULAS:
            row_counter = 0
            first_row_with_del = 0
            last_row_with_del = 0
            p_found_first_row = False
            last_row_4_test = myconstants.PARAMETER_MAX_ROWS_TEST_IN_REPORT
            range_from_excel = wb.Sheets[curr_sheet_name].Range(wb.Sheets[curr_sheet_name].Cells(1, 1), wb.Sheets[curr_sheet_name].Cells(last_row_4_test, 1)).Value

            # Ищем первый признак 'delete'
            for row_counter in range(len(range_from_excel)):
                row_del_flag_value = range_from_excel[row_counter][0]
                if row_del_flag_value is None:
                    p_found_first_row = False
                    break
                
                if (type(row_del_flag_value) == str):
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

                wb.Sheets[curr_sheet_name].Range(wb.Sheets[curr_sheet_name].Cells( \
                    first_row_with_del, 1), wb.Sheets[curr_sheet_name].Cells(last_row_with_del, 1)).Rows.EntireRow.Delete()
    # -----------------------------------
            # Скрываем строки и столбцы с признаком 'hide'
            for curr_sheet_name in get_sheets_list(wb):
                if curr_sheet_name not in [myconstants.RAW_DATA_SHEET_NAME, myconstants.UNIQE_LISTS_SHEET_NAME, myconstants.SETTINGS_SHEET_NAME]:
                    # Скрываем строки с признаком 'hide'
                    for curr_row in range(1, myconstants.NUM_ROWS_FOR_HIDE + 1):
                        cell_value = wb.Sheets[curr_sheet_name].Cells(curr_row, 1).Value
                        if type(cell_value) == str and cell_value is not None and cell_value.replace(" ", "") == myconstants.HIDE_MARKER:
                            pass
                            wb.Sheets[curr_sheet_name].Rows(curr_row).Hidden = True
                    # Скрываем столбцы с признаком 'hide'
                    for curr_col in range(1, myconstants.NUM_COLUMNS_FOR_HIDE + 1):
                        cell_value = wb.Sheets[curr_sheet_name].Cells(1, curr_col).Value
                        if type(cell_value) == str and cell_value is not None and cell_value.replace(" ", "") == myconstants.HIDE_MARKER:
                            wb.Sheets[curr_sheet_name].Columns(curr_col).Hidden = True
                        else:
                            pass
    # -----------------------------------
   

if __name__ == "__main__":
    #    param_name = myconstants.LAST_REPORT_PARAM_NAME
    #    print(load_param(param_name, "<пусто>"))
    #    save_param(param_name, 7)
    #    print(load_param(param_name, "<пусто>"))
    print(get_files_list("RawData", "", ".xlsx"))
