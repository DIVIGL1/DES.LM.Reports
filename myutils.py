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

    return params_dict


def load_param(param_name, default="", filename="saved.pkl"):
    params_dict = get_all_params(filename=filename)
    
    return params_dict.get(param_name, default)


def get_files_list(path2files="", files_starts="", files_ends=".xlsx", reverse=True):
    path2files = os.path.join(os.getcwd(), path2files)
    files_list = \
        [one_file[len(files_starts):][:-len(files_ends)]
            for one_file
                in os.listdir(path2files)
                    if (one_file.lower().startswith(files_starts.lower()) and
                        one_file.lower().endswith(files_ends.lower()) and
                        one_file[0] != "~"
                    )
        ]
    
    files_list = sorted(files_list, reverse=reverse)
    return files_list


def get_report_parameters():
    myconstants.ROUND_FTE_VALUE = mytablefuncs.get_parameter_value(myconstants.ROUND_FTE_SECTION_NAME, myconstants.ROUND_FTE_DEFVALUE)
    myconstants.MEANOURSPERMONTH_VALUE = mytablefuncs.get_parameter_value(myconstants.MEANHOURSPERMONTH_SECTION_NAME, myconstants.MEANOURSPERMONTH_DEFVALUE)
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
    

def get_excel_andwork_book(excel_file_name):
    oexcel = win32com.client.Dispatch("Excel.Application")
    oexcel.Visible = oexcel.WorkBooks.Count > 0
    oexcel.DisplayAlerts = False
        
    wb = oexcel.Workbooks.Open(excel_file_name)
    currwindow = oexcel.ActiveWindow
    currwindow.WindowState = myconstants.EXCELWINDOWSTATE_MIN
    n_save_excel_calc_status = oexcel.Calculation
    oexcel.Calculation = myconstants.EXCEL_MANUAL_CALC
    
    return oexcel, currwindow, wb, n_save_excel_calc_status


def get_sheets_list(wb):
    return [one_sheet.Name for one_sheet in wb.Sheets]


def rel_path(path):
    home_dir = get_home_dir()
    ret_value = os.path.relpath(path, home_dir)

    return ret_value


def get_home_dir():
    return os.path.abspath(os.curdir)


def is_loading_error(test_data, ui_handle):
    if type(test_data) == str:
        return True

    ui_handle.set_status("------------------------------")
    ui_handle.set_status("В данных нет ни одной строки!")
    ui_handle.set_status("Сформировать отчёт невозможно!")
    ui_handle.set_status("-------------------------------")



if __name__ == "__main__":
    print(get_files_list("RawData", "", ".xlsx"))
