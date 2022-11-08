import pandas as pd
import datetime as dt
import os

import myconstants
import myutils


def get_parameter_value(paramname, defvalue=""):
    # Читаем настройки
    if not os.path.isfile(myconstants.START_PARAMETERS_FILE):
        ret_value = defvalue
    else:
        settings_df = pd.read_excel(myconstants.START_PARAMETERS_FILE, engine='openpyxl')
        settings_df.dropna(how='all', inplace=True)
        tmp_df = settings_df[settings_df["ParameterName"] == paramname]["ParameterValue"]
        if tmp_df.shape[0] == 0:
            ret_value = defvalue
        else:
            ret_value = tmp_df.to_list()[0]
        
        if type(ret_value) == str:
            ret_value = ret_value.replace("\\", "/")
            if ret_value[-1] == "/":
                ret_value = ret_value[:-1:]
    
    return ret_value


def test_secret_files_list():
    ret_value = ""
    for one_table in myconstants.TEST_SECRET_FILES_LIST:
        full_file_path = myconstants.SECRET_COSTS_LOCATION + "/" + one_table
        if os.path.isfile(full_file_path):
            test_df = pd.read_excel(full_file_path, engine='openpyxl')
            test_df.dropna(how='all', inplace=True)
            if (test_df.select_dtypes(include='number') != 0).sum().sum() != 0:
                ret_value = f"{myconstants.TEXT_LINES_SEPARATOR}\n" + \
                            f"В таблице {one_table} в колонке присутствуют числовые значения.\n" + \
                            f"Для исключения случайного распространения конфиденциальной информации\n" + \
                            f"при пересылке файла, в отчете будет принудительно удалена закладка\n" + \
                            f"с исходными данными, а формулы в отчете будут переведены в значения.\n" + \
                            f"{myconstants.TEXT_LINES_SEPARATOR}"
                break

    return ret_value


def load_parameter_table(tablename):
    # Загружаем соответствующую таблицу с параметрами
    if (tablename == myconstants.COSTS_TABLE) and os.path.isfile(myconstants.SECRET_COSTS_LOCATION + "/" + myconstants.COSTS_TABLE):
        full_file_path = myconstants.SECRET_COSTS_LOCATION + "/" + myconstants.COSTS_TABLE
    else:
        if (tablename == myconstants.PROJECTS_LIST_ADD_INFO) and os.path.isfile(
                myconstants.SECRET_COSTS_LOCATION + "/" + myconstants.PROJECTS_LIST_ADD_INFO):
            full_file_path = myconstants.SECRET_COSTS_LOCATION + "/" + myconstants.PROJECTS_LIST_ADD_INFO
        else:
            full_file_path = get_parameter_value(myconstants.PARAMETERS_SECTION_NAME) + "/" + tablename

    parameter_df = pd.read_excel(full_file_path, engine='openpyxl')
    parameter_df.dropna(how='all', inplace=True)
    unique_key_field = myconstants.PARAMETERS_ALL_TABLES[tablename][1]
    if parameter_df.duplicated([unique_key_field]).sum() > 0:
        report_str = ""
        counter = 0
        for element in parameter_df[parameter_df.duplicated([unique_key_field])][unique_key_field].values:
            counter += 1
            report_str = report_str + f"   {counter}. {element}\n"

        return (
            f"\n\n\n" +
            f"{myconstants.TEXT_LINES_SEPARATOR}\n" +
            f"В таблице {tablename} в колонке {unique_key_field} обнаружены повторяющиеся значения:\n" +
            report_str +
            f"Сформировать отчёт невозможно, так как повторы искажают вычисления.\n"
            f"\n"
            f"Необходимо избавиться от повторов!\n"
            f"{myconstants.TEXT_LINES_SEPARATOR}"
        )
    
    return parameter_df


def load_raw_data(raw_file, ui_handle):
    # Загружаем сырые данные
    ui_handle.set_status("Начинаем загрузку и обработку исходных данных.")
    _, _, _, p_add_vfte, _, _, _ = myutils.get_report_parameters()
    if p_add_vfte:
        # Проверим наличие файла:
        virtual_fte_file = \
            os.path.join(
                os.path.join(os.getcwd(), get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)),
                myconstants.VIRTUAL_FTE_FILE_NAME)
        if not os.path.isfile(virtual_fte_file):
            ui_handle.set_status("Не обнаружен файл с искусственными FTE.")
            df = pd.read_excel(raw_file, engine='openpyxl')
        else:
            df = pd.concat(
                [pd.read_excel(raw_file, engine='openpyxl'), pd.read_excel(virtual_fte_file, engine='openpyxl')],
                sort=False, axis=0, ignore_index=True)
    else:
        df = pd.read_excel(raw_file, engine='openpyxl')

    ui_handle.set_status("Удаляем 'na'. Переименовываем столбцы и удаляем лишние.")
    df.dropna(how='all', inplace=True)
    # Необходимо проверить файл на соответствие структуре, на случай если скопировали не тот файл:
    set1 = (set(myconstants.RAW_DATA_COLUMNS.keys()) ^ {"Unnamed: 15"}) ^ {"Unnamed: 16"}
    set1 = set1 ^ {"Unnamed: 15"}
    set1 = set1 ^ {"Unnamed: 16"}
    set2 = set(df.columns)
    set2 = set2 ^ {"Unnamed: 15"}
    set2 = set2 ^ {"Unnamed: 16"}
    if not (set1 == set2):
        # Структура файла не правильная!
        return (
            f"\n\n\n" +
            f"{myconstants.TEXT_LINES_SEPARATOR}\n" +
            f"Выбранный файл имеет не правильную структуру!\n" +
            f"Сформировать отчёт невозможно!\n"
            f"{myconstants.TEXT_LINES_SEPARATOR}"
        )

    df.rename(columns=myconstants.RAW_DATA_COLUMNS, inplace=True)
    exist_drop_columns_list = list(set(myconstants.RAW_DATA_DROP_COLUMNS) & set(df.dtypes.keys()))
    df.drop(columns=exist_drop_columns_list, inplace=True)

    return df


def udata_2_date(data):

    if data != data:
        ret_date = data
    elif type(data) == str:
        ret_date = dt.datetime.strptime("01." + data, '%d.%m.%Y')
    elif type(data) == float:
        ret_date = dt.datetime.strptime("01." + str(data), '%d.%m.%Y')
    else:
        ret_date = data
    
    return ret_date


def calc_fact_fte(FactHours, Northern, CHour, NHour, Project, PlanFTE):
    if Project.find(myconstants.FACT_IS_PLAN_MARKER) >= 0:
        fact_fte = PlanFTE
    else:
        month_hours = NHour if Northern else CHour
        fact_fte = FactHours / month_hours
    return fact_fte


def add_combine_columns(df):
    df["Project7Letters"] = df["Project"].str[:7]
    
    df["FN_Proj"] = df["FN"] + "#" + df["Project7Letters"]
    df["FN_Proj_Month"] = df["FN"] + "#" + df["Project7Letters"] + "#" + df["Month"]
    
    df["FN_Proj_User"] = df["FN"] + "#" + df["Project7Letters"] + "#" + df["User"]
    df["FN_Proj_User_Month"] = df["FN"] + "#" + df["Project7Letters"] + "#" + df["User"] + "#" + df["Month"]
    
    df["Pdr_User"] = df["Division"] + "#" + df["User"]
    df["Pdr_User_Month"] = df["Division"] + "#" + df["User"] + "#" + df["Month"]
    
    df["Pdr_User_Proj"] = df["Division"] + "#" + df["User"] + "#" + df["Project7Letters"]
    df["Pdr_User_Proj_Month"] = df["Division"] + "#" + df["User"] + "#" + df["Project7Letters"] + "#" + df["Month"]
    
    df["ProjMang_Proj"] = df["ProjectManager"] + "#" + df["Project7Letters"]
    df["ProjMang_Proj_Month"] = df["ProjectManager"] + "#" + df["Project7Letters"] + "#" + df["Month"]
    
    df["ProjMang_Proj_User"] = df["ProjectManager"] + "#" + df["Project7Letters"] + "#" + df["User"]
    df["ProjMang_Proj_User_Month"] = df["ProjectManager"] + "#" + df["Project7Letters"] + "#" + df["User"] + "#" + df["Month"]

    df["ShortProject_Month"] = df["ShortProject"] + "#" + df["Month"]

    df["Division_Month"] = df["Division"] + "#" + df["Month"]
    df["User_Month"] = df["User"] + "#" + df["Month"]
    df["ProjectType_Month"] = df["ProjectType"] + "#" + df["Month"]
    df["ProjectManager_Month"] = df["ProjectManager"] + "#" + df["Month"]

    df["Pdr_User_ProjType"] = df["Division"] + "#" + df["User"] + "#" + df["ProjectType"]
    df["Pdr_User_ProjType_Month"] = df["Division"] + "#" + df["User"] + "#" + df["ProjectType"] + "#" + df["Month"]

    df["ProjectSubTypeDescription_Month"] = df["ProjectSubTypeDescription"] + "#" + df["Month"]

    df["ProjectSubType_Month"] = df["ProjectSubType"] + "#" + df["Month"]
    df["Pdr_User_ProjSubType"] = df["Division"] + "#" + df["User"] + "#" + df["ProjectSubType"]
    df["Pdr_User_ProjSubType_Month"] = df["Division"] + "#" + df["User"] + "#" + df["ProjectSubType"] + "#" + df["Month"]

    df["pVacasia"] = df["User"].apply(lambda param: True if param.replace(" ", "").lower()[:8] == myconstants.VACANCY_NAME_TEXT.lower() else False)

    df["Portfolio_Month"] = df["Portfolio"] + "#" + df["Month"]
    df["IS_Service_type_Month"] = df["IS_Service_type"] + "#" + df["Month"]
    df["IS_Product_type_Month"] = df["IS_Product_type"] + "#" + df["Month"]
    
    df["Pdr_Proj"] = df["Division"] + "#" + df["Project7Letters"]
    df["Pdr_Proj_Month"] = df["Division"] + "#" + df["Project7Letters"] + "#" + df["Month"]

    df["Proj_Pdr"] = df["Project7Letters"] + "#" + df["Division"]
    df["Proj_Pdr_Month"] = df["Project7Letters"] + "#" + df["Division"] + "#" + df["Month"]

    df["FN_Month"] = df["FN"] + "#" + df["Month"]


def prepare_data(
        raw_file_name,
        p_delete_vip,
        p_delete_not_prod_units,
        p_projects_with_add_info,
        p_delete_without_fact,
        p_curr_month_half,
        p_delete_pers_data,
        p_delete_vacation,
        ui_handle
):
    data_df = load_raw_data(raw_file_name, ui_handle)
    if type(data_df) == str:
        ui_handle.set_status(data_df)
        ui_handle.enable_buttons()
        return None
    month_hours_df = load_parameter_table(myconstants.MONTH_WORKING_HOURS_TABLE)
    if type(month_hours_df) == str:
        ui_handle.set_status(month_hours_df)
        ui_handle.enable_buttons()
        return None
    divisions_names_df = load_parameter_table(myconstants.DIVISIONS_NAMES_TABLE)
    if type(divisions_names_df) == str:
        ui_handle.set_status(divisions_names_df)
        ui_handle.enable_buttons()
        return None
    fns_names_df = load_parameter_table(myconstants.FNS_NAMES_TABLE)
    if type(fns_names_df) == str:
        ui_handle.set_status(fns_names_df)
        ui_handle.enable_buttons()
        return None
    p_fns_subst_df = load_parameter_table(myconstants.P_FN_SUBST_TABLE)
    if type(p_fns_subst_df) == str:
        ui_handle.set_status(p_fns_subst_df)
        ui_handle.enable_buttons()
        return None
    projects_sub_types_df = load_parameter_table(myconstants.PROJECTS_SUB_TYPES_TABLE)
    if type(projects_sub_types_df) == str:
        ui_handle.set_status(projects_sub_types_df)
        ui_handle.enable_buttons()
        return None
    projects_types_descr_df = load_parameter_table(myconstants.PROJECTS_TYPES_DESCR)
    if type(projects_types_descr_df) == str:
        ui_handle.set_status(projects_types_descr_df)
        ui_handle.enable_buttons()
        return None
    projects_sub_types_descr_df = load_parameter_table(myconstants.PROJECTS_SUB_TYPES_DESCR)
    if type(projects_sub_types_descr_df) == str:
        ui_handle.set_status(projects_sub_types_descr_df)
        ui_handle.enable_buttons()
        return None
    costs_df = load_parameter_table(myconstants.COSTS_TABLE)
    if type(costs_df) == str:
        ui_handle.set_status(costs_df)
        ui_handle.enable_buttons()
        return None
    emails_df = load_parameter_table(myconstants.EMAILS_TABLE)
    if type(emails_df) == str:
        ui_handle.set_status(emails_df)
        ui_handle.enable_buttons()
        return None
    vip_df = load_parameter_table(myconstants.VIP_TABLE)
    if type(vip_df) == str:
        ui_handle.set_status(vip_df)
        ui_handle.enable_buttons()
        return None
    portfolio_df = load_parameter_table(myconstants.PORTFEL_TABLE)
    if type(portfolio_df) == str:
        ui_handle.set_status(portfolio_df)
        ui_handle.enable_buttons()
        return None
    is_dog_name_df = load_parameter_table(myconstants.ISDOGNAME_TABLE)
    if type(is_dog_name_df) == str:
        ui_handle.set_status(is_dog_name_df)
        ui_handle.enable_buttons()
        return None
    projects_list_add_info = load_parameter_table(myconstants.PROJECTS_LIST_ADD_INFO)
    if type(projects_list_add_info) == str:
        ui_handle.set_status(projects_list_add_info)
        ui_handle.enable_buttons()
        return None

    projects_list_add_info.rename(columns = myconstants.PROJECTS_LIST_ADD_INFO_RENAME_COLUMNS_LIST, inplace = True)
    projects_list_add_info = projects_list_add_info[projects_list_add_info["Project4AddInfo"].notna()]
    projects_list_add_info.fillna(0.00, inplace = True)

    ui_handle.set_status(f"Загружены таблицы с параметрами (всего строк данных: {data_df.shape[0]})")
    if data_df.shape[0] == 0:
        ui_handle.set_status("")
        ui_handle.set_status("")
        ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
        ui_handle.set_status("В данных нет ни одной строки!")
        ui_handle.set_status("Сформировать отчёт невозможно!")
        ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
        ui_handle.enable_buttons()
        return None
    for column_name in set(data_df.dtypes.keys()) - set(myconstants.DONT_REPLACE_ENTER):
        if data_df.dtypes[column_name] == type(str):
            data_df[column_name] = data_df[column_name].str.replace("\n", "")
            data_df[column_name] = data_df[column_name].str.strip()
    ui_handle.set_status(f"Удалены переносы строк (всего строк данных: {data_df.shape[0]})")

    data_df["ShortProject"] = data_df["Project"].str[:5]
    projects_list_add_info["Project4AddInfo"] = projects_list_add_info["Project4AddInfo"].str[:5]

    join_type = "inner" if p_projects_with_add_info else "left"
    data_df = data_df.merge(projects_list_add_info, left_on="ShortProject", right_on="Project4AddInfo", how=join_type)

    for one_column in myconstants.PROJECTS_LIST_ADD_INFO_RENAME_COLUMNS_LIST.values():
        data_df[one_column] = data_df[one_column].fillna(0.00)

    data_df["FDate"] = data_df["FDate"].apply(lambda param: udata_2_date(param))
    ui_handle.set_status(f"Обновлён формат данных даты первого дня месяца (всего строк данных: {data_df.shape[0]})")

    data_df['Northern'].replace(myconstants.BOOLEAN_VALUES_SUBST, inplace=True)
    data_df = data_df.merge(month_hours_df, left_on="FDate", right_on="FirstDate", how="inner")
    ui_handle.set_status(f"Проведено объединение с таблицей с рабочими часами (всего строк данных: {data_df.shape[0]})")
    if data_df.shape[0] == 0:
        ui_handle.set_status("")
        ui_handle.set_status("")
        ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
        ui_handle.set_status("В данных нет ни одной строки!")
        ui_handle.set_status("Сформировать отчёт невозможно!")
        ui_handle.set_status(myconstants.TEXT_LINES_SEPARATOR)
        ui_handle.enable_buttons()
        return None
    data_df["FDate"] = data_df["FDate"].dt.strftime('%Y_%m')
    
    data_df["SumUserFHours"] = data_df.groupby(["User", "FDate"])["FactHours"].transform("sum")

    ui_handle.set_status("... начинаем пересчет фактических часов в FTE.")
    data_df["PlanFTE"] = data_df["PlanFTE"].fillna(0)
    # Получим не округлённый FTE
    data_df["FactFTEUnRounded"] = \
        data_df[["FactHours", "Northern", "CHour", "NHour", "Project", "PlanFTE"]].apply(
            lambda param: calc_fact_fte(*param), axis=1)
    # Получим округлённый FTE
    data_df["FactFTE"] = data_df["FactFTEUnRounded"].apply(lambda x: round(x, myconstants.ROUND_FTE_VALUE))

    data_df["SumUserFactFTE"] = data_df.groupby(["User", "FDate"])["FactFTE"].transform("sum")
    data_df["SumUserFactFTEUR"] = data_df.groupby(["User", "FDate"])["FactFTEUnRounded"].transform("sum")

    data_df["HourTo1FTE"] = \
        data_df[["SumUserFactFTEUR", "FactHours"]].apply(lambda x: round((x[1] / (1 if x[0] == 0 else x[0])), myconstants.ROUND_FTE_VALUE), axis=1)
    data_df["HourTo1FTE_Math"] = \
        data_df[["SumUserFactFTEUR", "FactHours"]].apply(lambda x: round(x[1] / max(x[0], 1), myconstants.ROUND_FTE_VALUE), axis=1)

    if p_curr_month_half:
        sCurrMonth = f"{dt.datetime.now().year}-{dt.datetime.now().month:0{2}}-01"
        data_df.loc[(data_df["FirstDate"] == sCurrMonth), ["FactFTE"]] = data_df[data_df["FirstDate"] == sCurrMonth]["FactFTE"] * 2

    if p_delete_without_fact:
        data_df = data_df[data_df["FactFTE"] != 0]
        ui_handle.set_status("Удалены строки без данных о факте.")
        
    ui_handle.set_status(f"Пересчитано (всего строк данных: {data_df.shape[0]})")

    data_df = data_df.merge(divisions_names_df, left_on="DivisionRaw", right_on="FullDivisionName", how="left")
    ui_handle.set_status(f"Выполнено объединение с таблицей с подразделениями (всего строк данных: {data_df.shape[0]})")
    ui_handle.set_status("... ищем пустые и восстанавливаем.")
    data_df["Division"] = data_df[["ShortDivisionName", "DivisionRaw"]].apply(lambda param: param[1] if pd.isna(param[0]) else param[0], axis=1)
    ui_handle.set_status(f"Все подразделения заполнены (всего строк данных: {data_df.shape[0]})")

    data_df = data_df.merge(p_fns_subst_df, left_on="Project", right_on="ProjectNum", how="left")
    data_df["FNRaw"] = data_df[["RealFNName", "FNRaw"]].apply(lambda param: param[1] if pd.isna(param[0]) else param[0], axis=1)
    data_df = data_df.merge(fns_names_df, left_on="FNRaw", right_on="FullFNName", how="left")
    data_df["FN"] = data_df[["ShortFNName", "FNRaw"]].apply(lambda param: param[1] if pd.isna(param[0]) else param[0], axis=1)
    ui_handle.set_status(f"Данные объединены с таблицей с ФН (всего строк данных: {data_df.shape[0]})")

    data_df["JustUserName"] = data_df["User"].apply(lambda param: param.replace(myconstants.FIRED_NAME_TEXT, ""))
    data_df = data_df.merge(costs_df, left_on="JustUserName", right_on="CostUserName", how="left")
    data_df = data_df.merge(emails_df, left_on="JustUserName", right_on="FIO_4_email", how="left")
    for one_column in myconstants.EMAIL_INFO_COLUMNS:
        data_df[one_column] = data_df[one_column].fillna("")
    
    data_df["UserHourCost"] = data_df["UserHourCost"].apply(lambda param: 0.00 if pd.isna(param) else param)
    data_df["UserMonthCost"] = data_df["UserMonthCost"].apply(lambda param: 0.00 if pd.isna(param) else param)

    data_df["UHCost_KV1"] = data_df["UHCost_KV1"].apply(lambda param: 0.00 if pd.isna(param) else param)
    data_df["UMCost_KV1"] = data_df["UMCost_KV1"].apply(lambda param: 0.00 if pd.isna(param) else param)

    data_df["UHCost_KV2"] = data_df["UHCost_KV2"].apply(lambda param: 0.00 if pd.isna(param) else param)
    data_df["UMCost_KV2"] = data_df["UMCost_KV2"].apply(lambda param: 0.00 if pd.isna(param) else param)

    data_df["UHCost_KV3"] = data_df["UHCost_KV3"].apply(lambda param: 0.00 if pd.isna(param) else param)
    data_df["UMCost_KV3"] = data_df["UMCost_KV3"].apply(lambda param: 0.00 if pd.isna(param) else param)

    data_df["UHCost_KV4"] = data_df["UHCost_KV4"].apply(lambda param: 0.00 if pd.isna(param) else param)
    data_df["UMCost_KV4"] = data_df["UMCost_KV4"].apply(lambda param: 0.00 if pd.isna(param) else param)
    
    data_df["ProjectType"] = \
        data_df[["Project", "ProjectType"]].apply(
            lambda param: "S" if param[0].find(myconstants.FACT_IS_PLAN_MARKER) >= 0 else param[1], axis=1)
    data_df = data_df.merge(projects_types_descr_df, left_on="ProjectType", right_on="ProjectTypeName", how="left")
    ui_handle.set_status(f"Уточнены типы проектов (всего строк данных: {data_df.shape[0]})")

    data_df = data_df.merge(projects_sub_types_df, left_on="Project", right_on="ProjectName", how="left")
    data_df["ProjectSubType"] = \
        data_df[["ProjectType", "ProjectSubTypePart"]].apply(
            lambda param: param[0] + myconstants.OTHER_PROJECT_SUB_TYPE if pd.isna(param[1]) else param[1], axis=1)

    data_df = data_df.merge(portfolio_df, left_on="ShortProject", right_on="ID_DES.LM_project", how="left")
    data_df["Portfolio"] = data_df["Portfolio"].fillna("")
    data_df["Contract"] = data_df["Contract"].fillna("")
    data_df["IS_Service_type"] = data_df["IS_Service_type"].fillna("")
    data_df["IS_Product_type"] = data_df["IS_Product_type"].fillna("")
    
    # Возможны пропуски в некоторых колонках. Поставим там признак, чтобы бросался в глаза в отчёте:
    for one_column in myconstants.COLUMNS_FILLNA:
        data_df[one_column] = data_df[[one_column]].fillna(myconstants.FILLNA_STRING)

    data_df = data_df.merge(is_dog_name_df, left_on="ShortProject", right_on="ID_DES.LM_project", how="left", suffixes=("", "_will_dropped"))
    data_df["ISDogName"].fillna(data_df["Project"], inplace=True)
    
    for one_type in myconstants.NO_CONTRACT_TYPES:
        data_df.loc[data_df["ProjectType"] == one_type, "Contract"] = myconstants.NO_CONTRACT_TEXT
    for one_type in myconstants.NO_PORTFOLIO_TYPES:
        data_df.loc[data_df["ProjectType"] == one_type, "Portfolio"] = myconstants.NO_PORTFOLIO_TEXT
    
    data_df = data_df.merge(projects_sub_types_descr_df, left_on="ProjectSubType", right_on="ProjectSubTypeName", how="left")
    if p_delete_pers_data:
        data_df = data_df[data_df["ProjectSubTypePersData"] != 1]

    if p_delete_vip:
        vip_list = vip_df["FIO_VIP"].to_list()
        for one_vip in vip_list:
            data_df = data_df[data_df["JustUserName"] != one_vip]

    if p_delete_not_prod_units:
        data_df = data_df[data_df["pNotProductUnit"] != 1]

    ui_handle.set_status(f"... и типы ПОДпроектов (всего строк данных: {data_df.shape[0]})")

    if p_delete_vacation:
        vacancy_text = myconstants.VACANCY_NAME_TEXT
        vacancy_text = vacancy_text.lower()
        data_df["User"] = \
            data_df["User"].apply(
                lambda param: vacancy_text if param.replace(" ", "").lower()[:len(vacancy_text)] == vacancy_text else param)
        
        data_df = data_df[data_df["User"] != vacancy_text]
        ui_handle.set_status(f"Удалены вакансии (всего строк данных: {data_df.shape[0]})")
    
    add_combine_columns(data_df)

    ui_handle.set_status(f"Добавлены производные столбцы (конкатенация) (всего строк данных: {data_df.shape[0]})")
    
    return data_df[myconstants.RESULT_DATA_COLUMNS]
