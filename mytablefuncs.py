import pandas as pd
import datetime as dt
import os

import myconstants
import myutils
from myutils import(
    iif
)


def append_categories(row, p_is_found=False):
    if row.ProjectType in myconstants.SERVICE_TYPES:
        if row.IS_Product_type_SAP == myconstants.SAP_IS_TYPE_NAME:
            service_or_project = myconstants.SERVICE_SAP_CATEGORY_BEGINNING
        else:
            service_or_project = myconstants.SERVICE_NOT_SAP_CATEGORY_BEGINNING
    elif row.ProjectType in myconstants.PROJECT_TYPES:
        if row.IS_Product_type_SAP == myconstants.SAP_IS_TYPE_NAME:
            service_or_project = myconstants.PROJECT_SAP_CATEGORY_BEGINNING
        else:
            service_or_project = myconstants.PROJECT_NOT_SAP_CATEGORY_BEGINNING
    else:
        service_or_project = myconstants.UNKNOWN_CATEGORY_BEGINNING

    if (service_or_project + row.FN) in row.index:
        if row[service_or_project + row.FN] == "":
            return myconstants.UNKNOWN_CATEGORY_NAME if not p_is_found else myconstants.CATEGORY_WAS_NOT_FOUND
        else:
            return row[service_or_project + row.FN] if not p_is_found else myconstants.CATEGORY_WAS_FOUND
    else:
        return myconstants.UNKNOWN_CATEGORY_NAME if not p_is_found else myconstants.CATEGORY_WAS_NOT_FOUND


def get_parameter_value(param_name, default_value=None):
    # Подготовим значение для возврата на случай если потребуется значение по умолчанию:
    ret_value = myconstants.SECTIONS_DEFAULT_VALUES.get(
                            param_name,
                            myconstants.SECTIONS_DEFAULT_VALUES[myconstants.ALL_OTHER_PARAMETERS_SECTION_NAME]
    )
    # Читаем настройки:
    if not os.path.isfile(myconstants.START_PARAMETERS_FILE):
        if default_value is None:
            # Возвращаемое значение вычислено в начале процедуры.
            pass
        else:
            ret_value = default_value
    else:
        settings_df = pd.read_excel(myconstants.START_PARAMETERS_FILE, engine='openpyxl')
        settings_df.dropna(how='all', inplace=True)
        tmp_df = settings_df[settings_df["ParameterName"] == param_name]["ParameterValue"]
        if tmp_df.shape[0] == 0:
            # Возвращаемое значение вычислено в начале процедуры.
            pass
        else:
            ret_value = tmp_df.to_list()[0]
        
        if type(ret_value) == str:
            ret_value = ret_value.replace("\\", "/")
            if ret_value[-1] == "/":
                ret_value = ret_value[:-1:]
    
    return ret_value


def check_for_commercial_data_in_user_files():
    ret_value = ""
    for one_table in myconstants.USER_FILES_4_COMMERCIAL_DATA_TEST:
        user_files_dir = get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME)
        full_file_path = user_files_dir + "/" + one_table
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
    # Загружаем соответствующую таблицу с параметрами.
    user_file_path = "<>!&*"
    if tablename in myconstants.USER_FILES_LIST:
        # Проверим, можем ли загрузить таблицу из пользовательских настроек:
        user_files_dir = get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME)
        user_file_path = os.path.join(os.path.join(user_files_dir, tablename))

    if os.path.isfile(user_file_path):
        full_file_path = user_file_path
    else:
        section = get_parameter_value(myconstants.PARAMETERS_SECTION_NAME)
        full_file_path = os.path.join(os.path.join(section, tablename))

    parameter_df = pd.read_excel(full_file_path, engine='openpyxl')
    parameter_df.dropna(how='all', inplace=True)
    unique_key_field = myconstants.PARAMETERS_ALL_TABLES[tablename][1]

    # Уберём строки, у которых ключевое поле - пустая строка
    parameter_df[unique_key_field].fillna("", inplace=True)
    parameter_df = parameter_df[parameter_df[unique_key_field] != ""]

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


def load_raw_data(raw_file, p_virtual_FTE, ui_handle):
    # Загружаем сырые данные
    ui_handle.add_text_to_log_box("Начинаем загрузку и обработку исходных данных.")
    df_raw = open_and_test_raw_struct(raw_file)
    if type(df_raw) == str:
        return df_raw

    if p_virtual_FTE:
        # Определим год данных в отчёте
        data_year = str(df_raw["Дата"][0]).split(".")[1]
        ui_handle.add_text_to_log_box(f"Проверяем наличие файла с искусственными FTE за {data_year} год.")
        # Проверим наличие файла:
        vff_name = myconstants.VIRTUAL_FTE_FILE_NAME.replace("YEAR", data_year)
        virtual_fte_file = \
            os.path.join(
                os.path.join(os.getcwd(), get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME)),
                vff_name)
        if not os.path.isfile(virtual_fte_file):
            ui_handle.add_text_to_log_box("")
            ui_handle.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
            ui_handle.add_text_to_log_box(f"Не обнаружен файл с искусственными FTE за {data_year} год.")
            ui_handle.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
            ui_handle.add_text_to_log_box("")
            df = df_raw
        else:
            df_virtual = open_and_test_raw_struct(virtual_fte_file, short_text=True)
            if type(df_virtual) == str:
                return (
                    "\n\n\n" +
                    f"{myconstants.TEXT_LINES_SEPARATOR}\n" +
                    f"Неудачная загрузка файла с виртуальными FTE.\n" +
                    "\n" +
                    df_virtual +
                    f"\nФайл виртуальных FTE пропущен.\n" +
                    f"{myconstants.TEXT_LINES_SEPARATOR}"
                )

            df = pd.concat([df_raw, df_virtual], sort=False, axis=0, ignore_index=True)
    else:
        df = df_raw

    ui_handle.add_text_to_log_box("Удаляем 'na'. Переименовываем столбцы и удаляем лишние.")
    df.dropna(how='all', inplace=True)
    df.rename(columns=myconstants.RAW_DATA_COLUMNS, inplace=True)
    exist_drop_columns_list = list(set(myconstants.RAW_DATA_DROP_COLUMNS) & set(df.dtypes.keys()))
    df.drop(columns=exist_drop_columns_list, inplace=True)

    return df


def open_and_test_raw_struct(xls_raw_file, short_text=False):
    # Откроем файл:
    try:
        df = pd.read_excel(xls_raw_file, engine='openpyxl')
    except FileNotFoundError:
        ret_error_value = \
            f"Не найден файл: {myutils.rel_path(xls_raw_file)}."
    except:
        ret_error_value = \
            f"Формат файла {myutils.rel_path(xls_raw_file)}\n" \
            f"не является форматом xlsx.\n" + \
            f"Такой формат файлов не поддерживается."
    else:
        # Необходимо проверить файл на соответствие структуре, на случай если в папку скопировали не тот файл:
        if not set(df.columns) <= set(myconstants.RAW_DATA_COLUMNS.keys()):
            # Структура файла не правильная!
            ret_error_value = \
                f"Выбранный файл имеет не правильную структуру!"
        else:
            # Структура правильная. Возвращаем DataFrame.
            return df

    if not short_text:
        ret_error_value = (
            "\n\n\n" +
            f"{myconstants.TEXT_LINES_SEPARATOR}\n" +
            ret_error_value +
            f"\nСформировать отчёт невозможно!\n"
            f"{myconstants.TEXT_LINES_SEPARATOR}"
        )
    return ret_error_value


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
        p_add_all_projects_with_add_info,
        p_only_projects_with_add_info,
        p_delete_without_fact,
        p_curr_month_half,
        p_delete_pers_data,
        p_delete_vacation,
        p_virtual_FTE,
        ui_handle
):
    data_df = load_raw_data(raw_file_name, p_virtual_FTE, ui_handle)
    if type(data_df) == str:
        ui_handle.add_text_to_log_box(data_df)
        return None
    month_hours_df = load_parameter_table(myconstants.MONTH_WORKING_HOURS_TABLE)
    if type(month_hours_df) == str:
        ui_handle.add_text_to_log_box(month_hours_df)
        return None
    divisions_names_df = load_parameter_table(myconstants.DIVISIONS_NAMES_TABLE)
    if type(divisions_names_df) == str:
        ui_handle.add_text_to_log_box(divisions_names_df)
        return None
    fns_names_df = load_parameter_table(myconstants.FNS_NAMES_TABLE)
    if type(fns_names_df) == str:
        ui_handle.add_text_to_log_box(fns_names_df)
        return None
    p_fns_subst_df = load_parameter_table(myconstants.P_FN_SUBST_TABLE)
    if type(p_fns_subst_df) == str:
        ui_handle.add_text_to_log_box(p_fns_subst_df)
        return None
    projects_sub_types_df = load_parameter_table(myconstants.PROJECTS_SUB_TYPES_TABLE)
    if type(projects_sub_types_df) == str:
        ui_handle.add_text_to_log_box(projects_sub_types_df)
        return None
    projects_types_descr_df = load_parameter_table(myconstants.PROJECTS_TYPES_DESCR)
    if type(projects_types_descr_df) == str:
        ui_handle.add_text_to_log_box(projects_types_descr_df)
        return None
    projects_sub_types_descr_df = load_parameter_table(myconstants.PROJECTS_SUB_TYPES_DESCR)
    if type(projects_sub_types_descr_df) == str:
        ui_handle.add_text_to_log_box(projects_sub_types_descr_df)
        return None
    costs_df = load_parameter_table(myconstants.COSTS_TABLE)
    if type(costs_df) == str:
        ui_handle.add_text_to_log_box(costs_df)
        return None
    emails_df = load_parameter_table(myconstants.EMAILS_TABLE)
    if type(emails_df) == str:
        ui_handle.add_text_to_log_box(emails_df)
        return None
    vip_df = load_parameter_table(myconstants.VIP_TABLE)
    if type(vip_df) == str:
        ui_handle.add_text_to_log_box(vip_df)
        return None
    portfolio_df = load_parameter_table(myconstants.PORTFEL_TABLE)
    if type(portfolio_df) == str:
        ui_handle.add_text_to_log_box(portfolio_df)
        return None
    is_dog_name_df = load_parameter_table(myconstants.ISDOGNAME_TABLE)
    if type(is_dog_name_df) == str:
        ui_handle.add_text_to_log_box(is_dog_name_df)
        return None
    projects_list_add_info = load_parameter_table(myconstants.PROJECTS_LIST_ADD_INFO)
    if type(projects_list_add_info) == str:
        ui_handle.add_text_to_log_box(projects_list_add_info)
        return None
    users_categs_list = load_parameter_table(myconstants.USERS_CATEGS_LIST)
    if type(users_categs_list) == str:
        ui_handle.add_text_to_log_box(users_categs_list)
        return None
    categories_types = load_parameter_table(myconstants.CATEGORIES_TYPES)
    if type(categories_types) == str:
        ui_handle.add_text_to_log_box(categories_types)
        return None
    categories_costs = load_parameter_table(myconstants.CATEGORIES_COSTS)
    if type(categories_costs) == str:
        ui_handle.add_text_to_log_box(categories_costs)
        return None

    if not p_only_projects_with_add_info:
        # Если нет отметки, что нужно выбрать только определённые проекты
        # тогда сделаем из таблицы пустую. Это нужно, чтобы исключить из неё возможные дубли.
        projects_list_add_info = projects_list_add_info[projects_list_add_info.index == -1]
    else:
        # Отмечено, что нужно выбрать только определённые проекты.
        # Наименование столбца содержащего признаки для фильтрации:
        # Найдём реальное название (с учетом регистра):
        tbl_clmns = projects_list_add_info.columns
        all_columns = [clmn.upper() for clmn in tbl_clmns]
        grp_clmn_name = tbl_clmns[all_columns.index(myconstants.GROUP_COLUMN_FOR_FILTER)]
        projects_list_add_info[grp_clmn_name] = projects_list_add_info[grp_clmn_name].fillna("").astype(str)

        # Из таблицы с дополнительными данными о проектах удалим всё лишне:
        group_value = ui_handle.comboBoxPGroups.currentText()
        projects_list_add_info = projects_list_add_info[projects_list_add_info[grp_clmn_name] == group_value]

    projects_list_add_info.rename(columns = myconstants.PROJECTS_LIST_ADD_INFO_RENAME_COLUMNS_LIST, inplace = True)
    projects_list_add_info.fillna(0.00, inplace=True)

    ui_handle.add_text_to_log_box(f"Загружены таблицы с параметрами (всего строк 'сырых' данных: {data_df.shape[0]})")
    if data_df.shape[0] == 0:
        ui_handle.add_text_to_log_box("")
        ui_handle.add_text_to_log_box("")
        ui_handle.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
        ui_handle.add_text_to_log_box("В данных нет ни одной строки!")
        ui_handle.add_text_to_log_box("Сформировать отчёт невозможно!")
        ui_handle.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
        return None

    if p_add_all_projects_with_add_info and p_delete_without_fact:
        # Эта обработка осуществляется только в случае если "чекнуты" оба checkboxes!
        # Мы должны добавить строку для каждого проекта,
        # который есть в списке - один раз для каждого месяца
        # 1. получим все уникальные месяцы:
        months_list = data_df["FDate"].unique()
        # 2. получим все уникальные проекты:
        add_projects = projects_list_add_info[myconstants.PROJECTS_LIST_ADD_INFO_RENAME_COLUMNS_LIST[myconstants.PROJECTS_LIST_ADD_INFO_RENAME_KEY_COLUMN]].values
        # 3. возьмём одну строчку из данных и заполним её "пустыми" значениями:
        one_row = data_df.loc[0]
        for idx, elm_type in enumerate(data_df.dtypes):
            if elm_type in ["float64", "int64"]:
                one_row[idx] = 0
            elif elm_type in ["object"]:
                one_row[idx] = ""
            elif elm_type in ["bool"]:
                one_row[idx] = False
            else:
                one_row[idx] = ""

            one_row["FNRaw"] = ""
            one_row["User"] = myconstants.NOT_EXIST_ELEMENT
        # 4. добавим строку по каждому проекту для каждого месяца:
        for one_month in months_list:
            for one_project in add_projects:
                one_row["FDate"] = one_month
                one_row["Project"] = one_project
                one_row["ProjectType"] = "+"

                data_df = data_df.append(one_row)

    for column_name in set(data_df.dtypes.keys()) - set(myconstants.DONT_REPLACE_ENTER):
        if data_df.dtypes[column_name] == type(str):
            data_df[column_name] = data_df[column_name].str.replace("\n", "")
            data_df[column_name] = data_df[column_name].str.strip()
    ui_handle.add_text_to_log_box(f"Удалены переносы строк (всего строк обработанных данных: {data_df.shape[0]})")

    if data_df['Northern'].isna().sum().sum():
        ui_handle.add_text_to_log_box(f"Найдены строки без признака 'Женщина с условиями работы на Крайнем Севере'.")
        ui_handle.add_text_to_log_box(f"В этих строчках принудительно подставлено значение 'Нет'.")
        ui_handle.add_text_to_log_box(f"")

        data_df['Northern'] = data_df['Northern'].fillna(False)

    data_df["ShortProject"] = data_df["Project"].str[:5]
    projects_list_add_info["Project4AddInfo"] = projects_list_add_info["Project4AddInfo"].str[:5]

    data_df["FDate"] = data_df["FDate"].apply(lambda param: udata_2_date(param))
    ui_handle.add_text_to_log_box(f"Обновлён формат данных даты первого дня месяца (всего строк обработанных данных: {data_df.shape[0]})")

    data_df['Northern'] = data_df['Northern'].apply(lambda x: myconstants.BOOLEAN_VALUES_SUBST[x])
    data_df = data_df.merge(month_hours_df, left_on="FDate", right_on="FirstDate", how="inner")
    ui_handle.add_text_to_log_box(f"Проведено объединение с таблицей с рабочими часами (всего строк обработанных данных: {data_df.shape[0]})")
    if data_df.shape[0] == 0:
        ui_handle.add_text_to_log_box("")
        ui_handle.add_text_to_log_box("")
        ui_handle.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
        ui_handle.add_text_to_log_box("В данных нет ни одной строки!")
        ui_handle.add_text_to_log_box("Сформировать отчёт невозможно!")
        ui_handle.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
        return None
    data_df["FDate"] = data_df["FDate"].dt.strftime('%Y_%m')
    
    data_df["SumUserFHours"] = data_df.groupby(["User", "FDate"])["FactHours"].transform("sum")

    ui_handle.add_text_to_log_box("... начинаем пересчет фактических часов в FTE.")
    data_df["PlanFTE"] = data_df["PlanFTE"].fillna(0)
    # Получим не округлённый FTE
    data_df["FactFTEUnRounded"] = \
        data_df[["FactHours", "Northern", "CHour", "NHour", "Project", "PlanFTE"]].apply(
            lambda param: calc_fact_fte(*param), axis=1)
    # Получим округлённый FTE
    data_df["FactFTE"] = data_df["FactFTEUnRounded"].apply(lambda x: round(x, myconstants.ROUND_FTE_VALUE))

    data_df["SumUserFactFTE"] = data_df.groupby(["User", "FDate"])["FactFTE"].transform("sum")
    data_df["SumUserFactFTEUR"] = data_df.groupby(["User", "FDate"])["FactFTEUnRounded"].transform("sum")

    # Обработаем проекты "заполнители".
    # 1. Отберём эти проекты:
    idx = ((data_df.Project.str.find(myconstants.FACT_FILLER) > 0) & (data_df["PlanFTE"] != 0))
    # И сделаем в них замену:
    #  - для не округлённого FTE:
    data_df.loc[idx, "FactFTEUnRounded"] = \
        data_df[idx][["PlanFTE", "SumUserFactFTEUR"]].apply(lambda d: max(d.PlanFTE - d.SumUserFactFTEUR, 0), axis=1)
    #  - для округлённого FTE:
    data_df.loc[idx, "FactFTE"] = \
        data_df[idx]["FactFTEUnRounded"].apply(lambda d: round(d, myconstants.ROUND_FTE_VALUE))
    #  - для часов:
    data_df.loc[idx, "FactHours"] = \
        data_df[idx][["FactFTEUnRounded", "SumUserFHours", "SumUserFactFTEUR"]].apply(
            lambda d: max(0, round((d.SumUserFHours / d.SumUserFactFTEUR) - d.SumUserFHours, 1)), axis=1
    )

    # Так как фактические FTE и часы могли поменяться из-за "заполнителей", то их необходимо пересчитать.
    data_df["SumUserFactFTE"] = data_df.groupby(["User", "FDate"])["FactFTE"].transform("sum")
    data_df["SumUserFactFTEUR"] = data_df.groupby(["User", "FDate"])["FactFTEUnRounded"].transform("sum")
    data_df["SumUserFHours"] = data_df.groupby(["User", "FDate"])["FactHours"].transform("sum")

    # Вычислим коэффициенты перевода:
    # Первый коэффициент: Часы приводятся к 1FTE в любом случае.
    # Пример: в месяце рабочих 176 часов.
    # 1) Если человек отработал на двух проектах по 100 часов, то этот коэффициент приведёт эти часы к 176.
    # 2) Если человек отработал на двух проектах по 80 часов, то этот коэффициент приведёт эти часы так же к 176.
    data_df["HourTo1FTE"] = \
        data_df[["SumUserFactFTEUR", "FactHours"]].apply(lambda x: iif(x[1]==0, x[1], x[1] / x[0]), axis=1)

    # Второй коэффициент: Если суммарно часы превышают 1FTE, то они приводятся к 1FTE, а если нет,
    # то суммарное количество часов сохраняется. То есть, не приводится к 1FTE, а сохраняется меньше.
    # Пример: в месяце рабочих 176 часов.
    # 1) Если человек отработал на двух проектах по 100 часов, то этот коэффициент приведёт эти часы к 176.
    # 2) Если человек отработал на двух проектах по 80 часов, то этот коэффициент их не изменит и сохранит часы суммарно равные 160.
    data_df["HourTo1FTE_Math"] = \
        data_df[["SumUserFactFTEUR", "FactHours"]].apply(lambda x: x[1] / max(x[0], 1), axis=1)

    ui_handle.add_text_to_log_box(f"Добавлена доп информация по проектам.")

    if p_only_projects_with_add_info:
        data_df = data_df.merge(projects_list_add_info, left_on="ShortProject", right_on="Project4AddInfo", how="inner")
        if data_df.shape[0] == 0:
            ui_handle.add_text_to_log_box(
                f"\n\n\n" +
                f"{myconstants.TEXT_LINES_SEPARATOR}\n" +
                f"В результирующей таблице нет данных.\n" +
                f"Скорее всего, это связано с фильтром по проектам.\n" +
                f"Сформировать отчёт невозможно.\n"
                f"{myconstants.TEXT_LINES_SEPARATOR}"
            )
            return None
    else:
        data_df = data_df.merge(projects_list_add_info, left_on="ShortProject", right_on="Project4AddInfo", how="left")

    if p_curr_month_half:
        sCurrMonth = f"{dt.datetime.now().year}-{dt.datetime.now().month:0{2}}-01"
        data_df.loc[(data_df["FirstDate"] == sCurrMonth), ["FactFTE"]] = data_df[data_df["FirstDate"] == sCurrMonth]["FactFTE"] * 2

    if p_delete_without_fact:
        data_df = data_df[((data_df["FactFTE"] != 0) | ((data_df["FactFTE"] == 0) & (data_df["ProjectType"] == "+")))]
        ui_handle.add_text_to_log_box("Удалены строки без данных о факте.")
        ui_handle.add_text_to_log_box(f"Пересчитано (всего строк обработанных данных: {data_df.shape[0]})")

        if data_df.shape[0] == 0:
            ui_handle.add_text_to_log_box(
                f"\n\n\n" +
                f"{myconstants.TEXT_LINES_SEPARATOR}\n" +
                f"В результирующей таблице нет данных.\n" +
                f"Скорее всего, это связано с отсутствие фактических данных.\n" +
                f"Сформировать отчёт невозможно.\n"
                f"{myconstants.TEXT_LINES_SEPARATOR}"
            )
            return None

    data_df = data_df.merge(divisions_names_df, left_on="DivisionRaw", right_on="FullDivisionName", how="left")
    ui_handle.add_text_to_log_box(f"Выполнено объединение с таблицей с подразделениями (всего строк обработанных данных: {data_df.shape[0]})")
    ui_handle.add_text_to_log_box("... ищем пустые и восстанавливаем.")
    data_df["Division"] = data_df[["ShortDivisionName", "DivisionRaw"]].apply(lambda param: param[1] if pd.isna(param[0]) else param[0], axis=1)
    ui_handle.add_text_to_log_box(f"Все подразделения заполнены (всего строк обработанных данных: {data_df.shape[0]})")

    data_df = data_df.merge(p_fns_subst_df, left_on="Project", right_on="ProjectNum", how="left")
    data_df["FNRaw"] = data_df[["RealFNName", "FNRaw"]].apply(lambda param: param[1] if pd.isna(param[0]) else param[0], axis=1)
    data_df["FNRaw"] = data_df["FNRaw"].fillna("")
    data_df = data_df.merge(fns_names_df, left_on="FNRaw", right_on="FullFNName", how="left")
    data_df["FN"] = data_df[["ShortFNName", "FNRaw"]].apply(lambda param: param[1] if pd.isna(param[0]) else param[0], axis=1)
    ui_handle.add_text_to_log_box(f"Данные объединены с таблицей с ФН (всего строк обработанных данных: {data_df.shape[0]})")

    data_df["JustUserName"] = data_df["User"].apply(lambda param: param.replace(myconstants.FIRED_NAME_TEXT, ""))
    if ui_handle.checkBoxSelectUsers.isChecked() and ui_handle.comboBoxSelectUsers.currentText() !="":
        # Отмечено, что нужно выбрать только определённые группы пользователей.
        # Наименование столбца содержащего признаки для фильтрации:
        group_field_name = myconstants.GROUP_COLUMNS_STARTER + ui_handle.comboBoxSelectUsers.currentText()
        costs_df[group_field_name] = costs_df[group_field_name].fillna("").astype(str).replace(r'\s+', '', regex=True)

        costs_df = costs_df[costs_df[group_field_name] != ""]
        costs_df = costs_df[["CostUserName"] + myconstants.COSTS_DATA_COLUMNS]
        data_df = data_df.merge(costs_df, left_on="JustUserName", right_on="CostUserName", how="inner")
        if data_df.shape[0] == 0:
            ui_handle.add_text_to_log_box(
                f"\n\n\n" +
                f"{myconstants.TEXT_LINES_SEPARATOR}\n" +
                f"В результирующей таблице нет данных.\n" +
                f"Скорее всего, это связано с фильтрами по людям.\n" +
                f"Сформировать отчёт невозможно.\n"
                f"{myconstants.TEXT_LINES_SEPARATOR}"
            )
            return None
        else:
            ui_handle.add_text_to_log_box(f"Установлен фильтр по людям (всего строк обработанных данных: {data_df.shape[0]})")
    else:
        if ui_handle.checkBoxSelectUsers.isChecked():
            ui_handle.add_text_to_log_box(
                f"{myconstants.TEXT_LINES_SEPARATOR}\n" +
                f"Не выбрана ни одна группа сотрудников.\n" +
                f"{myconstants.TEXT_LINES_SEPARATOR}"
            )

        data_df = data_df.merge(costs_df, left_on="JustUserName", right_on="CostUserName", how="left")

    projects_list_add_info.rename(columns = myconstants.PROJECTS_LIST_ADD_INFO_RENAME_COLUMNS_LIST, inplace = True)
    projects_list_add_info.fillna(0.00, inplace = True)

    data_df = data_df.merge(emails_df, left_on="JustUserName", right_on="FIO_4_email", how="left")
    emails_df = emails_df[["FIO_4_email", "user_email"]].rename(columns={"FIO_4_email": "mngr_FIO", "user_email": "manager_email"})
    data_df = data_df.merge(emails_df, left_on="ProjectManager", right_on="mngr_FIO", how="left")
    data_df[myconstants.EMAIL_INFO_COLUMNS] = data_df[myconstants.EMAIL_INFO_COLUMNS].fillna("")

    data_df[myconstants.COLUMNS_TO_SET_ZERO_IF_NULL] = data_df[myconstants.COLUMNS_TO_SET_ZERO_IF_NULL].fillna(0.00)

    data_df["ProjectType"] = \
        data_df[["Project", "ProjectType"]].apply(
            lambda param: "S" if param[0].find(myconstants.FACT_IS_PLAN_MARKER) >= 0 else param[1], axis=1)

    data_df = data_df.merge(projects_types_descr_df, left_on="ProjectType", right_on="ProjectTypeName", how="left")
    ui_handle.add_text_to_log_box(f"Уточнены типы проектов (всего строк обработанных данных: {data_df.shape[0]})")

    data_df = data_df.merge(projects_sub_types_df, left_on="Project", right_on="ProjectName", how="left")
    data_df["ProjectSubType"] = \
        data_df[["ProjectType", "ProjectSubTypePart"]].apply(
            lambda param: param[0] + myconstants.OTHER_PROJECT_SUB_TYPE if pd.isna(param[1]) else param[1], axis=1)

    data_df = data_df.merge(portfolio_df, left_on="ShortProject", right_on="ID_DES.LM_project", how="left")
    data_df["Portfolio"] = data_df["Portfolio"].fillna("")
    data_df["Contract"] = data_df["Contract"].fillna("")
    data_df["IS_Service_type"] = data_df["IS_Service_type"].fillna("")
    data_df["IS_Product_type"] = data_df["IS_Product_type"].fillna("")
    data_df["IS_Product_type_SAP"] = data_df["IS_Product_type"].apply(lambda x: "SAP" if x == "SAP" else "notSAP")

    # Возможны пропуски в некоторых колонках. Поставим там признак, чтобы бросался в глаза в отчёте:
    data_df[myconstants.COLUMNS_FILLNA] = data_df[myconstants.COLUMNS_FILLNA].fillna(myconstants.FILLNA_STRING)

    data_df = data_df.merge(is_dog_name_df, left_on="ShortProject", right_on="ID_DES.LM_project", how="left", suffixes=("", "_will_dropped"))
    data_df["ISDogName"].fillna("", inplace=True)
    
    for one_type in myconstants.NO_CONTRACT_TYPES:
        data_df.loc[data_df["ProjectType"] == one_type, "Contract"] = myconstants.NO_CONTRACT_TEXT
    for one_type in myconstants.NO_PORTFOLIO_TYPES:
        data_df.loc[data_df["ProjectType"] == one_type, "Portfolio"] = myconstants.NO_PORTFOLIO_TEXT
    
    data_df = data_df.merge(projects_sub_types_descr_df, left_on="ProjectSubType", right_on="ProjectSubTypeName", how="left")
    data_df["ProjectSubTypeDescription"].fillna("", inplace=True)
    ui_handle.add_text_to_log_box(f"... и типы ПОДпроектов (всего строк обработанных данных: {data_df.shape[0]})")

    # Добавляем категории пользователей
    users_categs_list.fillna("", inplace=True)
    data_df = data_df.merge(users_categs_list, left_on="JustUserName", right_on="CategUserName", how="left")
    data_df["UCateg4ThisFN"] = data_df.apply(lambda data: append_categories(data), axis=1)
    data_df["UCateg4ThisFN_WasFound"] = data_df.apply(lambda data: append_categories(data, True), axis=1)

    data_df["UCateg4ThisFN"].fillna(myconstants.UNKNOWN_CATEGORY_NAME, inplace=True)
    data_df["UCateg4ThisFN_WasFound"].fillna(myconstants.CATEGORY_WAS_NOT_FOUND, inplace=True)
    data_df["CommonCateg_SAP"].fillna(myconstants.UNKNOWN_CATEGORY_NAME, inplace=True)
    data_df["CommonCateg_notSAP"].fillna(myconstants.UNKNOWN_CATEGORY_NAME, inplace=True)

    data_df["CombinedUCateg4ThisFN"] = data_df[["UCateg4ThisFN", "CommonCateg_SAP", "CommonCateg_notSAP"]].apply(
        lambda data: data.CommonCateg if data.UCateg4ThisFN == "" else data.UCateg4ThisFN, axis=1
    )
    ui_handle.add_text_to_log_box(f"Добавлены категории пользователей (всего строк обработанных данных: {data_df.shape[0]})")

    if p_delete_pers_data:
        data_df = data_df[data_df["ProjectSubTypePersData"] != 1]
        ui_handle.add_text_to_log_box(f"Исключены персональные данные (всего строк обработанных данных: {data_df.shape[0]})")

    if p_delete_vip:
        vip_list = vip_df["FIO_VIP"].to_list()
        for one_vip in vip_list:
            data_df = data_df[data_df["JustUserName"] != one_vip]
        ui_handle.add_text_to_log_box(f"Исключены VIP (всего строк обработанных данных: {data_df.shape[0]})")

    if p_delete_not_prod_units:
        data_df = data_df[data_df["pNotProductUnit"] != 1]
        ui_handle.add_text_to_log_box(f"Исключены не производственные подразделения (всего строк обработанных данных: {data_df.shape[0]})")

    if p_delete_vacation:
        vacancy_text = myconstants.VACANCY_NAME_TEXT
        vacancy_text = vacancy_text.lower()
        data_df["User"] = \
            data_df["User"].apply(
                lambda param: vacancy_text if param.replace(" ", "").lower()[:len(vacancy_text)] == vacancy_text else param)
        
        data_df = data_df[data_df["User"] != vacancy_text]
        ui_handle.add_text_to_log_box(f"Удалены вакансии (всего строк обработанных данных: {data_df.shape[0]})")

    add_combine_columns(data_df)
    ui_handle.add_text_to_log_box(f"Добавлены необходимые производные столбцы (конкатенация) (всего строк данных: {data_df.shape[0]})")
    
    return data_df[myconstants.RESULT_DATA_COLUMNS]
