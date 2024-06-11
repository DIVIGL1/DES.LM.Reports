RAW_DATE_COLUMN_NAME = "Дата"
RAW_FACT_COLUMN_NAME = "Фактические трудозатраты (час.) (Сумма)"
RAW_PLAN_COLUMN_NAME = "План, FTE"

RAW_DATA_COLUMNS = {
    RAW_DATE_COLUMN_NAME: "FDate",
    "Функциональное направление": "FNRaw",
    "МВЗ": "MVZ",
    "Направление": "DivisionRaw",
    "Подразделение": "SubDivision",
    "Пользователь": "User",
    "Северный работник": "Northern",
    "Проект": "Project",
    "Статус проекта": "ProjectState",
    "Менеджер проекта": "ProjectManager",
    "Вид проекта": "ProjectType",
    "Кол-во штатных единиц": "KoBo",
    RAW_PLAN_COLUMN_NAME: "PlanFTE",
    "Договор": "Contract",
    RAW_FACT_COLUMN_NAME: "FactHours",
    # "Unnamed: 15": "Unnamed15",
    # "Unnamed: 16": "Unnamed16"
}

GROUP_COLUMNS_LIST = [
    "Дата",
    "Функциональное направление",
    "МВЗ",
    "Направление",
    "Подразделение",
    "Пользователь",
    "Северный работник",
    "Проект",
    "Статус проекта",
    "Менеджер проекта",
    "Вид проекта",
    "Кол-во штатных единиц",
    "Договор",
]

RESULT_DATA_COLUMNS = [
    "FDate",
    "Month",
    "FN",
    "Division",
    "User",
    "Project",
    "ProjectType",
    "ProjectSubType",
    "ProjectManager",
    "PlanFTE",
    "FactFTE",
    "FN_Proj",
    "FN_Proj_Month",
    "FN_Proj_User",
    "FN_Proj_User_Month",
    "Pdr_User",
    "Pdr_User_Month",
    "Pdr_User_Proj",
    "Pdr_User_Proj_Month",
    "ProjMang_Proj",
    "ProjMang_Proj_Month",
    "ProjMang_Proj_User",
    "ProjMang_Proj_User_Month",
    "ShortProject",
    "ShortProject_Month",
    "Division_Month",
    "User_Month",
    "ProjectType_Month",
    "ProjectManager_Month",
    "Pdr_User_ProjType",
    "Pdr_User_ProjType_Month",
    "FactHours",
    "ProjectTypeDescription",
    "ProjectSubTypeDescription",
    "ProjectSubType_Month",
    "Pdr_User_ProjSubType",
    "Pdr_User_ProjSubType_Month",
    "ProjectSubTypeDescription_Month",
    "pVacasia",
    "Northern",
    "JustUserName",
    "UserHourCost",
    "UserMonthCost",
    "Personal_email",
    "user_email",
    "boss_email",
    "Contract",
    "Portfolio",
    "Portfolio_Month",
    "IS_Service_type",
    "IS_Service_type_Month",
    "IS_Product_type",
    "IS_Product_type_Month",
    "Pdr_Proj",
    "Pdr_Proj_Month",
    "Proj_Pdr",
    "Proj_Pdr_Month",
    "FN_Month",
    "UHCost_KV1",
    "UMCost_KV1",
    "UHCost_KV2",
    "UMCost_KV2",
    "UHCost_KV3",
    "UMCost_KV3",
    "UHCost_KV4",
    "UMCost_KV4",
    "ISDogName",
    "SumUserFHours",
    "SumUserFactFTE",
    "SumUserFactFTEUR",
    "HourTo1FTE",
    "HourTo1FTE_Math",
    "SumInCome01",
    "SumInCome02",
    "SumInCome03",
    "SumInCome04",
    "SumInCome05",
    "SumInCome06",
    "SumInCome07",
    "SumInCome08",
    "SumInCome09",
    "SumInCome10",
    "SumInCome11",
    "SumInCome12",
    "SumPodr01",
    "SumPodr02",
    "SumPodr03",
    "SumPodr04",
    "SumPodr05",
    "SumPodr06",
    "SumPodr07",
    "SumPodr08",
    "SumPodr09",
    "SumPodr10",
    "SumPodr11",
    "SumPodr12",
    "manager_email",
    "ISys_SAP",
    "UCateg4ThisFN",
    "UCateg4ThisFN_WasFound",
    "CommonCateg_SAP",
    "CommonCateg_notSAP",
    "CategHCost",
    "CategMCost",
    "BOSS_NAME",
    "pPodryadchik",
    "SupportLevel",
    "ServiceShortName",
    "Pdr_User_ProjSTypeD",
    "Pdr_User_ProjSTypeD_Month",
]

COLUMNS_TO_SET_ZERO_IF_NULL = [
    "UserHourCost",
    "UserMonthCost",
    "UHCost_KV1",
    "UMCost_KV1",
    "UHCost_KV2",
    "UMCost_KV2",
    "UHCost_KV3",
    "UMCost_KV3",
    "UHCost_KV4",
    "UMCost_KV4",
    "SumInCome01",
    "SumInCome02",
    "SumInCome03",
    "SumInCome04",
    "SumInCome05",
    "SumInCome06",
    "SumInCome07",
    "SumInCome08",
    "SumInCome09",
    "SumInCome10",
    "SumInCome11",
    "SumInCome12",
    "SumPodr01",
    "SumPodr02",
    "SumPodr03",
    "SumPodr04",
    "SumPodr05",
    "SumPodr06",
    "SumPodr07",
    "SumPodr08",
    "SumPodr09",
    "SumPodr10",
    "SumPodr11",
    "SumPodr12",
]

MONTHS = {
    1: "Январь",
    2: "Февраль",
    3: "Март",
    4: "Апрель",
    5: "Май",
    6: "Июнь",
    7: "Июль",
    8: "Август",
    9: "Сентябрь",
    10: "Октябрь",
    11: "Ноябрь",
    12: "Декабрь",
}

MONTHS_RP = {
    1: "Января",
    2: "Февраля",
    3: "Марта",
    4: "Апреля",
    5: "Мая",
    6: "Июня",
    7: "Июля",
    8: "Августа",
    9: "Сентября",
    10: "Октября",
    11: "Ноября",
    12: "Декабря",
}

LONG_MONTHS2NUM = {
    "Январь": 1,
    "Февраль": 2,
    "Март": 3,
    "Апрель": 4,
    "Май": 5,
    "Июнь": 6,
    "Июль": 7,
    "Август": 8,
    "Сентябрь": 9,
    "Октябрь": 10,
    "Ноябрь": 11,
    "Декабрь": 12,
}

MONTHS2NUM = {
    "янв": "01",
    "фев": "02",
    "мар": "03",
    "апр": "04",
    "май": "05",
    "июн": "06",
    "июл": "07",
    "авг": "08",
    "сен": "09",
    "окт": "10",
    "ноя": "11",
    "дек": "12",
}

NUMS2MONTH = {
    "01": "янв",
    "02": "фев",
    "03": "мар",
    "04": "апр",
    "05": "май",
    "06": "июн",
    "07": "июл",
    "08": "авг",
    "09": "сен",
    "10": "окт",
    "11": "ноя",
    "12": "дек",
}

YEAR_FIELD_IN_USERS_COST = "UCYear"
YEAR_FIELD_IN_CATEGS_COST = "CCYear"
DATA_TRANSFORMATION = {
    YEAR_FIELD_IN_USERS_COST: int,
    YEAR_FIELD_IN_CATEGS_COST: int,
}

MONTHS_STR_NUMS = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]

CATEG_COLUMNS = [
    "CategKey",
    "CategName",
    "FN4CategName",
]

CATEG_COLUMN_HOUR_COST = "ChC"
CATEG_COLUMN_MONTH_COST = "CmC"

START_PARAMETERS_FILE = "Settings.xlsx"
START_PARAMETERS_FILE_4_DEVELOPER = "Settings4Developer.xlsx"

COSTS_DATA_COLUMNS = [
    "UserHourCost",
    "UserMonthCost",
    "UHCost_KV1",
    "UMCost_KV1",
    "UHCost_KV2",
    "UMCost_KV2",
    "UHCost_KV3",
    "UMCost_KV3",
    "UHCost_KV4",
    "UMCost_KV4",
]

RAW_DATA_DROP_COLUMNS = ["MVZ", "KoBo", "Unnamed15", "Unnamed16"]
SHEETS_DONT_DELETE_FORMULAS = ["ИсходныеДанные", "УникальныеСписки", "Настройки"]
FLAG_DONT_DELETE_FORMULAS_ON_THE_SHEET = "_(f)"
DELETE_SHEETS_LIST_IF_NO_FORMULAS = ["ИсходныеДанные", "УникальныеСписки", "S1Delete", "S2Delete", "S3Delete"]
DONT_REPLACE_ENTER = ["Month", "Northern", "FDate", "KoBo", "PlanFTE"]
COLUMNS_FILLNA = ["Division", "FN", "Portfolio"]
FILLNA_STRING = "???"
NOT_EXIST_ELEMENT = "< не существующий элемент >"

BOOLEAN_VALUES_SUBST = {
    "ЛОЖЬ": False,
    "Ложь": False,
    False: False,
    "ИСТИНА": True,
    "Истина": True,
    True: True,
}

PARAMETERS_SECTION_NAME = "Parameters"
RAW_DATA_SECTION_NAME = "RawDataPath"
REPORTS_SECTION_NAME = "ReportsPath"
REPORTS_PREPARED_SECTION_NAME = "ReportsPrepared"
USER_PARAMETERS_SECTION_NAME = "UserParametersPath"
ARCHIVE_SECTION_NAME = "ArchiveDir"
ALL_OTHER_PARAMETERS_SECTION_NAME = "AllOtherSections"

SECTIONS_DEFAULT_VALUES = {
    PARAMETERS_SECTION_NAME: "Parameters",
    RAW_DATA_SECTION_NAME: "RawData",
    REPORTS_SECTION_NAME: "Reports",
    REPORTS_PREPARED_SECTION_NAME: "ReportsPrepared",
    USER_PARAMETERS_SECTION_NAME: "C:/DES.LM.UserFiles",
    ARCHIVE_SECTION_NAME: "Archive",
    ALL_OTHER_PARAMETERS_SECTION_NAME: "OtherData"
}

REQUIRED_SYSTEM_SECTIONS = [
    (PARAMETERS_SECTION_NAME, "Параметры"),
    (RAW_DATA_SECTION_NAME, "Исходные (сырые) данные"),
    (REPORTS_SECTION_NAME, "Отчёты"),
    (REPORTS_PREPARED_SECTION_NAME, "Сформированные отчёты"),
    (USER_PARAMETERS_SECTION_NAME, "Таблицы с данными пользователя"),
]

ROUND_FTE_SECTION_NAME = "RoundFTE"
ROUND_FTE_DEFVALUE = 3
ROUND_FTE_VALUE = ROUND_FTE_DEFVALUE
MEANHOURSPERMONTH_SECTION_NAME = "MeanHoursPerMonth"
MEANOURSPERMONTH_DEFVALUE = 1973 / 12
MEANOURSPERMONTH_VALUE = MEANOURSPERMONTH_DEFVALUE

MONTH_WORKING_HOURS_TABLE = "WHours.xlsx"
USERS_COST_TABLE = "UCosts.xlsx"
DIVISIONS_NAMES_TABLE = "ShortDivisionNames.xlsx"
FNS_NAMES_TABLE = "ShortFNNames.xlsx"
P_FN_SUBST_TABLE = "FNSusbst.xlsx"
VIRTUAL_FTE_FILE_NAME = "Virtual FTE (YEAR).xlsx"
EMAILS_TABLE = "EMails.xlsx"
VIP_TABLE = "VIP.xlsx"
PORTFEL_TABLE = "BProg.xlsx"
ISDOGNAME_TABLE = "CrossingIS.xlsx"
PROJECTS_TYPES_DESCR = "ProjectsTypesDescriptions.xlsx"
PROJECTS_SUB_TYPES_TABLE = "ProjectsSubTypes.xlsx"
PROJECTS_SUB_TYPES_DESCR = "ProjectsSubTypesDescriptions.xlsx"
PROJECTS_LIST_ADD_INFO = "ProjectsAddInfo.xlsx"
YEARS_LIST_TABLE = "Years4Periods.xlsx"
MONTHS_LIST_TABLE = "DataPeriods.xlsx"
USERS_CATEGS_LIST = "UCategories.xlsx"
CATEGORIES_TYPES = "CategoriesTypes.xlsx"
CATEGORIES_COST_TABLE = "CCosts.xlsx"
USER_FILES_STRUCT_TABLE = "UserTablesStruct.xlsx"


PROJECTS_LIST_ADD_INFO_RAW_KEY_COLUMN = "Наименование проекта (только текст)"
PROJECTS_LIST_ADD_INFO_RENAME_COLUMNS_LIST = {
    PROJECTS_LIST_ADD_INFO_RAW_KEY_COLUMN: "Project4AddInfo",
    "Выруч_Янв": "SumInCome01",
    "Выруч_Фев": "SumInCome02",
    "Выруч_Мар": "SumInCome03",
    "Выруч_Апр": "SumInCome04",
    "Выруч_Май": "SumInCome05",
    "Выруч_Июн": "SumInCome06",
    "Выруч_Июл": "SumInCome07",
    "Выруч_Авг": "SumInCome08",
    "Выруч_Сен": "SumInCome09",
    "Выруч_Окт": "SumInCome10",
    "Выруч_Ноя": "SumInCome11",
    "Выруч_Дек": "SumInCome12",
    "Подр_Янв": "SumPodr01",
    "Подр_Фев": "SumPodr02",
    "Подр_Мар": "SumPodr03",
    "Подр_Апр": "SumPodr04",
    "Подр_Май": "SumPodr05",
    "Подр_Июн": "SumPodr06",
    "Подр_Июл": "SumPodr07",
    "Подр_Авг": "SumPodr08",
    "Подр_Сен": "SumPodr09",
    "Подр_Окт": "SumPodr10",
    "Подр_Ноя": "SumPodr11",
    "Подр_Дек": "SumPodr12",
}

PROJECTS_LIST_ADD_INFO_RAW_KEY_COLUMN = "Наименование проекта (только текст)"
GROUP_COLUMN_FOR_FILTER = "Группа"

PROJECTS_LIST_ADD_UNIQ_KEY_FIELD1 = GROUP_COLUMN_FOR_FILTER
PROJECTS_LIST_ADD_UNIQ_KEY_FIELD2 = PROJECTS_LIST_ADD_INFO_RAW_KEY_COLUMN

GROUP_COLUMN_FOR_FILTER = GROUP_COLUMN_FOR_FILTER.upper()
TEXT_4_ALL_GROUPS = "< Любые проекты >"
TEXT_4_ALL_USERS = "< Все пользователи >"
GROUP_COLUMNS_STARTER = "#"

USERS_COST_TABLE_UNIQ_KEY_FIELD1 = "CostUserName"
USERS_COST_TABLE_UNIQ_KEY_FIELD2 = YEAR_FIELD_IN_USERS_COST

CATEGS_COST_TABLE_UNIQ_KEY_FIELD1 = "CategName"
CATEGS_COST_TABLE_UNIQ_KEY_FIELD2 = "FN4CategName"
CATEGS_COST_TABLE_UNIQ_KEY_FIELD3 = YEAR_FIELD_IN_CATEGS_COST


# Словарь с описанием таблиц-параметров.
# В качестве значений используется кортежи, содержащие:
#  0. Описание
#  1. Столбец с уникальным идентификатором
#  2. Столбец с полным наименованием проекта (если он есть в таблице, если нет то указывается пустая строка)
#  3. Столбец, в котором будет проставлен уникальный код проекта (со 2-го по 5 символ)
PARAMETERS_ALL_TABLES = {
    MONTH_WORKING_HOURS_TABLE: ("Таблица с количеством рабочих часов в месяцах", "FirstDate", "", ""),
    DIVISIONS_NAMES_TABLE: ("Таблица с наименованиями подразделений", "FullDivisionName", "", ""),
    FNS_NAMES_TABLE: ("Таблица с наименованиями функциональных направлений", "FullFNName", "", ""),
    P_FN_SUBST_TABLE: ("Таблица подстановок функциональных направлений для 'проектов'", "ID.ProjectNum", "ProjectNum", "ID.ProjectNum"),
    PROJECTS_SUB_TYPES_TABLE: ("Таблица с наименованиями подтипов проектов", "ID.ProjectFromName", "ProjectName", "ID.ProjectFromName"),
    PROJECTS_TYPES_DESCR: ("Таблица с расшифровкой типов (букв) проектов", "ProjectTypeName", "", ""),
    PROJECTS_SUB_TYPES_DESCR: ("Таблица с расшифровок подтипов проектов", "ProjectSubTypeName", "", ""),
    USERS_COST_TABLE: ("Таблица часовых ставок", "UserCostKey", "", ""),
    EMAILS_TABLE: ("Таблица адресов электронной почты", "FIO_4_email", "", ""),
    VIP_TABLE: ("Таблица списка VIP", "FIO_VIP", "", ""),
    PORTFEL_TABLE: ("Таблица списка портфелей проектов", "ID_DES.LM_project", "Full_Project_name", "ID_DES.LM_project"),
    ISDOGNAME_TABLE: ("Таблица наименований ИС из контракта", "ID_DES.LM_project", "DESLM_Project", "ID_DES.LM_project"),
    PROJECTS_LIST_ADD_INFO: ("Таблица с дополнительной информацией о контрактах", "P_AddInfo_U_Key", PROJECTS_LIST_ADD_INFO_RAW_KEY_COLUMN, "ID_P_AddInfo"),
    YEARS_LIST_TABLE: ("Таблица годов, по которым можно формировать отчёт", "Years4Period", "", ""),
    MONTHS_LIST_TABLE: ("Таблица периодов (месяцев), по которым можно формировать отчёт", "DataPeriodName", "", ""),
    USERS_CATEGS_LIST: ("Таблица с перечнем категорий сотрудников", "CategUserName", "", ""),
    CATEGORIES_TYPES: ("Таблица с перечнем типов категорий сотрудников", "CategName4Type", "", ""),
    CATEGORIES_COST_TABLE: ("Таблица со списком ставок для каждой категории сотрудников", "CategKey", "", ""),
    USER_FILES_STRUCT_TABLE: ("Таблица со структурой пользовательских данных", "UserTable", "", ""),
}

LAST_INTERNET_PARAMS_NAME = "Скачанная версия справочников из Интернет"
LAST_INTERNET_PARAMS_VERSION = 0.01

LAST_INTERNET_REPORTS_NAME = "Скачанная версия ФОРМ отчётов из Интернет"
LAST_INTERNET_REPORTS_VERSION = 0.01

LAST_INTERNET_EMAILS_NAME = "Скачанная версия электронных адресов из Интернет"
LAST_INTERNET_EMAILS_VERSION = 0.01

EMAIL_INFO_COLUMNS = ["manager_email", "Personal_email", "user_email", "boss_email", "BOSS_NAME"]

PARAMETER_STR_YEAR = "YEAR"
PARAMETER_STR_MONTH1 = "MONTH1"
PARAMETER_STR_MONTH2 = "MONTH2"
PARAMETER_STR_FIRST_REPORT_DAY = "FIRST_REPORT_DAY"
PARAMETER_STR_LAST_REPORT_DAY = "LAST_REPORT_DAY"
PARAMETER_STR_KEY_WITH_PERIOD = "inputValue"

MONTHS_LIST_TABLE_PARAM_COLUMNS = ["StartMonth", "EndMonth"]

PARAMETERS_FOR_GETTING_DATA_FOR_URL ={
        "name": "period",
        "reportParameterType": "PERIOD",
        "inputValue": (
                f"{PARAMETER_STR_YEAR}-{PARAMETER_STR_MONTH1}-{PARAMETER_STR_FIRST_REPORT_DAY}" +
                "T00:00:00.000+03:00-" +
                f"{PARAMETER_STR_YEAR}-{PARAMETER_STR_MONTH2}-{PARAMETER_STR_LAST_REPORT_DAY}" +
                "T00:00:00.000+03:00"
        )
}


USER_FILES_LIST = [USERS_COST_TABLE, PROJECTS_LIST_ADD_INFO, EMAILS_TABLE]
USER_FILES_EXCLUDE_PREFFIX = "excluded__"
USER_FILES_4_COMMERCIAL_DATA_TEST = [USERS_COST_TABLE, PROJECTS_LIST_ADD_INFO]

FACT_IS_PLAN_MARKER = "(факт=плану)"
FACT_FILLER = "(факт=1)"
OTHER_PROJECT_SUB_TYPE = "_Прочее"

REPORT_FILE_PREFFIX = "Отчет - "
EXCEL_FILES_ENDS = ".xlsx".lower()

RAW_DATA_SHEET_NAME = "ИсходныеДанные"
UNIQE_LISTS_SHEET_NAME = "УникальныеСписки"
SETTINGS_SHEET_NAME = "Настройки"

PARAMETER_SAVED_MAIN_WINDOW_POZ = "Позиция и размер MainWindow"
PARAMETER_DEFAULT_MAIN_WINDOW_L = 50
PARAMETER_DEFAULT_MAIN_WINDOW_T = 50
PARAMETER_DEFAULT_MAIN_WINDOW_W = 0
PARAMETER_DEFAULT_MAIN_WINDOW_H = 0

PARAMETER_SAVED_VALUE_LEFT_AND_RIGHT_BOXES = "Ширина левого и правого блоков объектов"
PARAMETER_DEFAULT_VALUE_LEFT_AND_RIGHT_BOXES = [10, 10]

PARAMETER_SAVED_VALUE_TOP_AND_BOTTOM_BOXES = "Ширина верхнего и нижнего блоков объектов"
PARAMETER_DEFAULT_VALUE_TOP_AND_BOTTOM_BOXES = [10, 10]

PARAMETER_TIMES_TO_PRESS_F12 = 3

PARAMETER_WAITING_USER_ACTION = "Ожидание действий пользователя"

NO_CONTRACT_TYPES = ["А", "S", "В", "И", "Н"]
NO_PORTFOLIO_TYPES = ["А", "S", "В", "И", "Н"]
NO_CONTRACT_TEXT = "Не предусмотрен"
NO_PORTFOLIO_TEXT = "Не предусмотрен"
VACANCY_NAME_TEXT = "Вакансия"
PODRYADCHIK_NAME_TEXT = "Подрядчик"
FIRED_NAME_TEXT = "(Уволен) "
TEXT_LINES_SEPARATOR = "-" * 110
PARAMETER_FILENAME_OF_LAST_REPORT = "Последний сформированный отчёт"

PARAMETER_SAVED_DRAG_AND_DROP_VARIANT = "Как выполняем Drag&Drop?"

PARAMETER_SAVED_VALUE_DELETE_VIP = "Удалить VIP?"
PARAMETER_SAVED_VALUE_DELETE_CURRMONTHHALF = "Текущий месяц рассчитывать от половины нормы часов?"
PARAMETER_SAVED_VALUE_DELETE_NONPROD = "Удалять не производственные подразделения?"
PARAMETER_SAVED_VALUE_USE_ALL_P_INLIST_WITH_ADD = "Учесть все проекты из файла с доп информацией?"
PARAMETER_SAVED_VALUE_ONLY_P_WITH_ADD = "Оставлять только проекты, по которым есть доп информация?"
PARAMETER_SAVED_VALUE_SELECT_USERS = "Выбрать только людей из группы?"
PARAMETER_SAVED_VALUE_DELETE_EMPTYFACT = "Удалять строки не содержащие факта (факт=0)?"
PARAMETER_SAVED_VALUE_DELETE_PERSDATA = "Удалять проекты с персональными данными?"
PARAMETER_SAVED_VALUE_DELETE_VAC = "Удалять ли вакансии из отчёта?"
PARAMETER_SAVED_VALUE_DELETE_PODR = "Удалять ли подрядчиков из отчёта?"
PARAMETER_SAVED_VALUE_ADD_VFTE = "Добавить к списку сырых данных искусственные FTE?"
PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS = "Сохранять отчёт без формул?"
PARAMETER_SAVED_VALUE_DEL_RAWSHEET = "Удалить лист с исходными данными?"
PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL = "Открывать ли сформированный отчет в Excel?"
PARAMETER_SAVED_VALUE_COMBO_BOXES_TEXTS = "Значения выбранные в выпадающих списках"
PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_GROUPS = "Значение выбранное в выпадающем списке с группами проектов"
PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_USERS = "Значение выбранное в выпадающем списке с пользователями"
PARAMETER_SAVED_VALUE_LAST_SELECTED_YEAR = "Последний выбранный год для данных из DES.LM"
PARAMETER_SAVED_VALUE_LAST_SELECTED_MONTHS_PARAMETERS_NUM = "Номер последнего выбранного параметра с месяцами для данных из DES.LM"
PARAMETER_SAVED_VALUE_REPORT_START_MONTH = "Номер первого месяца, по которому строится отчёт"
PARAMETER_SAVED_VALUE_REPORT_END_MONTH = "Номер последнего месяца, по которому строится отчёт"


DO_IT_PREFFIX = ""

PARAMETER_SAVED_VALUE_DRAG_AND_DROP_VARIANT_DEFVALUE = 4

PARAMETER_SAVED_VALUE_DELETE_VIP_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_CURRMONTHHALF_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_NONPROD_DEFVALUE = False
PARAMETER_SAVED_VALUE_USE_ALL_P_INLIST_WITH_ADD_DEFVALUE = False
PARAMETER_SAVED_VALUE_ONLY_P_WITH_ADD_DEFVALUE = False
PARAMETER_SAVED_VALUE_SELECT_USERS_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_EMPTYFACT_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_PERSDATA_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_VAC_DEFVALUE = True
PARAMETER_SAVED_VALUE_DELETE_PODR_DEFVALUE = True
PARAMETER_SAVED_VALUE_ADD_VFTE_DEFVALUE = False
PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS_DEFVALUE = True
PARAMETER_SAVED_VALUE_DEL_RAWSHEET_DEFVALUE = False
PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL_DEFVALUE = True
PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_GROUPS_DEFVALUE = TEXT_4_ALL_GROUPS
PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_USERS_DEFVALUE = TEXT_4_ALL_USERS
PARAMETER_SAVED_VALUE_LAST_SELECTED_MONTHS_PARAMETERS_NUM_DEFVALUE = 0
PARAMETER_SAVED_VALUE_REPORT_START_MONTH_DEFVALUE = 1
PARAMETER_SAVED_VALUE_REPORT_END_MONTH_DEFVALUE = 1

PARAMETER_SAVED_SELECTED_REPORT = "Номер последнего выбранного отчёта"
PARAMETER_SAVED_SELECTED_RAW_FILE = "Имя файла последнего выбранного файла с сырыми данными"
PARAMETER_MAX_ROWS_TEST_IN_REPORT = 50000
EXCEL_MANUAL_CALC = -4135
EXCEL_AUTOMATIC_CALC = -4105
DELETE_ROW_MARKER = "delete"
HIDE_MARKER = "hide"
EXCELWINDOWSTATE_MAX = -4137  # Максимизировано
EXCELWINDOWSTATE_MIN = -4140  # Минимизировано
NUM_ROWS_FOR_HIDE = 300
NUM_COLUMNS_FOR_HIDE = 150

RES_FOLDER = "Res"

SERVICE_TYPES = "ТСД"
PROJECT_TYPES = "П"

SAP_IS_TYPE_NAME = "SAP"
NOT_SAP_IS_TYPE_NAME = "notSAP"

SERVICE_SAP_CATEGORY_BEGINNING = "SC_SAP_"
SERVICE_NOT_SAP_CATEGORY_BEGINNING = "SC_notSAP_"

PROJECT_SAP_CATEGORY_BEGINNING = "PC_SAP_"
PROJECT_NOT_SAP_CATEGORY_BEGINNING = "PC_notSAP_"

UNKNOWN_CATEGORY_BEGINNING = "XXXXX?_"

UNKNOWN_CATEGORY_NAME = " - категория не указана - "
CATEGORY_WAS_NOT_FOUND = "-"
CATEGORY_WAS_FOUND = "+"

UNKNOWN_IS_SERVICE_TYPE = "---"

REPLACE_EQ_SHEET_MARKER = "+"
MAKE_FORMULAS_MARKER = "(MBF)"
REPORT_YEAR_MARKER = "(YEAR)"
REPORT_MONTHS_PERIOD_MARKER = "(MONTHS_PERIOD)"

COMMON_VERSION = 6  # Должно быть целым
APP_VERSION = "v:12.020.270611.20"
