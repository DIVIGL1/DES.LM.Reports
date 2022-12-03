RAW_DATA_COLUMNS = {
    "Дата": "FDate",
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
    "План, FTE": "PlanFTE",
    "Договор": "Contract",
    "Фактические трудозатраты (час.) (Сумма)": "FactHours",
    "Unnamed: 15": "Unnamed15",
    "Unnamed: 16": "Unnamed16"
}

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

START_PARAMETERS_FILE = "Settings.xlsx"

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
DELETE_SHEETS_LIST_IF_NO_FORMULAS = ["ИсходныеДанные", "УникальныеСписки"]
DONT_REPLACE_ENTER = ["Month"]
COLUMNS_FILLNA = ["Division", "FN", "Portfolio"]
FILLNA_STRING = "???"

BOOLEAN_VALUES_SUBST = {"ЛОЖЬ": False, "ИСТИНА": True}

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
]

ROUND_FTE_SECTION_NAME = "RoundFTE"
ROUND_FTE_DEFVALUE = 3
ROUND_FTE_VALUE = ROUND_FTE_DEFVALUE
MEANHOURSPERMONTH_SECTION_NAME = "MeanHoursPerMonth"
MEANOURSPERMONTH_DEFVALUE = 1973 / 12
MEANOURSPERMONTH_VALUE = MEANOURSPERMONTH_DEFVALUE

MONTH_WORKING_HOURS_TABLE = "WHours.xlsx"
COSTS_TABLE = "UCosts.xlsx"
DIVISIONS_NAMES_TABLE = "ShortDivisionNames.xlsx"
FNS_NAMES_TABLE = "ShortFNNames.xlsx"
P_FN_SUBST_TABLE = "FNSusbst.xlsx"
VIRTUAL_FTE_FILE_NAME = "Virtual FTE.xlsx"
EMAILS_TABLE = "EMails.xlsx"
VIP_TABLE = "VIP.xlsx"
PORTFEL_TABLE = "BProg.xlsx"
ISDOGNAME_TABLE = "CrossingIS.xlsx"
PROJECTS_TYPES_DESCR = "ProjectsTypesDescriptions.xlsx"
PROJECTS_SUB_TYPES_TABLE = "ProjectsSubTypes.xlsx"
PROJECTS_SUB_TYPES_DESCR = "ProjectsSubTypesDescriptions.xlsx"
PROJECTS_LIST_ADD_INFO = "ProjectsAddInfo.xlsx"
PROJECTS_LIST_ADD_INFO_RENAME_COLUMNS_LIST = {
    "Наименование проекта (только текст)": "Project4AddInfo",
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
GROUP_COLUMN_FOR_FILTER = "Группа"
GROUP_COLUMN_FOR_FILTER = GROUP_COLUMN_FOR_FILTER.upper()
TEXT_4_ALL_GROUPS = "< Все группы >"
TEXT_4_ALL_USERS = "< Все пользователи >"
GROUP_COLUMNS_STARTER = "#"

PARAMETERS_ALL_TABLES = {
    MONTH_WORKING_HOURS_TABLE: ("Таблица с количеством рабочих часов в месяцах", "FirstDate"),
    DIVISIONS_NAMES_TABLE: ("Таблица с наименованиями подразделений", "FullDivisionName"),
    FNS_NAMES_TABLE: ("Таблица с наименованиями функциональных направлений", "FullFNName"),
    P_FN_SUBST_TABLE: ("Таблица подстановок названий функциональных направлений", "ProjectNum"),
    PROJECTS_SUB_TYPES_TABLE: ("Таблица с наименованиями подтипов проектов", "ProjectName"),
    PROJECTS_TYPES_DESCR: ("Таблица с расшифровкой типов (букв) проектов", "ProjectTypeName"),
    PROJECTS_SUB_TYPES_DESCR: ("Таблица с расшифровок подтипов проектов", "ProjectSubTypeName"),
    COSTS_TABLE: ("Таблица часовых ставок", "CostUserName"),
    EMAILS_TABLE: ("Таблица адресов электронной почты", "FIO_4_email"),
    VIP_TABLE: ("Таблица списка VIP", "FIO_VIP"),
    PORTFEL_TABLE: ("Таблица списка портфелей проектов", "ID_DES.LM_project"),
    ISDOGNAME_TABLE: ("Таблица наименований ИС из контракта", "ID_DES.LM_project"),
    PROJECTS_LIST_ADD_INFO: ("Таблица наименований ИС из контракта", "Наименование проекта (только текст)"),
}

EMAIL_INFO_COLUMNS = ["manager_email", "Personal_email", "user_email", "boss_email"]

USER_FILES_LIST = [COSTS_TABLE, PROJECTS_LIST_ADD_INFO, EMAILS_TABLE]
USER_FILES_EXCLUDE_PREFFIX = "excluded__"
USER_FILES_4_COMMERCIAL_DATA_TEST = [COSTS_TABLE, PROJECTS_LIST_ADD_INFO]

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
FIRED_NAME_TEXT = "(Уволен) "
TEXT_LINES_SEPARATOR = "-" * 110
PARAMETER_FILENAME_OF_LAST_REPORT = "Последний сформированный отчёт"

PARAMETER_SAVED_DRAG_AND_DROP_VARIANT = "Как выполняем Drag&Drop?"

PARAMETER_SAVED_VALUE_DELETE_VIP = "Удалить VIP?"
PARAMETER_SAVED_VALUE_DELETE_CURRMONTHHALF = "Текущий месяц рассчитывать от половины нормы часов?"
PARAMETER_SAVED_VALUE_DELETE_NONPROD = "Удалять не производственные подразделения?"
PARAMETER_SAVED_VALUE_ONLY_P_WITH_ADD = "Оставлять только проекты, по которым есть доп информация?"
PARAMETER_SAVED_VALUE_SELECT_USERS = "Выбрать только людей из группы?"
PARAMETER_SAVED_VALUE_DELETE_EMPTYFACT = "Удалять строки не содержащие факта (факт=0)?"
PARAMETER_SAVED_VALUE_DELETE_PERSDATA = "Удалять проекты с персональными данными?"
PARAMETER_SAVED_VALUE_DELETE_VAC = "Удалять ли вакансии из отчёта?"
PARAMETER_SAVED_VALUE_ADD_VFTE = "Добавить к списку сырых данных искусственные FTE?"
PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS = "Сохранять отчёт без формул?"
PARAMETER_SAVED_VALUE_DEL_RAWSHEET = "Удалить лист с исходными данными?"
PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL = "Открывать ли сформированный отчет в Excel?"
PARAMETER_SAVED_VALUE_COMBO_BOXES_TEXTS = "Значения выбранные в выпадающих списках"
PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_GROUPS = "Значение выбранное в выпадающем списке с группами проектов"
PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_USERS = "Значение выбранное в выпадающем списке с пользователями"


DO_IT_PREFFIX = ""

PARAMETER_SAVED_VALUE_DRAG_AND_DROP_VARIANT_DEFVALUE = 1

PARAMETER_SAVED_VALUE_DELETE_VIP_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_CURRMONTHHALF_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_NONPROD_DEFVALUE = False
PARAMETER_SAVED_VALUE_ONLY_P_WITH_ADD_DEFVALUE = False
PARAMETER_SAVED_VALUE_SELECT_USERS_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_EMPTYFACT_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_PERSDATA_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_VAC_DEFVALUE = True
PARAMETER_SAVED_VALUE_ADD_VFTE_DEFVALUE = False
PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS_DEFVALUE = True
PARAMETER_SAVED_VALUE_DEL_RAWSHEET_DEFVALUE = False
PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL_DEFVALUE = True
PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_GROUPS_DEFVALUE = TEXT_4_ALL_GROUPS
PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_USERS_DEFVALUE = TEXT_4_ALL_USERS

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

APP_VERSION = "v:6.03.031222.04"
