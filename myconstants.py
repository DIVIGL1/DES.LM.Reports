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
]

START_PARAMETERS_FILE = "Settings.xlsx"

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

ROUND_FTE_SECTION_NAME = "RoundFTE"
ROUND_FTE_DEFVALUE = 3
ROUND_FTE_VALUE = ROUND_FTE_DEFVALUE
MEANOURSPERMONTH_SECTION_NAME = "MeanOursPerMonth"
MEANOURSPERMONTH_DEFVALUE = 1973 / 12
MEANOURSPERMONTH_VALUE = MEANOURSPERMONTH_DEFVALUE

MONTH_WORKING_HOURS_TABLE = "WHours.xlsx"
DIVISIONS_NAMES_TABLE = "ShortDivisionNames.xlsx"
FNS_NAMES_TABLE = "ShortFNNames.xlsx"
P_FN_SUBST_TABLE = "FNSusbst.xlsx"
VIRTUAL_FTE_FILE_NAME = "Искусственные FTE.xlsx"
COSTS_TABLE = "UCosts.xlsx"
EMAILS_TABLE = "EMails.xlsx"
VIP_TABLE = "VIP.xlsx"
PORTFEL_TABLE = "BProg.xlsx"
ISDOGNAME_TABLE = "CrossingIS.xlsx"
PROJECTS_TYPES_DESCR = "ProjectsTypesDescriptions.xlsx"
PROJECTS_SUB_TYPES_TABLE = "ProjectsSubTypes.xlsx"
PROJECTS_SUB_TYPES_DESCR = "ProjectsSubTypesDescriptions.xlsx"

EMAIL_INFO_COLUMNS = ["Personal_email", "user_email", "boss_email"]
SECRET_COSTS_LOCATION = "C:/Tmp"

FACT_IS_PLAN_MARKER = "(факт=плану)"
OTHER_PROJECT_SUB_TYPE = "_Прочее"

REPORT_FILE_PREFFIX = "Отчет - "
EXCEL_FILES_ENDS = ".xlsx"

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

NO_CONTRACT_TYPES = ["А", "S", "В", "И", "Н"]
NO_PORTFOLIO_TYPES = ["А", "S", "В", "И", "Н"]
NO_CONTRACT_TEXT = "Не предусмотрен"
NO_PORTFOLIO_TEXT = "Не предусмотрен"
VACANCY_NAME_TEXT = "Вакансия"
FIRED_NAME_TEXT = "(Уволен) "
TEXT_LINES_SEPARATOR = "-" * 110
PARAMETER_FILENAME_OF_LAST_REPORT = "Последний сформированный отчёт"

PARAMETER_SAVED_VALUE_DELETE_VIP = "Удалить VIP?"
PARAMETER_SAVED_VALUE_DELETE_NONPROD = "Удалять не производственные подразделения?"
PARAMETER_SAVED_VALUE_DELETE_EMPTYFACT = "Удалять строки не содержащие факта (факт=0)?"
PARAMETER_SAVED_VALUE_DELETE_CURRMONTHHALF = "Текущий месяц расчитывать от половины нормы часов?"
PARAMETER_SAVED_VALUE_DELETE_PERSDATA = "Удалять проекты с персональными данныи?"
PARAMETER_SAVED_VALUE_DELETE_VAC = "Удалять ли вакансии из отчёта?"
PARAMETER_SAVED_VALUE_ADD_VFTE = "Дабавить к списку сырых данных искусственные FTE?"
PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS = "Сохранять отчёт без формул?"
PARAMETER_SAVED_VALUE_DEL_RAWSHEET = "Удалить лист с исходными данными?"
PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL = "Открывать ли сформированный отчет в Excel?"

DO_IT_PREFFIX = ""

PARAMETER_SAVED_VALUE_DELETE_VIP_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_NONPROD_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_EMPTYFACT_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_CURRMONTHHALF_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_PERSDATA_DEFVALUE = False
PARAMETER_SAVED_VALUE_DELETE_VAC_DEFVALUE = True
PARAMETER_SAVED_VALUE_ADD_VFTE_DEFVALUE = False
PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS_DEFVALUE = True
PARAMETER_SAVED_VALUE_DEL_RAWSHEET_DEFVALUE = False
PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL_DEFVALUE = True

PARAMETER_SAVED_SELECTED_REPORT = "Номер последнего выбранного отчёта"
PARAMETER_MAX_ROWS_TEST_IN_REPORT = 50000
EXCEL_MANUAL_CALC = -4135
EXCEL_AUTOMATIC_CALC = -4105
DELETE_ROW_MARKER = "delete"
HIDE_MARKER = "hide"
EXCELWINDOWSTATE_MAX = -4137 # Максимизировано
EXCELWINDOWSTATE_MIN = -4140 # Минимизировано
NUM_ROWS_FOR_HIDE = 300
NUM_COLUMNS_FOR_HIDE = 150

APP_VERSION = "v:2.40.041122.10"
