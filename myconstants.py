from pickle import TRUE


RAW_DATA_COLUMNS = {
                    "Дата" : "FDate",
                    "Функциональное направление" : "FNRaw",
                    "МВЗ" : "MVZ",
                    "Направление" : "DivisionRaw",
                    "Подразделение" : "SubDivision",
                    "Пользователь" : "User",
                    "Северный работник" : "Northern",
                    "Проект" : "Project",
                    "Статус проекта" : "ProjectState",
                    "Менеджер проекта" : "ProjectManager",
                    "Вид проекта" : "ProjectType",
                    "Кол-во штатных единиц" : "KoBo",
                    "План, FTE" : "PlanFTE",
                    "Договор" : "Contract",
                    "Фактические трудозатраты (час.) (Сумма)" : "FactHours",
                    "Unnamed: 15" : "Unnamed15",
                    "Unnamed: 16" : "Unnamed16"
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
                        "ProjectSubTypeDescription_Month"
                      ]

RAW_DATA_DROP_COLUMNS = ["MVZ", "KoBo", "Contract", "Unnamed15", "Unnamed16"]
SHEETS_DONT_DELETE_FORMULAS = ["ИсходныеДанные", "УникальныеСписки", "Настройки"]
DONT_REPLACE_ENTER = ["Month"]

BOOLEAN_VALUES_SUBST = {"ЛОЖЬ": 0, "ИСТИНА": 1}

PARAMETERS_SECTION_NAME = "Parameters"
RAW_DATA_SECTION_NAME = "RawDataPath"
ROUND_FTE_SECTION_NAME = "RoundFTE"
REPORTS_SECTION_NAME = "ReportsPath"
REPORTS_PREPARED_SECTION_NAME = "ReportsPrepared"

MONTH_WORKING_HOURS_TABLE = "WHours.xlsx"
DIVISIONS_NAMES_TABLE = "ShortDivisionNames.xlsx"
FNS_NAMES_TABLE = "ShortFNNames.xlsx"
VIRTUAL_FTE_FILE_NAME = "Искусственные FTE.xlsx"

PROJECTS_TYPES_DESCR = "ProjectsTypesDescriptions.xlsx"
PROJECTS_SUB_TYPES_TABLE = "ProjectsSubTypes.xlsx"
PROJECTS_SUB_TYPES_DESCR = "ProjectsSubTypesDescriptions.xlsx"

ROUND_FTE_VALUE = 5
FACT_IS_PLAN_MARKER = "(факт=плану)"
OTHER_PROJECT_SUB_TYPE = "_Прочее"

REPORT_FILE_PREFFIX = "Отчет - "
EXCEL_FILES_ENDS = ".xlsx"

RAW_DATA_SHEET_NAME = "ИсходныеДанные"
UNIQE_LISTS_SHEET_NAME = "УникальныеСписки"
SETTINGS_SHEET_NAME = "Настройки"

VACANCY_NAME_TEXT = "Вакансия"
TEXT_LINES_SEPARATOR = "-" * 110
PARAMETER_FILENAME_OF_LAST_REPORT = "Последний сформированный отчёт"

PARAMETER_SAVED_VALUE_DELETE_VAC = "Удалять ли вакансии из отчёта?"
PARAMETER_SAVED_VALUE_ADD_VFTE = "Дабавить к списку сырых данных искусственные FTE?"
PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS = "Сохранять отчёт без формул?"
PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL = "Открывать ли сформированный отчет в Excel?"

PARAMETER_SAVED_VALUE_DELETE_VAC_DEFVALUE = True
PARAMETER_SAVED_VALUE_ADD_VFTE_DEFVALUE = False
PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS_DEFVALUE = False
PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL_DEFVALUE = True

PARAMETER_SAVED_SELECTED_REPORT = "Номер последнего выбранного отчёта"
PARAMETER_MAX_ROWS_TEST_IN_REPORT = 50000
EXCEL_MANUAL_CALC = -4135
EXCEL_AUTOMATIC_CALC = -4105
DELETE_ROW_MARKER = "delete"
HIDE_MARKER = "hide"
MAXIMIZED = -4137
NUM_ROWS_FOR_HIDE = 20
NUM_COLUMNS_FOR_HIDE = 50

