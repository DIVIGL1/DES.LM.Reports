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
                    "Фактические трудозатраты (час.) (Сумма)" : "FactHour",
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
                        "ProjectManager_Month"
                      ]

RAW_DATA_DROP_COLUMNS = ["MVZ", "KoBo", "Contract", "Unnamed15", "Unnamed16"]
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
PROJECTS_SUB_TYPES_TABLE = "ProjectsSubTypes.xlsx"

ROUND_FTE_VALUE = 5
FACT_IS_PLAN_MARKER = "(факт=плану)"
OTHER_PROJECT_SUB_TYPE = "_Прочее"

REPORT_FILE_PREFFIX = "Отчет - "
EXCEL_FILES_ENDS = ".xlsx"
RAW_DATA_SHEET_NAME = "ИсходныеДанные"
UNIQE_LISTS_SHEET_NAME = "УникальныеСписки"
VACANCY_NAME_TEXT = "Вакансия"
TEXT_LINES_SEPARATOR = "-" * 110
PARAMETER_FILENAME_OF_LAST_REPORT = "Последний сформированный отчёт"
PARAMETER_SAVED_VALUE_DELETE_VAC = "Удалять ли вакансии из отчёта?"
PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL = "Открывать ли сформированный отчет в Excel?"
PARAMETER_SAVED_SELECTED_REPORT = "Номер последнего выбранного отчёта"
