# pyuic5 -x Qt5Project/Windows2.ui -o myQt_form.py
# pyuic5 -x Qt5Project/_tmp_QLV.ui -o _tmp_QLV_form.py

import os
import sys

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import pyqtSignal, QObject

import myconstants
import myQt_form
from mytablefuncs import get_parameter_value, load_parameter_table
from myutils import (
    load_param, save_param, get_files_list, iif,
    open_download_dir, get_later_raw_file,
    copy_file_as_drop_process, is_admin,
    open_dir_in_explore, get_home_dir,
    open_file_in_application, test_create_dir,
    open_user_files_dir, get_resource_path
)

class animatedGifLabel(QtWidgets.QLabel):
    def __init__(self, pict_name):
        super().__init__()
        size = 10
        self.setGeometry(QtCore.QRect(0, 0, size, size))
        # self.setText("animatedGifLabel")

        self.movie = QtGui.QMovie(get_resource_path(pict_name + ".gif"))
        self.movie.setScaledSize(QtCore.QSize(15, 15))
        self.setMovie(self.movie)
        self.stop()

    def start(self):
        self.setVisible(True)
        self.movie.start()

    def stop(self):
        self.setVisible(False)
        self.movie.stop()


class Communicate(QObject):
    commander = pyqtSignal(str)


class QtMainWindow(myQt_form.Ui_MainWindow):
    closeApp = pyqtSignal()
    exit_in_process = None
    parent = None
    model = None

    def setup_reports_list(self, reports_list=None):
        if reports_list is None:
            reports_list = []
    
        self.model = QtGui.QStandardItemModel()
        self.listViewReports.setModel(self.model)

        for one_report in reports_list:
            item = QtGui.QStandardItem(one_report)
            self.model.appendRow(item)
        
        item = self.model.item(load_param(myconstants.PARAMETER_SAVED_SELECTED_REPORT, 1) - 1)

        self.listViewReports.setCurrentIndex(self.model.indexFromItem(item))

        return True

    def on_dblclick_reports_list(self):
        self.on_click_do_it()

    def on_dblclick_raw_data(self):
        self.on_click_do_it()

    def on_click_do_it(self, p_dont_clear_log_box=False):
        self.set_status_bar_text("Начато формирование отчёта...")
        self.parent.parent.reporter.create_report(p_dont_clear_log_box=p_dont_clear_log_box)

    def on_click_open_last_report(self):
        if self.pushButtonOpenLastReport.isEnabled() and self.pushButtonOpenLastReport.isVisible():
            # Открываем последний сгенерированный отчёт.
            last_report_filename = load_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT)
            if last_report_filename:
                open_file_in_application(last_report_filename)

    def open_report_form(self):
        # Открываем шаблон отчётной формы:
        report_file_name = \
            os.path.join(
                os.path.join(os.getcwd(), get_parameter_value(myconstants.REPORTS_SECTION_NAME)),
                myconstants.REPORT_FILE_PREFFIX + self.listViewReports.currentIndex().data() + myconstants.EXCEL_FILES_ENDS
            )

        open_file_in_application(report_file_name)

    def open_raw_file(self):
        # Открываем файл с 'сырыми' данными, выгруженными из DES.LM:
        raw_file_name = \
            os.path.join(
                os.path.join(os.getcwd(), get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)),
                self.listViewRawData.currentIndex().data() + myconstants.EXCEL_FILES_ENDS
            )
        open_file_in_application(raw_file_name)

    def on_click_checkboxes(self):
        if self.listViewReports.currentIndex().data() is None:
            return
        
        s_preff = self.get_preff()
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_VIP, self.checkBoxDeleteVIP.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_CURRMONTHHALF, self.checkBoxCurrMonthAHalf.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_NONPROD, self.checkBoxDeleteNotProduct.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_EMPTYFACT, self.checkBoxDeleteWithoutFact.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_PERSDATA, self.checkBoxDelPDn.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_VAC, self.checkBoxDeleteVac.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_ADD_VFTE, self.checkBoxAddVFTE.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS, self.checkBoxSaveWithOutFotmulas.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DEL_RAWSHEET, self.checkBoxDelRawData.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL, self.checkBoxOpenExcel.isChecked())
        
        self.checkBoxDelRawData.setVisible(self.checkBoxSaveWithOutFotmulas.isChecked())

    def on_click_radioButtonDD(self):
        drug_and_drop_type = (
            self.radioButtonDD1.isChecked() * 1 +
            self.radioButtonDD2.isChecked() * 2 +
            self.radioButtonDD3.isChecked() * 3 +
            self.radioButtonDD4.isChecked() * 4
        )

        save_param(myconstants.PARAMETER_SAVED_DRAG_AND_DROP_VARIANT, drug_and_drop_type)

    def on_combobox_changed(self):
        s_preff = self.get_preff()
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_GROUPS, self.comboBoxPGroups.currentText())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_USERS, self.comboBoxSelectUsers.currentText())

    def setup_radio_buttons_dd(self):
        value = load_param(
            myconstants.PARAMETER_SAVED_DRAG_AND_DROP_VARIANT,
            myconstants.PARAMETER_SAVED_VALUE_DRAG_AND_DROP_VARIANT_DEFVALUE
        )

        self.radioButtonDD1.setChecked(value == 1)
        self.radioButtonDD2.setChecked(value == 2)
        self.radioButtonDD3.setChecked(value == 3)
        self.radioButtonDD4.setChecked(value == 4)

    def on_click_reports_list(self):
        self.setup_check_boxes()
        self.setup_checkbox_only_projects_with_add()
        self.setup_checkbox_select_users()

    def get_preff(self):
        if self.listViewReports.currentIndex().data() is None:
            s_preff = ""
        else:
            s_preff = self.listViewReports.currentIndex().data() + " --> "
        return(s_preff)

    def setup_check_boxes(self):
        s_preff = self.get_preff()

        self.checkBoxDeleteVIP.setChecked(
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_VIP, myconstants.PARAMETER_SAVED_VALUE_DELETE_VIP_DEFVALUE))
        self.checkBoxDeleteNotProduct.setChecked(
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_NONPROD, myconstants.PARAMETER_SAVED_VALUE_DELETE_NONPROD_DEFVALUE))
        self.checkBoxDeleteWithoutFact.setChecked(
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_EMPTYFACT, myconstants.PARAMETER_SAVED_VALUE_DELETE_EMPTYFACT_DEFVALUE))
        self.checkBoxCurrMonthAHalf.setChecked(
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_CURRMONTHHALF, myconstants.PARAMETER_SAVED_VALUE_DELETE_CURRMONTHHALF_DEFVALUE))
        self.checkBoxDelPDn.setChecked(
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_PERSDATA, myconstants.PARAMETER_SAVED_VALUE_DELETE_PERSDATA_DEFVALUE))
        self.checkBoxDeleteVac.setChecked(
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_VAC, myconstants.PARAMETER_SAVED_VALUE_DELETE_VAC_DEFVALUE))
        self.checkBoxAddVFTE.setChecked(
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_ADD_VFTE, myconstants.PARAMETER_SAVED_VALUE_ADD_VFTE_DEFVALUE))
        self.checkBoxSaveWithOutFotmulas.setChecked(
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS, myconstants.PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS_DEFVALUE))
        self.checkBoxDelRawData.setChecked(
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DEL_RAWSHEET, myconstants.PARAMETER_SAVED_VALUE_DEL_RAWSHEET_DEFVALUE))
        self.checkBoxOpenExcel.setChecked(
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL, myconstants.PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL_DEFVALUE))
        
        self.checkBoxDelRawData.setVisible(self.checkBoxSaveWithOutFotmulas.isChecked())

    def setup_checkbox_only_projects_with_add(self):
        s_preff = self.get_preff()
        saved_value = load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_ONLY_P_WITH_ADD, myconstants.PARAMETER_SAVED_VALUE_ONLY_P_WITH_ADD_DEFVALUE)
        self.checkBoxOnlyProjectsWithAdd.setChecked(saved_value)
        self.comboBoxPGroups.setVisible(saved_value)
        if saved_value:
            self.setup_combobox_pgroups()

    def onclick_checkbox_only_projects_with_add(self):
        # Обработка клика по чекбоксу включения/выключения дополнительных параметров проектов
        s_preff = self.get_preff()
        p_selected = self.checkBoxOnlyProjectsWithAdd.isChecked()
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_ONLY_P_WITH_ADD, p_selected)
        if p_selected:
            self.setup_combobox_pgroups()
        self.comboBoxPGroups.setVisible(p_selected)

    def setup_checkbox_select_users(self):
        s_preff = self.get_preff()
        saved_value = load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_SELECT_USERS, myconstants.PARAMETER_SAVED_VALUE_SELECT_USERS_DEFVALUE)
        self.checkBoxSelectUsers.setChecked(saved_value)
        self.comboBoxSelectUsers.setVisible(saved_value)
        if saved_value:
            self.setup_combobox_select_users()

    def onclick_checkbox_select_users(self):
        # Обработка клика по чекбоксу включения/выключения выбора групп пользователей
        s_preff = self.get_preff()
        p_selected = self.checkBoxSelectUsers.isChecked()
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_SELECT_USERS, p_selected)
        self.comboBoxSelectUsers.setVisible(p_selected)
        if p_selected:
            self.setup_combobox_select_users()

    def setup_combobox_pgroups(self):
        self.comboBoxPGroups.clear()
        df = load_parameter_table(myconstants.PROJECTS_LIST_ADD_INFO)
        groups_list = [
            myconstants.TEXT_4_ALL_GROUPS,
        ]
        if type(df) == str:
            pass
        else:
            all_columns = [clmn.upper() for clmn in df.columns]
            if myconstants.GROUP_COLUMN_FOR_FILTER in all_columns:
                tbl_clmns = df.columns
                all_grp_columns = [clmn.upper() for clmn in tbl_clmns]
                selected_clmn = tbl_clmns[all_grp_columns.index(myconstants.GROUP_COLUMN_FOR_FILTER)]

                df = df[[selected_clmn]].fillna("").astype(str)
                df[selected_clmn] = df[selected_clmn].apply(lambda s: iif(s.count(" ") == len(s), "", s))

                groups = sorted(df[selected_clmn].unique())
                groups = [grn for grn in groups if grn != ""]

                groups_list = groups_list + groups

        self.comboBoxPGroups.addItems(groups_list)

        s_preff = self.get_preff()
        saved_selected_text = load_param(
            s_preff + myconstants.PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_GROUPS,
            myconstants.PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_GROUPS_DEFVALUE
        )

        index = self.comboBoxPGroups.findText(saved_selected_text, QtCore.Qt.MatchFixedString)
        if index >= 0:
            self.comboBoxPGroups.setCurrentIndex(index)

    def setup_combobox_select_users(self):
        self.comboBoxSelectUsers.clear()
        df = load_parameter_table(myconstants.COSTS_TABLE)
        if type(df) == str:
            return None
        else:
            all_grp_columns = [clmn[1:] for clmn in df.columns if clmn[0] == myconstants.GROUP_COLUMNS_STARTER]

        self.comboBoxSelectUsers.addItems(all_grp_columns)

        s_preff = self.get_preff()
        saved_selected_text = load_param(
            s_preff + myconstants.PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_USERS,
            myconstants.PARAMETER_SAVED_VALUE_COMBO_BOX_TEXT_USERS_DEFVALUE
        )

        index = self.comboBoxSelectUsers.findText(saved_selected_text, QtCore.Qt.MatchFixedString)
        if index >= 0:
            self.comboBoxSelectUsers.setCurrentIndex(index)

    def setup_form(self, reports_list):
        self.reports_list = reports_list

        self.setup_reports_list(reports_list)
        last_raw_file = load_param(myconstants.PARAMETER_SAVED_SELECTED_RAW_FILE, "")

        self.parent.refresh_raw_files_list(last_raw_file)

        self.pushButtonDoIt.clicked.connect(self.on_click_do_it)
        
        self.checkBoxDeleteVIP.clicked.connect(self.on_click_checkboxes)
        self.checkBoxDeleteNotProduct.clicked.connect(self.on_click_checkboxes)

        self.checkBoxOnlyProjectsWithAdd.clicked.connect(self.onclick_checkbox_only_projects_with_add)
        self.checkBoxSelectUsers.clicked.connect(self.onclick_checkbox_select_users)

        self.checkBoxDeleteWithoutFact.clicked.connect(self.on_click_checkboxes)
        self.checkBoxCurrMonthAHalf.clicked.connect(self.on_click_checkboxes)
        self.checkBoxDelPDn.clicked.connect(self.on_click_checkboxes)
        self.checkBoxDeleteVac.clicked.connect(self.on_click_checkboxes)
        self.checkBoxAddVFTE.clicked.connect(self.on_click_checkboxes)
        self.checkBoxOpenExcel.clicked.connect(self.on_click_checkboxes)
        self.checkBoxSaveWithOutFotmulas.clicked.connect(self.on_click_checkboxes)
        self.checkBoxDelRawData.clicked.connect(self.on_click_checkboxes)

        self.radioButtonDD1.clicked.connect(self.on_click_radioButtonDD)
        self.radioButtonDD2.clicked.connect(self.on_click_radioButtonDD)
        self.radioButtonDD3.clicked.connect(self.on_click_radioButtonDD)
        self.radioButtonDD4.clicked.connect(self.on_click_radioButtonDD)

        self.comboBoxPGroups.activated.connect(self.on_combobox_changed)
        self.comboBoxSelectUsers.activated.connect(self.on_combobox_changed)

        # Формируем обработку пунктов меню:
        self.CreateReport.triggered.connect(lambda: self.menu_action("CreateReport"))
        self.OpenLastReport.triggered.connect(lambda: self.menu_action("OpenLastReport"))
        self.OpenSavedReportsFolder.triggered.connect(lambda: self.menu_action("OpenSavedReportsFolder"))

        self.OpenDownLoads.triggered.connect(lambda: self.menu_action("OpenDownLoads"))
        self.OpenUserFilesFolder.triggered.connect(lambda: self.menu_action("OpenUserFilesFolder"))
        self.GetLastFileFromDownLoads.triggered.connect(lambda: self.menu_action("GetLastFileFromDownLoads"))
        self.MoveRawFile2Archive.triggered.connect(lambda: self.menu_action("MoveRawFile2Archive"))
        self.WaitFileAndCreateReport.triggered.connect(lambda: self.menu_action("WaitFileAndCreateReport"))

        self.EditReportForm.triggered.connect(lambda: self.menu_action("EditReportForm"))
        self.EditRawFile.triggered.connect(lambda: self.menu_action("EditRawFile"))

        self.StopWaitingFile.triggered.connect(lambda: self.menu_action("StopWaitingFile"))
        #----------------------------------
        section = myconstants.PARAMETERS_SECTION_NAME
        self.WHours.triggered.connect(lambda: self.menu_action("OpenExcel", section, "WHours"))
        self.ShortDivisionNames.triggered.connect(lambda: self.menu_action("OpenExcel", section, "ShortDivisionNames"))
        self.ShortFNNames.triggered.connect(lambda: self.menu_action("OpenExcel", section, "ShortFNNames"))
        self.FNSusbst.triggered.connect(lambda: self.menu_action("OpenExcel", section, "FNSusbst"))
        self.ProjectsSubTypes.triggered.connect(lambda: self.menu_action("OpenExcel", section, "ProjectsSubTypes"))
        self.ProjectsTypesDescriptions.triggered.connect(lambda: self.menu_action("OpenExcel", section, "ProjectsTypesDescriptions"))
        self.ProjectsSubTypesDescriptions.triggered.connect(lambda: self.menu_action("OpenExcel", section, "ProjectsSubTypesDescriptions"))
        self.BProg.triggered.connect(lambda: self.menu_action("OpenExcel", section, "BProg"))
        self.CrossingIS.triggered.connect(lambda: self.menu_action("OpenExcel", section, "CrossingIS"))
        self.VIP.triggered.connect(lambda: self.menu_action("OpenExcel", section, "VIP"))
        #----------------------------------
        self.SystemUCosts.triggered.connect(lambda: self.menu_action("OpenExcel", section, "UCosts"))
        self.SystemProjectsAddInfo.triggered.connect(lambda: self.menu_action("OpenExcel", section, "ProjectsAddInfo"))
        self.SystemEMails.triggered.connect(lambda: self.menu_action("OpenExcel", section, "EMails"))

        self.UserUCosts.triggered.connect(lambda: self.menu_action("OpenExcel", "UserParameters", "UCosts"))
        self.UserProjectsAddInfo.triggered.connect(lambda: self.menu_action("OpenExcel", "UserParameters", "ProjectsAddInfo"))
        self.UserEMails.triggered.connect(lambda: self.menu_action("OpenExcel", "UserParameters", "EMails"))

        self.UCostsSwitcher.triggered.connect(lambda: self.menu_action("ExcludeUserFile", "UserParameters", "UCosts"))
        self.ProjectsAddInfoSwitcher.triggered.connect(lambda: self.menu_action("ExcludeUserFile", "UserParameters", "ProjectsAddInfo"))
        self.EMailsSwitcher.triggered.connect(lambda: self.menu_action("ExcludeUserFile", "UserParameters", "EMails"))
        # ----------------------------------
        self.Settings.triggered.connect(lambda: self.menu_action("OpenExcel", "", "Settings"))
        # ----------------------------------
        self.Exit.triggered.connect(lambda: self.menu_action("Exit"))
        # ----------------------------------
        self.SystemUCosts.setCheckable(True)
        self.UserUCosts.setCheckable(True)
        self.SystemProjectsAddInfo.setCheckable(True)
        self.UserProjectsAddInfo.setCheckable(True)
        self.SystemEMails.setCheckable(True)
        self.UserEMails.setCheckable(True)

        self.Exit.setShortcut("Alt+F4")
        self.update_user_files_menus()

        self.edit_pads_dict = {
            "Parameters4admin": [
                self.EditReportForm,
                self.WHours,
                self.ShortDivisionNames,
                self.ShortFNNames,
                self.FNSusbst,
                self.ProjectsSubTypes,
                self.ProjectsTypesDescriptions,
                self.ProjectsSubTypesDescriptions,
                self.BProg,
                self.VIP,
                self.CrossingIS,
                self.SystemUCosts,
                self.SystemProjectsAddInfo,
                self.SystemEMails,
                self.Settings,
            ],
            "Parameters4user": [
                self.EditRawFile,
                self.UCostsSelector,
                self.UserUCosts,
                self.UCostsSwitcher,
                self.ProjectsAddInfoSelector,
                self.UserProjectsAddInfo,
                self.ProjectsAddInfoSwitcher,
                self.EMailsSelector,
                self.UserEMails,
                self.EMailsSwitcher,
            ]
        }

        self.setup_check_boxes()
        self.setup_checkbox_only_projects_with_add()
        self.setup_checkbox_select_users()
        self.setup_combobox_pgroups()
        self.setup_combobox_select_users()

        self.setup_radio_buttons_dd()

        self.listViewReports.clicked.connect(self.on_click_reports_list)
        
        self.listViewReports.doubleClicked.connect(self.on_dblclick_reports_list)
        self.listViewRawData.doubleClicked.connect(self.on_dblclick_raw_data)

        self.pushButtonOpenLastReport.clicked.connect(self.on_click_open_last_report)

        self.status_text = ""
        self.previous_status_text = ""
        self.add_text_to_log_box("> " + myconstants.PARAMETER_WAITING_USER_ACTION)
        self.statusBar.showMessage(myconstants.PARAMETER_WAITING_USER_ACTION)
        self.lock_unlock_interface_items()

    def update_user_files_menus(self):
        # Расположение пользовательских файлов:
        user_files_path = get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME)

        # Обработаем список всех возможных пользовательских файлов:
        for one_file in myconstants.USER_FILES_LIST:
            # Полный путь до рассматриваемого файла (не заблокированного):
            user_file_path = os.path.join(os.path.join(user_files_path, one_file))
            # Имя "заблокированного" файла:
            excluded_file = myconstants.USER_FILES_EXCLUDE_PREFFIX + one_file
            # Полный путь до заблокированного файла:
            user_excluded_file_path = os.path.join(os.path.join(user_files_path, excluded_file))

            # "Флаг": существует ли основной файл (НЕ заблокированный):
            user_file_exist = os.path.isfile(user_file_path)
            # "Флаг": существует ли заблокированный файл:
            user_file_locked = os.path.isfile(user_excluded_file_path)
            # "Флаг": есть ли хоть какой-то из этих файлов:
            one_of_2_files_exists = (user_file_exist or user_file_locked)

            # Если рассматриваемый файл - это ставки пользователей, то:
            if one_file == myconstants.COSTS_TABLE:
                self.UserUCosts.setEnabled(one_of_2_files_exists)
                self.UCostsSwitcher.setEnabled(one_of_2_files_exists)

                self.SystemUCosts.setChecked(not user_file_exist)
                self.UserUCosts.setChecked(user_file_exist)
                if not user_file_exist:
                    self.UCostsSwitcher.setText("Включить пользовательские настройки")
                else:
                    self.UCostsSwitcher.setText("Отключить пользовательские настройки")

            # Если рассматриваемый файл - это дополнительная информация о проектах, то:
            if one_file == myconstants.PROJECTS_LIST_ADD_INFO:
                self.UserProjectsAddInfo.setEnabled(one_of_2_files_exists)
                self.ProjectsAddInfoSwitcher.setEnabled(one_of_2_files_exists)

                self.SystemProjectsAddInfo.setChecked(not user_file_exist)
                self.UserProjectsAddInfo.setChecked(user_file_exist)
                if not user_file_exist:
                    self.ProjectsAddInfoSwitcher.setText("Включить пользовательские настройки")
                else:
                    self.ProjectsAddInfoSwitcher.setText("Отключить пользовательские настройки")

            # Если рассматриваемый файл - это списки почтовых адресов, то:
            if one_file == myconstants.EMAILS_TABLE:
                self.UserEMails.setEnabled(one_of_2_files_exists)
                self.EMailsSwitcher.setEnabled(one_of_2_files_exists)

                self.SystemEMails.setChecked(not user_file_exist)
                self.UserEMails.setChecked(user_file_exist)
                if not user_file_exist:
                    self.EMailsSwitcher.setText("Включить пользовательские настройки")
                else:
                    self.EMailsSwitcher.setText("Отключить пользовательские настройки")

    def menu_action(self, action_type, p1="", p2=""):
        if action_type == "CreateReport":
            self.set_status_bar_text("Выбрана функция формирования отчёта")
            self.on_click_do_it()
            return
        if action_type == "OpenLastReport":
            self.set_status_bar_text("Выбрана функция открытия последнего сформированного отчёта")
            self.on_click_open_last_report()
            return
        if action_type == "OpenSavedReportsFolder":
            self.set_status_bar_text("Выбрана функция директории с сохранёнными отчётами")
            section_path = os.path.join(get_home_dir(), get_parameter_value(myconstants.REPORTS_PREPARED_SECTION_NAME))
            open_dir_in_explore(section_path)
            return
        if action_type == "MoveRawFile2Archive":
            self.set_status_bar_text("Выбрана функция переноса выделенного файла в архив")
            self.move_selected_raw_file_2_archive()
            return

        if action_type == "WaitFileAndCreateReport":
            self.set_status_bar_text("... ждём новый файл в папке 'Загрузка', после чего он будет скопирован и будет запущено формирование отчёта ...", 0)
            self.clear_log_box()
            self.add_text_to_log_box("Программа переведена в режим ожидания нового файла в папке 'Загрузка'.")
            self.add_text_to_log_box("Он будет скопирован и на его основании запустится формирование отчёта.")
            self.parent.parent.waiting_file_4_report = True
            self.lock_unlock_interface_items()
            return
        if action_type == "StopWaitingFile":
            self.set_status_bar_text("Прекращено ожидание файла с данными в папке 'Загрузка'")
            self.clear_log_box()
            self.add_text_to_log_box("> " + myconstants.PARAMETER_WAITING_USER_ACTION)
            self.parent.parent.waiting_file_4_report = False
            self.lock_unlock_interface_items()
            return

        if action_type == "OpenDownLoads":
            self.set_status_bar_text("Выбрана функция открытия директории 'Загрузки'")
            open_download_dir()
            return

        if action_type == "OpenUserFilesFolder":
            self.set_status_bar_text("Выбрана функция открытия директории с пользовательскими файлами")
            open_user_files_dir()
            return

        if action_type == "GetLastFileFromDownLoads":
            self.set_status_bar_text("Выбрана функция копирования последнего файла Excel из директории 'Загрузки'")
            raw_file = get_later_raw_file()
            if raw_file is None:
                return

            copy_file_as_drop_process(self.parent, [raw_file])
            return

        if action_type == "EditReportForm":
            self.set_status_bar_text("Выбрана функция редактирования выделенного шаблона отчёта")
            self.open_report_form()
            return
        if action_type == "EditRawFile":
            self.set_status_bar_text("Выбрана функция редактирования выделенного файла с данными")
            self.open_raw_file()
            return

        if action_type == "ExcludeUserFile":
            user_files_dir = get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME)
            user_file_path = os.path.join(os.path.join(user_files_dir, p2 + myconstants.EXCEL_FILES_ENDS))
            excluded_file = myconstants.USER_FILES_EXCLUDE_PREFFIX + p2 + myconstants.EXCEL_FILES_ENDS
            user_excluded_file_path = os.path.join(os.path.join(user_files_dir, excluded_file))

            user_file_exist = os.path.isfile(user_file_path)
            user_file_locked = os.path.isfile(user_excluded_file_path)

            if user_file_exist:
                self.set_status_bar_text("Выбрана функция, исключающая пользовательский файл с данными из обработки")
                # У нас существует основной файл и его надо переименовать в "заблокированный"
                # Наличие в это же время "заблокированного" файла нас не интересует.
                try:
                    os.rename(user_file_path, user_excluded_file_path)
                except:
                    self.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
                    self.add_text_to_log_box("Не удалось отключить пользовательский файл.")
                    self.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
            elif user_file_locked:
                self.set_status_bar_text("Выбрана функция, позволяющая использовать пользовательский файл в обработке")
                # У нас НЕ существует основного файла, но существует "заблокированный".
                # Переименуем его.
                try:
                    os.rename(user_excluded_file_path, user_file_path)
                except:
                    self.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
                    self.add_text_to_log_box("Не удалось подключить пользовательский файл.")
                    self.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)

            return

        if action_type == "OpenExcel":
            self.set_status_bar_text("Выбрана функция редактирования одного из файла настроек")
            if p1 == "":
                section = ""
            elif p1 == "UserParameters":
                # Проверим существование пользовательских файлов:
                user_files_dir = get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME)
                user_file_path = os.path.join(os.path.join(user_files_dir, p2 + myconstants.EXCEL_FILES_ENDS))
                excluded_file = myconstants.USER_FILES_EXCLUDE_PREFFIX + p2 + myconstants.EXCEL_FILES_ENDS
                user_excluded_file_path = os.path.join(os.path.join(user_files_dir, excluded_file))

                user_file_exist = os.path.isfile(user_file_path)
                user_file_locked = os.path.isfile(user_excluded_file_path)
                one_of_2_files_exists = (user_file_exist or user_file_locked)
                if not one_of_2_files_exists:
                    # Если оба файла не доступны, то и редактировать нечего.
                    return
                if not user_file_exist:
                    # Уточняем имя файла в параметре p1:
                    p2 = myconstants.USER_FILES_EXCLUDE_PREFFIX + p2

                section = get_parameter_value(myconstants.USER_PARAMETERS_SECTION_NAME)
            else:
                section = get_parameter_value(p1)

            xls_file_path = os.path.join(os.path.join(os.getcwd(), section, p2 + myconstants.EXCEL_FILES_ENDS))
            open_file_in_application(xls_file_path)
            return

        if action_type == "Exit":
            self.set_status_bar_text("Выбрано прекращение работы программы")
            self.parent.close()

        print(action_type)

    def move_selected_raw_file_2_archive(self):
        raw_file_name = self.listViewRawData.currentIndex().data()
        self.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
        self.move_one_raw_file_2_archive(raw_file_name)
        self.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)

    def move_one_raw_file_2_archive(self, raw_file_name):
        # Определим пути
        section_path = get_parameter_value(myconstants.RAW_DATA_SECTION_NAME)
        archive_dir = get_parameter_value(myconstants.ARCHIVE_SECTION_NAME)
        archive_path = os.path.join(section_path, archive_dir)
        if not test_create_dir(archive_path):
            return

        raw_file_name = raw_file_name + myconstants.EXCEL_FILES_ENDS
        src_archive_file_path = os.path.join(section_path, raw_file_name)
        dst_archive_file_path = os.path.join(archive_path, raw_file_name)

        try:
            os.replace(src_archive_file_path, dst_archive_file_path)
            self.add_text_to_log_box(f"   Файл {raw_file_name} перемещён в архив.")
        except:
            self.add_text_to_log_box(f"   Перемещение файла {raw_file_name} в архив не удалось - возникли ошибки.")

    def lock_unlock_interface_items(self):
        processing_report = self.parent.parent.reporter.report_creation_in_process
        processing_drag_and_drop = self.parent.parent.drag_and_prop_in_process
        waiting_file_4_report = self.parent.parent.waiting_file_4_report
        report_automation_in_process = self.parent.parent.report_automation_in_process
        last_report_exists = load_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT, "") != ""

        is_user_admin = is_admin()

        if waiting_file_4_report or report_automation_in_process:
            self.parent.setAcceptDrops(False)
            # ----------------------------------------------------------
            # Ждём файл. ...или запущен автоматический
            # процесс: копирование файла и формирование отчёта
            # ----------------------------------------------------------
            # Ситуация когда не надо ничего делать - просто ждём.

            if processing_report:
                self.parent.ag_switcher("report_preparation_ag")
            else:
                self.parent.ag_switcher("wait_file_ag")

            # В этом случае запрещено:
            self.pushButtonDoIt.setEnabled(False)
            self.CreateReport.setEnabled(False)

            self.GetLastFileFromDownLoads.setEnabled(False)
            self.MoveRawFile2Archive.setEnabled(False)

            self.OpenLastReport.setEnabled(False)

            self.WaitFileAndCreateReport.setEnabled(False)

            for one_pad in self.edit_pads_dict["Parameters4admin"]:
                one_pad.setEnabled(False)

            for one_pad in self.edit_pads_dict["Parameters4user"]:
                one_pad.setEnabled(False)

            # В этом случае разрешено:
            self.OpenSavedReportsFolder.setEnabled(True)
            self.OpenDownLoads.setEnabled(True)

            self.Exit.setEnabled(True)

            # А вот это разрешено только если НЕ запущен автомат:
            self.StopWaitingFile.setEnabled(True and not report_automation_in_process)

            # Кнопка не доступна, но видна если есть последний отчёт
            self.pushButtonOpenLastReport.setEnabled(False)
            self.pushButtonOpenLastReport.setVisible(last_report_exists)

        else:
            if not processing_report and processing_drag_and_drop:
                self.parent.setAcceptDrops(False)
                # ----------------------------------------------------------
                # Drag&Drop!
                # ----------------------------------------------------------
                # Ситуация когда не надо запускать отчёты и выполнять
                # другие функции с выводом на экран. Но можно открывать папки,
                # редактировать файлы, перемещать файлы в архив.

                self.parent.ag_switcher("drag_and_drop_ag")

                # В этом случае запрещено:
                self.pushButtonDoIt.setEnabled(False)

                self.CreateReport.setEnabled(False)
                self.Exit.setEnabled(False)
                self.GetLastFileFromDownLoads.setEnabled(False)

                self.Automation.setEnabled(False)

                # В этом случае разрешено:
                self.OpenLastReport.setEnabled(True)
                self.OpenSavedReportsFolder.setEnabled(True)
                self.OpenDownLoads.setEnabled(True)

                for one_pad in self.edit_pads_dict["Parameters4admin"]:
                    one_pad.setEnabled(is_user_admin)

                for one_pad in self.edit_pads_dict["Parameters4user"]:
                    one_pad.setEnabled(True)

                # Кнопка доступна только если есть отчёт
                self.pushButtonOpenLastReport.setEnabled(last_report_exists)
                self.pushButtonOpenLastReport.setVisible(last_report_exists)

            elif processing_report and not processing_drag_and_drop:
                self.parent.setAcceptDrops(False)
                # ----------------------------------------------------------
                # Формируется отчёт!
                # ----------------------------------------------------------
                # Ситуация когда не надо выполнять Drag&Drop,
                # не надо формировать другие отчёты,
                # Не надо ничего редактировать через Excel
                # и не надо перемещать файлы в архив.
                # Но можно открывать папки.

                self.parent.ag_switcher("report_preparation_ag")

                # В этом случае запрещено:
                self.pushButtonDoIt.setEnabled(False)
                self.pushButtonOpenLastReport.setEnabled(False)

                self.CreateReport.setEnabled(False)
                self.OpenLastReport.setEnabled(False)
                self.GetLastFileFromDownLoads.setEnabled(False)

                self.Automation.setEnabled(False)

                for one_pad in self.edit_pads_dict["Parameters4admin"]:
                    one_pad.setEnabled(False)

                for one_pad in self.edit_pads_dict["Parameters4user"]:
                    one_pad.setEnabled(False)

                # В этом случае разрешено:
                self.Exit.setEnabled(True)
                self.OpenDownLoads.setEnabled(True)
                self.OpenSavedReportsFolder.setEnabled(True)

                # Кнопка не доступна, но видна если есть последний отчёт
                self.pushButtonOpenLastReport.setEnabled(False)
                self.pushButtonOpenLastReport.setVisible(last_report_exists)

            elif not processing_report and not processing_drag_and_drop:
                self.parent.setAcceptDrops(True)
                # ----------------------------------------------------------
                # НЕ формируется отчёт и НЕ выполняется Drag&Drop...
                # ----------------------------------------------------------

                self.parent.ag_switcher("waiting_user_action_ag")

                # В этом случае разрешено всё:
                self.pushButtonDoIt.setEnabled(True)

                self.CreateReport.setEnabled(True)
                self.OpenLastReport.setEnabled(self.pushButtonOpenLastReport.isEnabled() and self.pushButtonOpenLastReport.isVisible())
                self.Exit.setEnabled(True)
                self.OpenDownLoads.setEnabled(True)
                self.OpenSavedReportsFolder.setEnabled(True)
                self.WaitFileAndCreateReport.setEnabled(True)
                self.GetLastFileFromDownLoads.setEnabled(True)

                self.Automation.setEnabled(True)
                self.MoveRawFile2Archive.setEnabled(True)
                self.StopWaitingFile.setEnabled(False)

                for one_pad in self.edit_pads_dict["Parameters4admin"]:
                    one_pad.setEnabled(is_user_admin)

                for one_pad in self.edit_pads_dict["Parameters4user"]:
                    one_pad.setEnabled(True)

                # Кнопка доступна и видна только если есть отчёт
                self.pushButtonOpenLastReport.setEnabled(last_report_exists)
                self.pushButtonOpenLastReport.setVisible(last_report_exists)

    def set_status_bar_text(self, text, sec=5):
        self.statusBar.showMessage(text, sec * 1000)

    def update_log_box_text(self):
        self.plainTextEdit.setPlainText(self.status_text)
    
    def add_text_to_log_box(self, status_text):
        start_text_value = self.status_text
        self.previous_status_text = self.status_text
        self.status_text = start_text_value + ("\n" if start_text_value != "" else "") + status_text
        self.parent.communicate.commander.emit("update_log_box_text")

    def change_last_log_box_text(self, status_text):
        start_text_value = self.previous_status_text
        self.status_text = start_text_value + ("\n" if start_text_value != "" else "") + status_text
        self.parent.communicate.commander.emit("update_log_box_text")

    def clear_log_box(self):
        self.status_text = ""
        self.previous_status_text = ""
        self.parent.communicate.commander.emit("update_log_box_text")
        
    def resize_text_and_button(self):
        last_report_filename = load_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT)
        if last_report_filename == "":
            self.pushButtonOpenLastReport.setVisible(False)
        else:
            self.pushButtonOpenLastReport.setVisible(True)
        
    def save_app_link(self, app):
        self.app = app


class MyWindow(QtWidgets.QMainWindow):
    ui = None
    f12_counter = 0
    start_flag = True
    ready_to_save_position = False
    ctrl_is_pressed = False
    alt_is_pressed = False

    def __init__(self, parent):
        self.parent = parent
        self.app = QtWidgets.QApplication(sys.argv)
        QtWidgets.QMainWindow.__init__(self, None)
        self.ui = QtMainWindow()
        self.ui.setupUi(self)
        self.ui.parent = self
        self.ui.save_app_link(self.app)
        self.setWindowTitle(f"DES.LM.Reporter ({myconstants.APP_VERSION})")
        self.ui.plainTextEdit.setWordWrapMode(QtGui.QTextOption.NoWrap)
        # Установим исходные (сохранённые) координаты и размеры:
        data = load_param(myconstants.PARAMETER_SAVED_MAIN_WINDOW_POZ, "")
        left_and_right_boxes_widths = load_param(myconstants.PARAMETER_SAVED_VALUE_LEFT_AND_RIGHT_BOXES,
                                                 myconstants.PARAMETER_DEFAULT_VALUE_LEFT_AND_RIGHT_BOXES)
        top_and_bottom_boxes_widths = load_param(myconstants.PARAMETER_SAVED_VALUE_TOP_AND_BOTTOM_BOXES,
                                                 myconstants.PARAMETER_DEFAULT_VALUE_TOP_AND_BOTTOM_BOXES)

        self.ui.HorisontalSplitter.setSizes(top_and_bottom_boxes_widths)

        if data:
            self.restoreGeometry(data)
            self.ui.VerticalSplitter.setSizes(left_and_right_boxes_widths)

        self.ui.VerticalSplitter.splitterMoved.connect(self.save_coordinates)
        self.ui.HorisontalSplitter.splitterMoved.connect(self.save_coordinates)

        self.wait_file_ag = animatedGifLabel("search")
        self.ui.statusBar.addPermanentWidget(self.wait_file_ag)

        self.report_preparation_ag = animatedGifLabel("book")
        self.ui.statusBar.addPermanentWidget(self.report_preparation_ag)

        self.drag_and_drop_ag = animatedGifLabel("download")
        self.ui.statusBar.addPermanentWidget(self.drag_and_drop_ag)

        self.waiting_user_action_ag = animatedGifLabel("coffee")
        self.ui.statusBar.addPermanentWidget(self.waiting_user_action_ag)

        self.communicate = Communicate()
        self.communicate.commander.connect(lambda command: self.communication_handler(command))


    def ag_switcher(self, ag_name):
        ag_list = [
            "wait_file_ag",
            "drag_and_drop_ag",
            "report_preparation_ag",
            "waiting_user_action_ag",
        ]

        for element in ag_list:
            if element == ag_name:
                self.communicate.commander.emit(element + " start")
            else:
                self.communicate.commander.emit(element + " stop")

    def communication_handler(self, command):
        if command == "wait_file_ag start":
            self.wait_file_ag.start()
        elif command == "wait_file_ag stop":
            self.wait_file_ag.stop()

        elif command == "drag_and_drop_ag start":
            self.drag_and_drop_ag.start()
        elif command == "drag_and_drop_ag stop":
            self.drag_and_drop_ag.stop()

        elif command == "waiting_user_action_ag start":
            self.waiting_user_action_ag.start()
        elif command == "waiting_user_action_ag stop":
            self.waiting_user_action_ag.stop()

        elif command == "report_preparation_ag start":
            self.report_preparation_ag.start()
        elif command == "report_preparation_ag stop":
            self.report_preparation_ag.stop()

        elif command == "update_log_box_text":
            self.ui.plainTextEdit.setPlainText(self.ui.status_text)

    def set_status_bar_text(self, text, sec=5):
        self.ui.set_status_bar_text(text, sec=5)

    def add_text_to_log_box(self, text):
        self.ui.add_text_to_log_box(text)

    def change_last_log_box_text(self, text):
        self.ui.change_last_log_box_text(text)

    def clear_log_box(self):
        self.ui.clear_log_box()

    def refresh_raw_files_list(self, select_row_with_text=""):
        # Получим название текущего элемента:
        p_passed_no_name = False
        select_row_num = 0
        if select_row_with_text == "":
            p_passed_no_name = True
            # Если название не указано, то сохраним строку, которая выделена сейчас:
            select_row_with_text = self.ui.listViewRawData.currentIndex().data()
            select_row_num = self.ui.listViewRawData.currentIndex().row() + 1

        # Получим список файлов из папки с "сырыми" данными:
        rawdata_list = get_files_list(get_parameter_value(myconstants.RAW_DATA_SECTION_NAME))

        self.ui.model = QtGui.QStandardItemModel()
        self.ui.listViewRawData.setModel(self.ui.model)
        self.ui.model.removeRows(0, self.ui.model.rowCount())

        item2select = None
        counter = 0
        # Добавим в список все файлы, найденные в папке:
        for curr_file_name in rawdata_list:
            counter += 1
            item = QtGui.QStandardItem(curr_file_name)
            self.ui.model.appendRow(item)
            if curr_file_name == select_row_with_text:
                item2select = item
            if (counter == select_row_num) and (item2select is None):
                item2select = item

        if item2select is None:
            # Не выбран элемент.
            if p_passed_no_name and counter != 0:
            # Ни какой элемент не был передан в функцию для выбора.
            # То есть это случай когда исчез последний элемент
            # и количество строк стало меньше, а значит надо
            # выбрать тот что стал последним.
                item2select = self.ui.model.item(counter - 1)
            else:
                item2select = self.ui.model.item(0)

        self.ui.listViewRawData.setCurrentIndex(self.ui.model.indexFromItem(item2select))

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        processing_report = self.parent.reporter.report_creation_in_process
        processing_drag_and_drop = self.parent.drag_and_prop_in_process
        if processing_report or processing_drag_and_drop:
            # Ничего не делаем в это время.
            return

        # Из полученных файлов выберем те, которые обрабатывать не будем:
        not_xls_files = [u.toLocalFile() for u in event.mimeData().urls() if u.toLocalFile()[-5:].lower() != ".xlsx"]
        # Из тех файлов, которые "прилетели" выберем только *.xlsx:
        xls_files = [u.toLocalFile() for u in event.mimeData().urls() if u.toLocalFile()[-5:].lower() == ".xlsx"]

        if xls_files or not_xls_files:
            self.ui.clear_log_box()
        if not_xls_files:
            self.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)
            if len(not_xls_files) == 1:
                self.add_text_to_log_box("Исключен из обработки:")
                self.add_text_to_log_box(f"   {not_xls_files[0]}")
            else:
                self.add_text_to_log_box("Исключены из обработки:")

                for num, one_file in enumerate(not_xls_files):
                    self.add_text_to_log_box(f"   {num + 1}. {one_file}")

            self.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)

        if not xls_files:
            pass
        else:
            if not not_xls_files:
                self.add_text_to_log_box(myconstants.TEXT_LINES_SEPARATOR)

            copy_file_as_drop_process(self, xls_files)

    def showEvent(self, event):
        super(MyWindow, self).showEvent(event)
        self.ready_to_save_position = True

    def resizeEvent(self, event):
        super(MyWindow, self).resizeEvent(event)
        self.save_coordinates()

    def moveEvent(self, event):
        super(MyWindow, self).moveEvent(event)
        self.save_coordinates()
    
    def save_coordinates(self):
        if self.ready_to_save_position:
            data = self.saveGeometry()
            save_param(myconstants.PARAMETER_SAVED_MAIN_WINDOW_POZ, data)

            left_and_right_boxes_widths = [self.ui.layoutWidget.width(), self.ui.layoutWidget3.width()]
            top_and_bottom_boxes_widths = [self.ui.layoutWidget1.height(), self.ui.layoutWidget2.height()]
            save_param(myconstants.PARAMETER_SAVED_VALUE_LEFT_AND_RIGHT_BOXES, left_and_right_boxes_widths)
            save_param(myconstants.PARAMETER_SAVED_VALUE_TOP_AND_BOTTOM_BOXES, top_and_bottom_boxes_widths)

    def keyReleaseEvent(self, event):
        if event.key() == QtCore.Qt.Key_Control:
            self.ctrl_is_pressed = False
        if event.key() == QtCore.Qt.Key_Alt:
            self.alt_is_pressed = False

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Control:
            self.ctrl_is_pressed = True
        if event.key() == QtCore.Qt.Key_Alt:
            self.alt_is_pressed = True

        if event.key() == QtCore.Qt.Key_F12:
            self.f12_counter += 1
            if self.f12_counter >= myconstants.PARAMETER_TIMES_TO_PRESS_F12:
                self.f12_counter = 0
                self.setGeometry(
                    myconstants.PARAMETER_DEFAULT_MAIN_WINDOW_L,
                    myconstants.PARAMETER_DEFAULT_MAIN_WINDOW_T,
                    myconstants.PARAMETER_DEFAULT_MAIN_WINDOW_W,
                    myconstants.PARAMETER_DEFAULT_MAIN_WINDOW_H
                )
                
        event.accept()
        
    def closeEvent(self, e):
        result = None
        if not self.ctrl_is_pressed:
            result = QtWidgets.QMessageBox.question(self, "Подтверждение закрытия окна",
                                                    "Вы действительно хотите закрыть программу?\n\n" +
                                                    "Если у Вас формируется отчёт,\n" +
                                                    "то скорее всего, его формирование не прекратится.",
                                                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                                    QtWidgets.QMessageBox.No)
        if self.ctrl_is_pressed or result == QtWidgets.QMessageBox.Yes:
            self.ui.exit_in_process = True
            e.accept()
            QtWidgets.QMainWindow.closeEvent(self, e)
        else:
            e.ignore()
            self.set_status_bar_text("Отказ от закрытия программы")

if __name__ == "__main__":
    pass
