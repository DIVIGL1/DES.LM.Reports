# pyuic5 -x Qt5Project/Windows2.ui -o myQt_form.py
# pyuic5 -x Qt5Project/_tmp_QLV.ui -o _tmp_QLV_form.py

import os
import sys
import subprocess
import shutil

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import pyqtSignal, QObject

import myconstants
import myQt_form
from mytablefuncs import get_parameter_value
from myutils import load_param, save_param, get_files_list


class Communicate(QObject):
    updateStatusText = pyqtSignal()

class qtMainWindow(myQt_form.Ui_MainWindow):
    closeApp = pyqtSignal()
    exit_in_process = None
    parent = None

    def setup_reports_list(self, reports_list=None):
        if reports_list is None:
            reports_list = []
    
        self.model = QtGui.QStandardItemModel()
        self.listView.setModel(self.model)

        for one_report in reports_list:
            item = QtGui.QStandardItem(one_report)
            self.model.appendRow(item)
        
        item = self.model.item(load_param(myconstants.PARAMETER_SAVED_SELECTED_REPORT, 1) - 1)

        self.listView.setCurrentIndex(self.model.indexFromItem(item))

        return (True)

    def on_click_DoIt(self):
        raw_file_name = self.listViewRawData.currentIndex().data()
        report_file_name = self.listView.currentIndex().data()

        if not self.parent.parent.report_parameters.is_all_parameters_exist():
            return

        self.pushButtonDoIt.setEnabled(False)
        self.pushButtonOpenLastReport.setEnabled(False)
        self.resize_text_and_button()
        self.parent.parent.report_parameters.update(raw_file_name, report_file_name)
        save_param(myconstants.PARAMETER_SAVED_SELECTED_REPORT, self.reports_list.index(report_file_name) + 1)
        save_param(myconstants.PARAMETER_SAVED_SELECTED_RAW_FILE, raw_file_name)

        self.parent.parent.reporter.create_report()

    def on_dblClick_Reports_List(self):
        self.on_click_DoIt()

    def on_click_OpenLastReport(self):
        last_report_filename = load_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT)
        if last_report_filename:
            subprocess.Popen(last_report_filename, shell=True)
    
    def on_click_CheckBoxes(self):
        if self.listView.currentIndex().data() is None:
            return
        
        s_preff = self.listView.currentIndex().data() + " --> "
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_VIP, self.checkBoxDeleteVIP.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_NONPROD, self.checkBoxDeleteNotProduct.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_EMPTYFACT, self.checkBoxDeleteWithoutFact.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_CURRMONTHHALF, self.checkBoxCurrMonthAHalf.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_PERSDATA, self.checkBoxDelPDn.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_VAC, self.checkBoxDeleteVac.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_ADD_VFTE, self.checkBoxAddVFTE.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS, self.checkBoxSaveWithOutFotmulas.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DEL_RAWSHEET, self.checkBoxDelRawData.isChecked())
        save_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL, self.checkBoxOpenExcel.isChecked())
        
        self.checkBoxDelRawData.setVisible(self.checkBoxSaveWithOutFotmulas.isChecked())

    def on_Click_Reports_List(self):
        self.setup_check_boxes()
        
    def setup_check_boxes(self):
        if self.listView.currentIndex().data() is None:
            s_preff = ""
        else:
            s_preff = self.listView.currentIndex().data() + " --> "
            
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

    def setup_form(self, reports_list):
        self.reports_list = reports_list
        self.pushButtonOpenLastReport.setVisible(False)

        self.setup_reports_list(reports_list)
        last_raw_file = load_param(myconstants.PARAMETER_SAVED_SELECTED_RAW_FILE, "")

        self.parent.refresh_raw_files_list(last_raw_file)

        self.pushButtonDoIt.clicked.connect(self.on_click_DoIt)
        
        self.checkBoxDeleteVIP.clicked.connect(self.on_click_CheckBoxes)
        self.checkBoxDeleteNotProduct.clicked.connect(self.on_click_CheckBoxes)
        self.checkBoxDeleteWithoutFact.clicked.connect(self.on_click_CheckBoxes)
        self.checkBoxCurrMonthAHalf.clicked.connect(self.on_click_CheckBoxes)
        self.checkBoxDelPDn.clicked.connect(self.on_click_CheckBoxes)
        self.checkBoxDeleteVac.clicked.connect(self.on_click_CheckBoxes)
        self.checkBoxAddVFTE.clicked.connect(self.on_click_CheckBoxes)
        self.checkBoxOpenExcel.clicked.connect(self.on_click_CheckBoxes)
        self.checkBoxSaveWithOutFotmulas.clicked.connect(self.on_click_CheckBoxes)
        self.checkBoxDelRawData.clicked.connect(self.on_click_CheckBoxes)
        
        self.setup_check_boxes()

        self.listView.clicked.connect(self.on_Click_Reports_List)
        
        self.listView.doubleClicked.connect(self.on_dblClick_Reports_List)
        self.listViewRawData.doubleClicked.connect(self.on_dblClick_Reports_List)

        self.pushButtonOpenLastReport.clicked.connect(self.on_click_OpenLastReport)
        self.status_text = ""
        self.previous_status_text = ""
        self.comminucate = Communicate()
        self.comminucate.updateStatusText.connect(self.update_status)

        self.listViewRawData.setDragEnabled(True)

    
    def update_status(self):
        self.plainTextEdit.setPlainText(self.status_text)
    
    def set_status(self, status_text):
        start_text_value = self.status_text
        self.previous_status_text = self.status_text
        self.status_text = start_text_value + ("\n" if start_text_value != "" else "") + status_text
        self.comminucate.updateStatusText.emit()

    def change_last_status_line(self, status_text):
        start_text_value = self.previous_status_text
        self.status_text = start_text_value + ("\n" if start_text_value != "" else "") + status_text
        self.comminucate.updateStatusText.emit()

    def clear_status(self):
        self.status_text = ""
        self.previous_status_text = ""
        self.comminucate.updateStatusText.emit()
        
    def enable_buttons(self):
        if self.exit_in_process:
            # Форма закрыта во время формирования отчёта - кнопок уже нет.
            return
        
        self.pushButtonDoIt.setEnabled(True)
        self.pushButtonOpenLastReport.setEnabled(True)
        
        self.comminucate.updateStatusText.emit()
        
        self.resize_text_and_button()
    
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

    def __init__(self, parent):
        self.parent = parent
        self.app = QtWidgets.QApplication(sys.argv)
        QtWidgets.QMainWindow.__init__(self, None)
        self.ui = qtMainWindow()
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

    def refresh_raw_files_list(self, select_row_with_text=""):
        # Получим название текущего элемента:
        if select_row_with_text == "":
            # Есл название не указано, то берём сейчас выделенную строку:
            select_row_with_text = self.ui.listViewRawData.currentIndex().data()
        # Получим список файлов из папки с "сырыми" данными:
        rawdata_list = get_files_list(get_parameter_value(myconstants.RAW_DATA_SECTION_NAME))

        # ???
        self.ui.model = QtGui.QStandardItemModel()
        self.ui.listViewRawData.setModel(self.ui.model)

        item2select = None
        # Добавим в список все файлы, найденные в папке:
        for curr_file_name in rawdata_list:
            item = QtGui.QStandardItem(curr_file_name)
            self.ui.model.appendRow(item)
            if curr_file_name == select_row_with_text:
                item2select = item
        else:
            if item2select is None:
                item2select = self.ui.model.item(0)

        self.ui.listViewRawData.setCurrentIndex(self.ui.model.indexFromItem(item2select))

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        for one_file_path in files:
            this_file_name = os.path.basename(one_file_path)
            raw_file_path = get_parameter_value(myconstants.RAW_DATA_SECTION_NAME) + "/" + this_file_name
            if os.path.isfile(raw_file_path):
                result = QtWidgets.QMessageBox.question(self, "Заменить файл?",
                                                        "В папке, где находятся данные, выгруженные из DES.LM" +
                                                        f"Файл с таким названием уже есть {this_file_name}\n\n" +
                                                        "Вы действительно хотите переписать его новым файлом?",
                                                        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                                        QtWidgets.QMessageBox.No)

                if result == QtWidgets.QMessageBox.No:
                    continue

            try:
                shutil.copyfile(one_file_path, raw_file_path)
                new_filename = os.path.splitext(os.path.basename(raw_file_path))[0]
                if len(files) == 1:
                    self.refresh_raw_files_list(new_filename)
            except (OSError, shutil.Error):
                QtWidgets.QMessageBox.question(self, "Ошибка копирования?",
                                               "Не удалось скопировать файл с данными выгруженными из DES.LM.",
                                               QtWidgets.QMessageBox.Yes)

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

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Control:
            self.ctrl_is_pressed = True

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
                                                    "Вы действительно хотите закрыть программу?\n\n"+
                                                    "Если у Вас формируется отчёт,\n"+
                                                    "то скорее всего, его формирование не прекратится.",
                                                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                                    QtWidgets.QMessageBox.No)
        if self.ctrl_is_pressed or result == QtWidgets.QMessageBox.Yes:
            self.ui.exit_in_process = True
            e.accept()
            QtWidgets.QMainWindow.closeEvent(self, e)
        else:
            e.ignore()


if __name__ == "__main__":
    pass
