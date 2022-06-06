# pyuic5 -x Qt5Project/Windows.ui -o _tmp_Qt5_form.py
# pyuic5 -x Qt5Project/_tmp_QLV.ui -o _tmp_QLV_form.py
#        MainWindow.setFixedSize(MainWindow.size().width(), MainWindow.size().height())
#        MainWindow.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)

import sys
import subprocess

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import pyqtSignal, QObject

import myconstants
from myutils import load_param, save_param

class Communicate(QObject):
    updateStatusText = pyqtSignal()

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setWindowModality(QtCore.Qt.ApplicationModal)
        MainWindow.resize(795, 578)
        MainWindow.setAcceptDrops(True)
        MainWindow.setStatusTip("")
        MainWindow.setWhatsThis("")
        MainWindow.setAccessibleDescription("")
        MainWindow.setAutoFillBackground(False)
        MainWindow.setToolButtonStyle(QtCore.Qt.ToolButtonTextUnderIcon)
        MainWindow.setAnimated(False)
        MainWindow.setDocumentMode(False)
        MainWindow.setTabShape(QtWidgets.QTabWidget.Rounded)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.checkBoxDeleteVac = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBoxDeleteVac.setEnabled(True)
        self.checkBoxDeleteVac.setGeometry(QtCore.QRect(20, 449, 241, 17))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDeleteVac.sizePolicy().hasHeightForWidth())
        self.checkBoxDeleteVac.setSizePolicy(sizePolicy)
        self.checkBoxDeleteVac.setChecked(True)
        self.checkBoxDeleteVac.setAutoRepeat(False)
        self.checkBoxDeleteVac.setObjectName("checkBoxDeleteVac")
        self.pushButtonDoIt = QtWidgets.QPushButton(self.centralwidget)
        self.pushButtonDoIt.setGeometry(QtCore.QRect(10, 541, 261, 31))
        self.pushButtonDoIt.setAutoDefault(False)
        self.pushButtonDoIt.setFlat(False)
        self.pushButtonDoIt.setObjectName("pushButtonDoIt")
        self.listView = QtWidgets.QListView(self.centralwidget)
        self.listView.setGeometry(QtCore.QRect(10, 20, 261, 231))
        self.listView.setStyleSheet("")
        self.listView.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.listView.setObjectName("listView")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 3, 251, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setTextFormat(QtCore.Qt.AutoText)
        self.label.setScaledContents(False)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setWordWrap(False)
        self.label.setObjectName("label")
        self.listViewRawData = QtWidgets.QListView(self.centralwidget)
        self.listViewRawData.setGeometry(QtCore.QRect(10, 277, 261, 121))
        self.listViewRawData.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.listViewRawData.setObjectName("listViewRawData")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(10, 260, 251, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setTextFormat(QtCore.Qt.AutoText)
        self.label_2.setScaledContents(False)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setWordWrap(False)
        self.label_2.setObjectName("label_2")
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit.setEnabled(True)
        self.plainTextEdit.setGeometry(QtCore.QRect(277, 20, 511, 551))
        self.plainTextEdit.setStyleSheet("color: rgb(0, 0, 0);")
        self.plainTextEdit.setReadOnly(True)
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(298, 3, 481, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setTextFormat(QtCore.Qt.AutoText)
        self.label_3.setScaledContents(False)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setWordWrap(False)
        self.label_3.setObjectName("label_3")
        self.checkBoxOpenExcel = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBoxOpenExcel.setEnabled(True)
        self.checkBoxOpenExcel.setGeometry(QtCore.QRect(20, 521, 241, 17))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxOpenExcel.sizePolicy().hasHeightForWidth())
        self.checkBoxOpenExcel.setSizePolicy(sizePolicy)
        self.checkBoxOpenExcel.setChecked(True)
        self.checkBoxOpenExcel.setAutoRepeat(False)
        self.checkBoxOpenExcel.setObjectName("checkBoxOpenExcel")
        self.pushButtonOpenLastReport = QtWidgets.QPushButton(self.centralwidget)
        self.pushButtonOpenLastReport.setGeometry(QtCore.QRect(277, 541, 511, 31))
        self.pushButtonOpenLastReport.setMinimumSize(QtCore.QSize(191, 23))
        self.pushButtonOpenLastReport.setCheckable(False)
        self.pushButtonOpenLastReport.setAutoDefault(False)
        self.pushButtonOpenLastReport.setFlat(False)
        self.pushButtonOpenLastReport.setObjectName("pushButtonOpenLastReport")
        self.checkBoxSaveWithOutFotmulas = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBoxSaveWithOutFotmulas.setEnabled(True)
        self.checkBoxSaveWithOutFotmulas.setGeometry(QtCore.QRect(20, 486, 241, 17))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxSaveWithOutFotmulas.sizePolicy().hasHeightForWidth())
        self.checkBoxSaveWithOutFotmulas.setSizePolicy(sizePolicy)
        self.checkBoxSaveWithOutFotmulas.setChecked(True)
        self.checkBoxSaveWithOutFotmulas.setAutoRepeat(False)
        self.checkBoxSaveWithOutFotmulas.setObjectName("checkBoxSaveWithOutFotmulas")
        self.checkBoxAddVFTE = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBoxAddVFTE.setEnabled(True)
        self.checkBoxAddVFTE.setGeometry(QtCore.QRect(20, 463, 241, 17))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxAddVFTE.sizePolicy().hasHeightForWidth())
        self.checkBoxAddVFTE.setSizePolicy(sizePolicy)
        self.checkBoxAddVFTE.setChecked(True)
        self.checkBoxAddVFTE.setAutoRepeat(False)
        self.checkBoxAddVFTE.setObjectName("checkBoxAddVFTE")
        self.checkBoxDelPDn = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBoxDelPDn.setEnabled(True)
        self.checkBoxDelPDn.setGeometry(QtCore.QRect(20, 434, 221, 17))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDelPDn.sizePolicy().hasHeightForWidth())
        self.checkBoxDelPDn.setSizePolicy(sizePolicy)
        self.checkBoxDelPDn.setChecked(True)
        self.checkBoxDelPDn.setAutoRepeat(False)
        self.checkBoxDelPDn.setObjectName("checkBoxDelPDn")
        self.checkBoxDeleteNotProduct = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBoxDeleteNotProduct.setEnabled(True)
        self.checkBoxDeleteNotProduct.setGeometry(QtCore.QRect(20, 420, 221, 17))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDeleteNotProduct.sizePolicy().hasHeightForWidth())
        self.checkBoxDeleteNotProduct.setSizePolicy(sizePolicy)
        self.checkBoxDeleteNotProduct.setChecked(True)
        self.checkBoxDeleteNotProduct.setAutoRepeat(False)
        self.checkBoxDeleteNotProduct.setObjectName("checkBoxDeleteNotProduct")
        self.checkBoxDelRawData = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBoxDelRawData.setEnabled(True)
        self.checkBoxDelRawData.setGeometry(QtCore.QRect(20, 500, 241, 17))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDelRawData.sizePolicy().hasHeightForWidth())
        self.checkBoxDelRawData.setSizePolicy(sizePolicy)
        self.checkBoxDelRawData.setChecked(True)
        self.checkBoxDelRawData.setAutoRepeat(False)
        self.checkBoxDelRawData.setObjectName("checkBoxDelRawData")
        self.checkBoxDeleteVIP = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBoxDeleteVIP.setEnabled(True)
        self.checkBoxDeleteVIP.setGeometry(QtCore.QRect(20, 400, 221, 21))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDeleteVIP.sizePolicy().hasHeightForWidth())
        self.checkBoxDeleteVIP.setSizePolicy(sizePolicy)
        self.checkBoxDeleteVIP.setChecked(True)
        self.checkBoxDeleteVIP.setAutoRepeat(False)
        self.checkBoxDeleteVIP.setObjectName("checkBoxDeleteVIP")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        MainWindow.setTabOrder(self.listView, self.listViewRawData)
        MainWindow.setTabOrder(self.listViewRawData, self.checkBoxDelPDn)
        MainWindow.setTabOrder(self.checkBoxDelPDn, self.checkBoxDeleteNotProduct)
        MainWindow.setTabOrder(self.checkBoxDeleteNotProduct, self.checkBoxDeleteVac)
        MainWindow.setTabOrder(self.checkBoxDeleteVac, self.checkBoxSaveWithOutFotmulas)
        MainWindow.setTabOrder(self.checkBoxSaveWithOutFotmulas, self.checkBoxAddVFTE)
        MainWindow.setTabOrder(self.checkBoxAddVFTE, self.plainTextEdit)
        MainWindow.setTabOrder(self.plainTextEdit, self.pushButtonDoIt)
        MainWindow.setTabOrder(self.pushButtonDoIt, self.checkBoxDelRawData)
        MainWindow.setTabOrder(self.checkBoxDelRawData, self.checkBoxOpenExcel)
        MainWindow.setTabOrder(self.checkBoxOpenExcel, self.pushButtonOpenLastReport)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "DES.LM.Reporter"))
        self.checkBoxDeleteVac.setText(_translate("MainWindow", "Удалить данные о вакансиях из отчёта"))
        self.pushButtonDoIt.setText(_translate("MainWindow", "Сформировать"))
        self.label.setText(_translate("MainWindow", "Список доступных отчетов:"))
        self.label_2.setText(_translate("MainWindow", "Выгрузка данных из DES.LM:"))
        self.plainTextEdit.setPlainText(_translate("MainWindow", "> Ожидание выбора пользователя."))
        self.label_3.setText(_translate("MainWindow", "Прогресс выполнения обработки данных и подготовки отчета:"))
        self.checkBoxOpenExcel.setText(_translate("MainWindow", "Сразу открыть в Excel полученный отчет"))
        self.pushButtonOpenLastReport.setText(_translate("MainWindow", "Открыть последний сформированный отчет в Excel"))
        self.checkBoxSaveWithOutFotmulas.setText(_translate("MainWindow", "Сохранить только значения (без формул)"))
        self.checkBoxAddVFTE.setText(_translate("MainWindow", "Добавить к данным искусственные FTE"))
        self.checkBoxDelPDn.setText(_translate("MainWindow", "Удалить проекты с ПерсДанными"))
        self.checkBoxDeleteNotProduct.setText(_translate("MainWindow", "Оставить только производство"))
        self.checkBoxDelRawData.setText(_translate("MainWindow", "Удалить лист с данными в файле отчета"))
        self.checkBoxDeleteVIP.setText(_translate("MainWindow", "Убрать VIP"))
    # ------------------------------------------------------------------- #

    closeApp = pyqtSignal()
    exit_in_process = False
    parent = None
    
    def setup_reports_list(self, reports_list=[]):
    
        self.model = QtGui.QStandardItemModel()
        self.listView.setModel(self.model)

        for one_report in reports_list:
            item = QtGui.QStandardItem(one_report)
            self.model.appendRow(item)
        
        item = self.model.item(load_param(myconstants.PARAMETER_SAVED_SELECTED_REPORT, 1) - 1)

        self.listView.setCurrentIndex(self.model.indexFromItem(item))
        
        return(True)

    def setup_rawdata_list(self, rawdata_list=[]):
    
        self.model = QtGui.QStandardItemModel()
        self.listViewRawData.setModel(self.model)

        for one_file in rawdata_list:
            item = QtGui.QStandardItem(one_file)
            self.model.appendRow(item)

        item = self.model.item(0)
        self.listViewRawData.setCurrentIndex(self.model.indexFromItem(item))
        
        return(True)

    def on_click_DoIt(self):
        raw_file_name = self.listViewRawData.currentIndex().data()
        report_file_name = self.listView.currentIndex().data()

        if not self.parent.parent.report_parameters.is_all_parametars_exist():
            return

        self.pushButtonDoIt.setEnabled(False)
        self.pushButtonOpenLastReport.setEnabled(False)
        self.resize_text_and_button()
        self.parent.parent.report_parameters.update(raw_file_name, report_file_name)
        save_param(myconstants.PARAMETER_SAVED_SELECTED_REPORT, self.reports_list.index(report_file_name) + 1)
        
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
            
        self.checkBoxDeleteVIP.setChecked(\
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_VIP, myconstants.PARAMETER_SAVED_VALUE_DELETE_VIP_DEFVALUE))
        self.checkBoxDeleteNotProduct.setChecked(\
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_NONPROD, myconstants.PARAMETER_SAVED_VALUE_DELETE_NONPROD_DEFVALUE))
        self.checkBoxDelPDn.setChecked(\
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_PERSDATA, myconstants.PARAMETER_SAVED_VALUE_DELETE_PERSDATA_DEFVALUE))
        self.checkBoxDeleteVac.setChecked(\
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_VAC, myconstants.PARAMETER_SAVED_VALUE_DELETE_VAC_DEFVALUE))
        self.checkBoxAddVFTE.setChecked(\
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_ADD_VFTE, myconstants.PARAMETER_SAVED_VALUE_ADD_VFTE_DEFVALUE))
        self.checkBoxSaveWithOutFotmulas.setChecked(\
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS, myconstants.PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS_DEFVALUE))
        self.checkBoxDelRawData.setChecked(\
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DEL_RAWSHEET, myconstants.PARAMETER_SAVED_VALUE_DEL_RAWSHEET_DEFVALUE))
        self.checkBoxOpenExcel.setChecked(\
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL, myconstants.PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL_DEFVALUE))
        
        self.checkBoxDelRawData.setVisible(self.checkBoxSaveWithOutFotmulas.isChecked())

    def setup_form(self, reports_list, raw_files_list):
        self.reports_list = reports_list
        self.pushButtonOpenLastReport.setVisible(False)

        self.setup_reports_list(reports_list)
        self.setup_rawdata_list(raw_files_list)

        self.pushButtonDoIt.clicked.connect(self.on_click_DoIt)
        
        self.checkBoxDeleteVIP.clicked.connect(self.on_click_CheckBoxes)
        self.checkBoxDeleteNotProduct.clicked.connect(self.on_click_CheckBoxes)
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
        slastreportfilename = load_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT)
        text_geometry = self.plainTextEdit.geometry()
        button_geometry = self.pushButtonOpenLastReport.geometry()
        x = text_geometry.left()
        y = text_geometry.top()
        w = text_geometry.width()
        if slastreportfilename == "":
            self.pushButtonOpenLastReport.setVisible(False)
            h = button_geometry.top() - text_geometry.top() + button_geometry.height() - 1
            self.plainTextEdit.setGeometry(QtCore.QRect(x, y, w, h))
        else:
            self.pushButtonOpenLastReport.setVisible(True)
            h = button_geometry.top() - text_geometry.top() - 5
            self.plainTextEdit.setGeometry(QtCore.QRect(x, y, w, h))
        
    def save_app_link(self, app):
        self.app = app
        
class MyWindow(QtWidgets.QMainWindow):
    ui = None
    
    def __init__(self, parent):
        self.parent = parent
        self._app = QtWidgets.QApplication(sys.argv)
        QtWidgets.QMainWindow.__init__(self, None)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.parent = self
        self.ui.save_app_link(self._app)
        self.setFixedSize(self.size().width(), self.size().height())
        
    def closeEvent(self, e):
        result = QtWidgets.QMessageBox.question(self, "Подтверждение закрытия окна", 
                                                "Вы действительно хотите закрыть программу?\n\nЕсли у Вас формируется отчёт,\nто скорее всего, его формирование не прекратится.", 
                                                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, 
                                                QtWidgets.QMessageBox.No)
        if result == QtWidgets.QMessageBox.Yes:
            self.ui.exit_in_process = True
            e.accept()
            QtWidgets.QMainWindow.closeEvent(self, e)
        else:
            e.ignore()

if __name__ == "__main__":
    pass
