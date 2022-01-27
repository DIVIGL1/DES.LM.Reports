# pyuic5 -x Qt5Project/Windows.ui -o _tmp_Qt5_form.py
#        MainWindow.setFixedSize(MainWindow.size().width(), MainWindow.size().height())
#        MainWindow.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)

import sys
import subprocess

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import pyqtSignal, QObject

import myconstants
import reportcreater
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
        self.checkBoxDeleteVac.setGeometry(QtCore.QRect(30, 493, 251, 17))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDeleteVac.sizePolicy().hasHeightForWidth())
        self.checkBoxDeleteVac.setSizePolicy(sizePolicy)
        self.checkBoxDeleteVac.setChecked(True)
        self.checkBoxDeleteVac.setAutoRepeat(False)
        self.checkBoxDeleteVac.setObjectName("checkBoxDeleteVac")
        self.pushButtonClose = QtWidgets.QPushButton(self.centralwidget)
        self.pushButtonClose.setGeometry(QtCore.QRect(10, 541, 75, 31))
        self.pushButtonClose.setObjectName("pushButtonClose")
        self.pushButtonDoIt = QtWidgets.QPushButton(self.centralwidget)
        self.pushButtonDoIt.setGeometry(QtCore.QRect(90, 541, 191, 31))
        self.pushButtonDoIt.setMinimumSize(QtCore.QSize(191, 23))
        self.pushButtonDoIt.setAutoDefault(False)
        self.pushButtonDoIt.setFlat(False)
        self.pushButtonDoIt.setObjectName("pushButtonDoIt")
        self.listView = QtWidgets.QListView(self.centralwidget)
        self.listView.setGeometry(QtCore.QRect(10, 20, 271, 251))
        self.listView.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.listView.setObjectName("listView")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 3, 261, 16))
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
        self.listViewRawData.setGeometry(QtCore.QRect(10, 290, 271, 201))
        self.listViewRawData.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.listViewRawData.setObjectName("listViewRawData")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(10, 273, 271, 16))
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
        self.plainTextEdit.setEnabled(False)
        self.plainTextEdit.setGeometry(QtCore.QRect(288, 20, 500, 551))
        self.plainTextEdit.setStyleSheet("color: rgb(0, 0, 0);")
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
        self.checkBoxOpenExcel.setGeometry(QtCore.QRect(30, 521, 251, 17))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxOpenExcel.sizePolicy().hasHeightForWidth())
        self.checkBoxOpenExcel.setSizePolicy(sizePolicy)
        self.checkBoxOpenExcel.setChecked(True)
        self.checkBoxOpenExcel.setAutoRepeat(False)
        self.checkBoxOpenExcel.setObjectName("checkBoxOpenExcel")
        self.pushButtonOpenLastReport = QtWidgets.QPushButton(self.centralwidget)
        self.pushButtonOpenLastReport.setGeometry(QtCore.QRect(287, 540, 501, 31))
        self.pushButtonOpenLastReport.setMinimumSize(QtCore.QSize(191, 23))
        self.pushButtonOpenLastReport.setCheckable(False)
        self.pushButtonOpenLastReport.setAutoDefault(False)
        self.pushButtonOpenLastReport.setFlat(False)
        self.pushButtonOpenLastReport.setObjectName("pushButtonOpenLastReport")
        self.checkBoxSaveWithOutFotmulas = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBoxSaveWithOutFotmulas.setEnabled(True)
        self.checkBoxSaveWithOutFotmulas.setGeometry(QtCore.QRect(30, 507, 251, 17))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxSaveWithOutFotmulas.sizePolicy().hasHeightForWidth())
        self.checkBoxSaveWithOutFotmulas.setSizePolicy(sizePolicy)
        self.checkBoxSaveWithOutFotmulas.setChecked(True)
        self.checkBoxSaveWithOutFotmulas.setAutoRepeat(False)
        self.checkBoxSaveWithOutFotmulas.setObjectName("checkBoxSaveWithOutFotmulas")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "DES.LM.Reporter"))
        self.checkBoxDeleteVac.setText(_translate("MainWindow", "Вакансии из отчета необходимо удалить"))
        self.pushButtonClose.setText(_translate("MainWindow", "Отменить"))
        self.pushButtonDoIt.setText(_translate("MainWindow", "Сформировать"))
        self.label.setText(_translate("MainWindow", "Список доступных отчетов:"))
        self.label_2.setText(_translate("MainWindow", "Выгрузка данных из DES.LM:"))
        self.plainTextEdit.setPlainText(_translate("MainWindow", "> Ожидание выбора пользователя."))
        self.label_3.setText(_translate("MainWindow", "Прогресс выполнения обработки данных и подготовки отчета:"))
        self.checkBoxOpenExcel.setText(_translate("MainWindow", "Сразу открыть отчет в Excel"))
        self.pushButtonOpenLastReport.setText(_translate("MainWindow", "Открыть последний сформированный отчет в Excel"))
        self.checkBoxSaveWithOutFotmulas.setText(_translate("MainWindow", "Сохранить только значения"))
        
    # ------------------------------------------------------------------- #

    closeApp = pyqtSignal()
    
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
        self.pushButtonDoIt.setEnabled(False)
        self.pushButtonClose.setEnabled(False)
        self.pushButtonOpenLastReport.setVisible(False)

        report_file_name = self.listView.currentIndex().data()
        raw_file_name = self.listViewRawData.currentIndex().data()

#        reportcreater.send_df_2_xls(report_file_name, raw_file_name, self.checkBoxDeleteVac.isChecked(), self.checkBoxOpenExcel.isChecked(), self)
        reportcreater.send_df_2_xls(report_file_name, raw_file_name, self)

    def on_click_Close(self):
        if self.pushButtonClose.isEnabled():
            self.pushButtonClose.setEnabled(False)
            self.pushButtonDoIt.setEnabled(False)
            self.app.quit()

    def on_dblClick_Reports_List(self):
        #save_param(myconstants.PARAMETER_SAVED_SELECTED_REPORT, self.checkBoxDeleteVac.isChecked())
        pass

    def on_click_OpenLastReport(self):
        last_report_filename = load_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT)
        if last_report_filename:
            subprocess.Popen(last_report_filename, shell=True)
    
    def on_click_CheckBoxes(self):
        save_param(myconstants.PARAMETER_SAVED_VALUE_DELETE_VAC, self.checkBoxDeleteVac.isChecked())
        save_param(myconstants.PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS, self.checkBoxSaveWithOutFotmulas.isChecked())
        save_param(myconstants.PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL, self.checkBoxOpenExcel.isChecked())

    def setup_form(self, reports_list, raw_files_list):
        self.reports_list = reports_list
        self.pushButtonOpenLastReport.setVisible(False)

        self.setup_reports_list(reports_list)
        self.setup_rawdata_list(raw_files_list)

        self.pushButtonDoIt.clicked.connect(self.on_click_DoIt)
        self.pushButtonClose.clicked.connect(self.on_click_Close)
        self.checkBoxDeleteVac.clicked.connect(self.on_click_CheckBoxes)
        self.checkBoxOpenExcel.clicked.connect(self.on_click_CheckBoxes)
        self.checkBoxSaveWithOutFotmulas.clicked.connect(self.on_click_CheckBoxes)
        self.listView.doubleClicked.connect(self.on_dblClick_Reports_List)
        
        self.checkBoxDeleteVac.setChecked(load_param(myconstants.PARAMETER_SAVED_VALUE_DELETE_VAC, True))
        self.checkBoxSaveWithOutFotmulas.setChecked(load_param(myconstants.PARAMETER_SAVED_VALUE_SAVE_WITHOUT_FORMULAS, False))
        self.checkBoxOpenExcel.setChecked(load_param(myconstants.PARAMETER_SAVED_VALUE_OPEN_IN_EXCEL, False))
        
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
        self.pushButtonClose.setEnabled(True)
        self.pushButtonDoIt.setEnabled(True)
        self.comminucate.updateStatusText.emit()
        self.pushButtonOpenLastReport.setVisible(load_param(myconstants.PARAMETER_FILENAME_OF_LAST_REPORT) != "")
        
    def save_app_link(self, app):
        self.app = app
        

def get_app_and_mainwindow():
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.save_app_link(app)
    MainWindow.setFixedSize(MainWindow.size().width(), MainWindow.size().height())
    MainWindow.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)

    return(app, ui, MainWindow)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
