# pyuic5 -x Qt5Project/Windows.ui -o _tmp_Qt5_form.py
# pyuic5 -x Qt5Project/Windows2.ui -o _tmp_Qt5_form.py
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

# ------------------------------------------------------------------- #
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setWindowModality(QtCore.Qt.ApplicationModal)
        MainWindow.resize(816, 650)
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
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setContentsMargins(5, 5, 5, 5)
        self.gridLayout.setObjectName("gridLayout")
        self.splitter_2 = QtWidgets.QSplitter(self.centralwidget)
        self.splitter_2.setOrientation(QtCore.Qt.Horizontal)
        self.splitter_2.setObjectName("splitter_2")
        self.layoutWidget = QtWidgets.QWidget(self.splitter_2)
        self.layoutWidget.setObjectName("layoutWidget")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout(self.layoutWidget)
        self.verticalLayout_5.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.splitter = QtWidgets.QSplitter(self.layoutWidget)
        self.splitter.setOrientation(QtCore.Qt.Vertical)
        self.splitter.setObjectName("splitter")
        self.layoutWidget1 = QtWidgets.QWidget(self.splitter)
        self.layoutWidget1.setObjectName("layoutWidget1")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.layoutWidget1)
        self.verticalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_3.setSpacing(1)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.label = QtWidgets.QLabel(self.layoutWidget1)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setTextFormat(QtCore.Qt.AutoText)
        self.label.setScaledContents(False)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setWordWrap(False)
        self.label.setObjectName("label")
        self.verticalLayout_3.addWidget(self.label)
        self.listView = QtWidgets.QListView(self.layoutWidget1)
        self.listView.setMinimumSize(QtCore.QSize(0, 260))
        self.listView.setStyleSheet("")
        self.listView.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.listView.setObjectName("listView")
        self.verticalLayout_3.addWidget(self.listView)
        self.layoutWidget2 = QtWidgets.QWidget(self.splitter)
        self.layoutWidget2.setObjectName("layoutWidget2")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.layoutWidget2)
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_2.setSpacing(1)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.layoutWidget2)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setTextFormat(QtCore.Qt.AutoText)
        self.label_2.setScaledContents(False)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setWordWrap(False)
        self.label_2.setObjectName("label_2")
        self.verticalLayout_2.addWidget(self.label_2)
        self.listViewRawData = QtWidgets.QListView(self.layoutWidget2)
        self.listViewRawData.setMinimumSize(QtCore.QSize(0, 100))
        self.listViewRawData.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.listViewRawData.setObjectName("listViewRawData")
        self.verticalLayout_2.addWidget(self.listViewRawData)
        self.verticalLayout_5.addWidget(self.splitter)
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setSpacing(0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.checkBoxDeleteVIP = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxDeleteVIP.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDeleteVIP.sizePolicy().hasHeightForWidth())
        self.checkBoxDeleteVIP.setSizePolicy(sizePolicy)
        self.checkBoxDeleteVIP.setMaximumSize(QtCore.QSize(100, 16777215))
        self.checkBoxDeleteVIP.setChecked(True)
        self.checkBoxDeleteVIP.setAutoRepeat(False)
        self.checkBoxDeleteVIP.setObjectName("checkBoxDeleteVIP")
        self.horizontalLayout.addWidget(self.checkBoxDeleteVIP)
        self.checkBoxCurrMonthAHalf = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxCurrMonthAHalf.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxCurrMonthAHalf.sizePolicy().hasHeightForWidth())
        self.checkBoxCurrMonthAHalf.setSizePolicy(sizePolicy)
        self.checkBoxCurrMonthAHalf.setChecked(True)
        self.checkBoxCurrMonthAHalf.setAutoRepeat(False)
        self.checkBoxCurrMonthAHalf.setObjectName("checkBoxCurrMonthAHalf")
        self.horizontalLayout.addWidget(self.checkBoxCurrMonthAHalf)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.checkBoxDeleteNotProduct = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxDeleteNotProduct.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDeleteNotProduct.sizePolicy().hasHeightForWidth())
        self.checkBoxDeleteNotProduct.setSizePolicy(sizePolicy)
        self.checkBoxDeleteNotProduct.setChecked(True)
        self.checkBoxDeleteNotProduct.setAutoRepeat(False)
        self.checkBoxDeleteNotProduct.setObjectName("checkBoxDeleteNotProduct")
        self.verticalLayout.addWidget(self.checkBoxDeleteNotProduct)
        self.checkBoxDeleteWithoutFact = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxDeleteWithoutFact.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDeleteWithoutFact.sizePolicy().hasHeightForWidth())
        self.checkBoxDeleteWithoutFact.setSizePolicy(sizePolicy)
        self.checkBoxDeleteWithoutFact.setChecked(True)
        self.checkBoxDeleteWithoutFact.setAutoRepeat(False)
        self.checkBoxDeleteWithoutFact.setObjectName("checkBoxDeleteWithoutFact")
        self.verticalLayout.addWidget(self.checkBoxDeleteWithoutFact)
        self.checkBoxDelPDn = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxDelPDn.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDelPDn.sizePolicy().hasHeightForWidth())
        self.checkBoxDelPDn.setSizePolicy(sizePolicy)
        self.checkBoxDelPDn.setChecked(True)
        self.checkBoxDelPDn.setAutoRepeat(False)
        self.checkBoxDelPDn.setObjectName("checkBoxDelPDn")
        self.verticalLayout.addWidget(self.checkBoxDelPDn)
        self.checkBoxDeleteVac = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxDeleteVac.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDeleteVac.sizePolicy().hasHeightForWidth())
        self.checkBoxDeleteVac.setSizePolicy(sizePolicy)
        self.checkBoxDeleteVac.setChecked(True)
        self.checkBoxDeleteVac.setAutoRepeat(False)
        self.checkBoxDeleteVac.setObjectName("checkBoxDeleteVac")
        self.verticalLayout.addWidget(self.checkBoxDeleteVac)
        self.checkBoxAddVFTE = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxAddVFTE.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxAddVFTE.sizePolicy().hasHeightForWidth())
        self.checkBoxAddVFTE.setSizePolicy(sizePolicy)
        self.checkBoxAddVFTE.setChecked(True)
        self.checkBoxAddVFTE.setAutoRepeat(False)
        self.checkBoxAddVFTE.setObjectName("checkBoxAddVFTE")
        self.verticalLayout.addWidget(self.checkBoxAddVFTE)
        self.checkBoxSaveWithOutFotmulas = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxSaveWithOutFotmulas.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxSaveWithOutFotmulas.sizePolicy().hasHeightForWidth())
        self.checkBoxSaveWithOutFotmulas.setSizePolicy(sizePolicy)
        self.checkBoxSaveWithOutFotmulas.setChecked(True)
        self.checkBoxSaveWithOutFotmulas.setAutoRepeat(False)
        self.checkBoxSaveWithOutFotmulas.setObjectName("checkBoxSaveWithOutFotmulas")
        self.verticalLayout.addWidget(self.checkBoxSaveWithOutFotmulas)
        self.checkBoxDelRawData = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxDelRawData.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDelRawData.sizePolicy().hasHeightForWidth())
        self.checkBoxDelRawData.setSizePolicy(sizePolicy)
        self.checkBoxDelRawData.setChecked(True)
        self.checkBoxDelRawData.setAutoRepeat(False)
        self.checkBoxDelRawData.setObjectName("checkBoxDelRawData")
        self.verticalLayout.addWidget(self.checkBoxDelRawData)
        self.checkBoxOpenExcel = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxOpenExcel.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxOpenExcel.sizePolicy().hasHeightForWidth())
        self.checkBoxOpenExcel.setSizePolicy(sizePolicy)
        self.checkBoxOpenExcel.setChecked(True)
        self.checkBoxOpenExcel.setAutoRepeat(False)
        self.checkBoxOpenExcel.setObjectName("checkBoxOpenExcel")
        self.verticalLayout.addWidget(self.checkBoxOpenExcel)
        self.pushButtonDoIt = QtWidgets.QPushButton(self.layoutWidget)
        self.pushButtonDoIt.setEnabled(True)
        self.pushButtonDoIt.setMinimumSize(QtCore.QSize(295, 30))
        self.pushButtonDoIt.setAutoDefault(False)
        self.pushButtonDoIt.setFlat(False)
        self.pushButtonDoIt.setObjectName("pushButtonDoIt")
        self.verticalLayout.addWidget(self.pushButtonDoIt)
        self.verticalLayout_5.addLayout(self.verticalLayout)
        self.layoutWidget3 = QtWidgets.QWidget(self.splitter_2)
        self.layoutWidget3.setObjectName("layoutWidget3")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.layoutWidget3)
        self.verticalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_4.setSpacing(1)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.label_3 = QtWidgets.QLabel(self.layoutWidget3)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setTextFormat(QtCore.Qt.AutoText)
        self.label_3.setScaledContents(False)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setWordWrap(False)
        self.label_3.setObjectName("label_3")
        self.verticalLayout_4.addWidget(self.label_3)
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.layoutWidget3)
        self.plainTextEdit.setEnabled(True)
        self.plainTextEdit.setMinimumSize(QtCore.QSize(500, 0))
        self.plainTextEdit.setStyleSheet("color: rgb(0, 0, 0);")
        self.plainTextEdit.setReadOnly(True)
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.verticalLayout_4.addWidget(self.plainTextEdit)
        self.pushButtonOpenLastReport = QtWidgets.QPushButton(self.layoutWidget3)
        self.pushButtonOpenLastReport.setMinimumSize(QtCore.QSize(191, 30))
        self.pushButtonOpenLastReport.setCheckable(False)
        self.pushButtonOpenLastReport.setAutoDefault(False)
        self.pushButtonOpenLastReport.setFlat(False)
        self.pushButtonOpenLastReport.setObjectName("pushButtonOpenLastReport")
        self.verticalLayout_4.addWidget(self.pushButtonOpenLastReport)
        self.gridLayout.addWidget(self.splitter_2, 0, 0, 1, 1)
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
        self.label.setText(_translate("MainWindow", "Список доступных отчетов:"))
        self.label_2.setText(_translate("MainWindow", "Выгрузка данных из DES.LM:"))
        self.checkBoxDeleteVIP.setText(_translate("MainWindow", "Убрать VIP"))
        self.checkBoxCurrMonthAHalf.setText(_translate("MainWindow", "Текущий месяц 50%"))
        self.checkBoxDeleteNotProduct.setText(_translate("MainWindow", "Оставить только производство"))
        self.checkBoxDeleteWithoutFact.setText(_translate("MainWindow", "Удалить строки с нулевым фактом"))
        self.checkBoxDelPDn.setText(_translate("MainWindow", "Удалить проекты с ПерсДанными"))
        self.checkBoxDeleteVac.setText(_translate("MainWindow", "Удалить данные о вакансиях из отчёта"))
        self.checkBoxAddVFTE.setText(_translate("MainWindow", "Добавить к данным искусственные FTE"))
        self.checkBoxSaveWithOutFotmulas.setText(_translate("MainWindow", "Сохранить только значения (без формул)"))
        self.checkBoxDelRawData.setText(_translate("MainWindow", "Удалить лист с данными в файле отчета"))
        self.checkBoxOpenExcel.setText(_translate("MainWindow", "Сразу открыть в Excel полученный отчет"))
        self.pushButtonDoIt.setText(_translate("MainWindow", "Сформировать"))
        self.label_3.setText(_translate("MainWindow", "Прогресс выполнения обработки данных и подготовки отчета:"))
        self.plainTextEdit.setPlainText(_translate("MainWindow", "> Ожидание выбора пользователя."))
        self.pushButtonOpenLastReport.setText(_translate("MainWindow", "Открыть последний сформированный отчет в Excel"))
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
        
        return (True)

    def setup_rawdata_list(self, rawdata_list=[]):
    
        self.model = QtGui.QStandardItemModel()
        self.listViewRawData.setModel(self.model)

        for one_file in rawdata_list:
            item = QtGui.QStandardItem(one_file)
            self.model.appendRow(item)

        item = self.model.item(0)
        self.listViewRawData.setCurrentIndex(self.model.indexFromItem(item))
        
        return (True)

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
            
        self.checkBoxDeleteVIP.setChecked(\
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_VIP, myconstants.PARAMETER_SAVED_VALUE_DELETE_VIP_DEFVALUE))
        self.checkBoxDeleteNotProduct.setChecked(\
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_NONPROD, myconstants.PARAMETER_SAVED_VALUE_DELETE_NONPROD_DEFVALUE))
        self.checkBoxDeleteWithoutFact.setChecked(\
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_EMPTYFACT, myconstants.PARAMETER_SAVED_VALUE_DELETE_EMPTYFACT_DEFVALUE))
        self.checkBoxCurrMonthAHalf.setChecked(\
            load_param(s_preff + myconstants.PARAMETER_SAVED_VALUE_DELETE_CURRMONTHHALF, myconstants.PARAMETER_SAVED_VALUE_DELETE_CURRMONTHHALF_DEFVALUE))
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
        if slastreportfilename == "":
            self.pushButtonOpenLastReport.setVisible(False)
        else:
            self.pushButtonOpenLastReport.setVisible(True)
        
    def save_app_link(self, app):
        self.app = app
        
class MyWindow(QtWidgets.QMainWindow):
    ui = None
    esc_counter = 0
    start_flag = True
    
    def __init__(self, parent):
        self.parent = parent
        self._app = QtWidgets.QApplication(sys.argv)
        QtWidgets.QMainWindow.__init__(self, None)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.parent = self
        self.ui.save_app_link(self._app)

    def resizeEvent(self, event):
        super(MyWindow, self).resizeEvent(event)
        self.save_coordinats()

    def moveEvent(self, event):
        super(MyWindow, self).moveEvent(event)
        self.save_coordinats()
    
    def save_coordinats(self):
        if self.start_flag:
            self.start_flag = False
            data = load_param(myconstants.PARAMETER_SAVED_MAIN_WINDOW_POZ, "")
            if data:
                self.restoreGeometry(data)
        else:
            data = self.saveGeometry()
            save_param(myconstants.PARAMETER_SAVED_MAIN_WINDOW_POZ, data)
        
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
