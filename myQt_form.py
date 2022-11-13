# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Qt5Project/Windows2.ui'
#
# Created by: PyQt5 UI code generator 5.15.7
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


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
        self.VerticalSplitter = QtWidgets.QSplitter(self.centralwidget)
        self.VerticalSplitter.setOrientation(QtCore.Qt.Horizontal)
        self.VerticalSplitter.setObjectName("VerticalSplitter")
        self.layoutWidget = QtWidgets.QWidget(self.VerticalSplitter)
        self.layoutWidget.setObjectName("layoutWidget")
        self.leftBox = QtWidgets.QVBoxLayout(self.layoutWidget)
        self.leftBox.setContentsMargins(0, 0, 0, 0)
        self.leftBox.setObjectName("leftBox")
        self.HorisontalSplitter = QtWidgets.QSplitter(self.layoutWidget)
        self.HorisontalSplitter.setOrientation(QtCore.Qt.Vertical)
        self.HorisontalSplitter.setObjectName("HorisontalSplitter")
        self.layoutWidget1 = QtWidgets.QWidget(self.HorisontalSplitter)
        self.layoutWidget1.setObjectName("layoutWidget1")
        self.topBox = QtWidgets.QVBoxLayout(self.layoutWidget1)
        self.topBox.setContentsMargins(0, 0, 0, 0)
        self.topBox.setSpacing(1)
        self.topBox.setObjectName("topBox")
        self.labelReports = QtWidgets.QLabel(self.layoutWidget1)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.labelReports.setFont(font)
        self.labelReports.setTextFormat(QtCore.Qt.AutoText)
        self.labelReports.setScaledContents(False)
        self.labelReports.setAlignment(QtCore.Qt.AlignCenter)
        self.labelReports.setWordWrap(False)
        self.labelReports.setObjectName("labelReports")
        self.topBox.addWidget(self.labelReports)
        self.listView = QtWidgets.QListView(self.layoutWidget1)
        self.listView.setMinimumSize(QtCore.QSize(200, 100))
        self.listView.setStyleSheet("")
        self.listView.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.listView.setObjectName("listView")
        self.topBox.addWidget(self.listView)
        self.layoutWidget2 = QtWidgets.QWidget(self.HorisontalSplitter)
        self.layoutWidget2.setObjectName("layoutWidget2")
        self.bottomBox = QtWidgets.QVBoxLayout(self.layoutWidget2)
        self.bottomBox.setContentsMargins(0, 0, 0, 0)
        self.bottomBox.setSpacing(1)
        self.bottomBox.setObjectName("bottomBox")
        self.labelRawData = QtWidgets.QLabel(self.layoutWidget2)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.labelRawData.setFont(font)
        self.labelRawData.setTextFormat(QtCore.Qt.AutoText)
        self.labelRawData.setScaledContents(False)
        self.labelRawData.setAlignment(QtCore.Qt.AlignCenter)
        self.labelRawData.setWordWrap(False)
        self.labelRawData.setObjectName("labelRawData")
        self.bottomBox.addWidget(self.labelRawData)
        self.listViewRawData = QtWidgets.QListView(self.layoutWidget2)
        self.listViewRawData.setMinimumSize(QtCore.QSize(200, 100))
        self.listViewRawData.setAcceptDrops(True)
        self.listViewRawData.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.listViewRawData.setDragEnabled(True)
        self.listViewRawData.setObjectName("listViewRawData")
        self.bottomBox.addWidget(self.listViewRawData)
        self.leftBox.addWidget(self.HorisontalSplitter)
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setSpacing(0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.line_3 = QtWidgets.QFrame(self.layoutWidget)
        self.line_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.verticalLayout.addWidget(self.line_3)
        self.label = QtWidgets.QLabel(self.layoutWidget)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.label_4 = QtWidgets.QLabel(self.layoutWidget)
        font = QtGui.QFont()
        font.setPointSize(2)
        self.label_4.setFont(font)
        self.label_4.setText("")
        self.label_4.setObjectName("label_4")
        self.verticalLayout.addWidget(self.label_4)
        self.radioButtonDD1 = QtWidgets.QRadioButton(self.layoutWidget)
        self.radioButtonDD1.setChecked(True)
        self.radioButtonDD1.setObjectName("radioButtonDD1")
        self.verticalLayout.addWidget(self.radioButtonDD1)
        self.radioButtonDD2 = QtWidgets.QRadioButton(self.layoutWidget)
        self.radioButtonDD2.setObjectName("radioButtonDD2")
        self.verticalLayout.addWidget(self.radioButtonDD2)
        self.radioButtonDD3 = QtWidgets.QRadioButton(self.layoutWidget)
        self.radioButtonDD3.setObjectName("radioButtonDD3")
        self.verticalLayout.addWidget(self.radioButtonDD3)
        self.radioButtonDD4 = QtWidgets.QRadioButton(self.layoutWidget)
        self.radioButtonDD4.setObjectName("radioButtonDD4")
        self.verticalLayout.addWidget(self.radioButtonDD4)
        self.line_2 = QtWidgets.QFrame(self.layoutWidget)
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.verticalLayout.addWidget(self.line_2)
        self.label_2 = QtWidgets.QLabel(self.layoutWidget)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.verticalLayout.addWidget(self.label_2)
        self.label_3 = QtWidgets.QLabel(self.layoutWidget)
        font = QtGui.QFont()
        font.setPointSize(2)
        self.label_3.setFont(font)
        self.label_3.setText("")
        self.label_3.setObjectName("label_3")
        self.verticalLayout.addWidget(self.label_3)
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
        self.checkBoxDeleteNotProduct.setObjectName("checkBoxDeleteNotProduct")
        self.verticalLayout.addWidget(self.checkBoxDeleteNotProduct)
        self.checkBoxOnlyProjectsWithAdd = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxOnlyProjectsWithAdd.setChecked(False)
        self.checkBoxOnlyProjectsWithAdd.setObjectName("checkBoxOnlyProjectsWithAdd")
        self.verticalLayout.addWidget(self.checkBoxOnlyProjectsWithAdd)
        self.comboBoxPGroups = QtWidgets.QComboBox(self.layoutWidget)
        self.comboBoxPGroups.setObjectName("comboBoxPGroups")
        self.verticalLayout.addWidget(self.comboBoxPGroups)
        self.checkBoxSelectUsers = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxSelectUsers.setObjectName("checkBoxSelectUsers")
        self.verticalLayout.addWidget(self.checkBoxSelectUsers)
        self.comboBoxSelectUsers = QtWidgets.QComboBox(self.layoutWidget)
        self.comboBoxSelectUsers.setObjectName("comboBoxSelectUsers")
        self.verticalLayout.addWidget(self.comboBoxSelectUsers)
        self.checkBoxDeleteWithoutFact = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBoxDeleteWithoutFact.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBoxDeleteWithoutFact.sizePolicy().hasHeightForWidth())
        self.checkBoxDeleteWithoutFact.setSizePolicy(sizePolicy)
        self.checkBoxDeleteWithoutFact.setChecked(True)
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
        self.checkBoxOpenExcel.setObjectName("checkBoxOpenExcel")
        self.verticalLayout.addWidget(self.checkBoxOpenExcel)
        self.line = QtWidgets.QFrame(self.layoutWidget)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.verticalLayout.addWidget(self.line)
        self.pushButtonDoIt = QtWidgets.QPushButton(self.layoutWidget)
        self.pushButtonDoIt.setEnabled(True)
        self.pushButtonDoIt.setMinimumSize(QtCore.QSize(200, 30))
        self.pushButtonDoIt.setAutoDefault(False)
        self.pushButtonDoIt.setFlat(False)
        self.pushButtonDoIt.setObjectName("pushButtonDoIt")
        self.verticalLayout.addWidget(self.pushButtonDoIt)
        self.leftBox.addLayout(self.verticalLayout)
        self.layoutWidget3 = QtWidgets.QWidget(self.VerticalSplitter)
        self.layoutWidget3.setObjectName("layoutWidget3")
        self.rightBox = QtWidgets.QVBoxLayout(self.layoutWidget3)
        self.rightBox.setContentsMargins(0, 0, 0, 0)
        self.rightBox.setSpacing(1)
        self.rightBox.setObjectName("rightBox")
        self.labelProgressStatus = QtWidgets.QLabel(self.layoutWidget3)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.labelProgressStatus.setFont(font)
        self.labelProgressStatus.setTextFormat(QtCore.Qt.AutoText)
        self.labelProgressStatus.setScaledContents(False)
        self.labelProgressStatus.setAlignment(QtCore.Qt.AlignCenter)
        self.labelProgressStatus.setWordWrap(False)
        self.labelProgressStatus.setObjectName("labelProgressStatus")
        self.rightBox.addWidget(self.labelProgressStatus)
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.layoutWidget3)
        self.plainTextEdit.setEnabled(True)
        self.plainTextEdit.setMinimumSize(QtCore.QSize(300, 0))
        self.plainTextEdit.setStyleSheet("color: rgb(0, 0, 0);")
        self.plainTextEdit.setReadOnly(True)
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.rightBox.addWidget(self.plainTextEdit)
        self.pushButtonOpenLastReport = QtWidgets.QPushButton(self.layoutWidget3)
        self.pushButtonOpenLastReport.setMinimumSize(QtCore.QSize(191, 30))
        self.pushButtonOpenLastReport.setCheckable(False)
        self.pushButtonOpenLastReport.setAutoDefault(False)
        self.pushButtonOpenLastReport.setFlat(False)
        self.pushButtonOpenLastReport.setObjectName("pushButtonOpenLastReport")
        self.rightBox.addWidget(self.pushButtonOpenLastReport)
        self.gridLayout.addWidget(self.VerticalSplitter, 0, 0, 1, 1)
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
        self.labelReports.setText(_translate("MainWindow", "Список доступных отчетов:"))
        self.labelRawData.setText(_translate("MainWindow", "Выгрузка данных из DES.LM:"))
        self.label.setText(_translate("MainWindow", "При обработке Drag&Drop:"))
        self.radioButtonDD1.setText(_translate("MainWindow", "только копировать"))
        self.radioButtonDD2.setText(_translate("MainWindow", "проверять структуру данных"))
        self.radioButtonDD3.setText(_translate("MainWindow", "... и переименовывать результат"))
        self.radioButtonDD4.setText(_translate("MainWindow", "... и переименовывать источник"))
        self.label_2.setText(_translate("MainWindow", "При формировании отчёта:"))
        self.checkBoxDeleteVIP.setText(_translate("MainWindow", "Убрать VIP"))
        self.checkBoxCurrMonthAHalf.setText(_translate("MainWindow", "Текущий месяц 50%"))
        self.checkBoxDeleteNotProduct.setText(_translate("MainWindow", "Оставить только производство"))
        self.checkBoxOnlyProjectsWithAdd.setText(_translate("MainWindow", "Оставить только проекты с доп иформацией"))
        self.checkBoxSelectUsers.setText(_translate("MainWindow", "Выбрать только людей из группы:"))
        self.checkBoxDeleteWithoutFact.setText(_translate("MainWindow", "Удалить строки с нулевым фактом"))
        self.checkBoxDelPDn.setText(_translate("MainWindow", "Удалить проекты с ПерсДанными"))
        self.checkBoxDeleteVac.setText(_translate("MainWindow", "Удалить данные о вакансиях из отчёта"))
        self.checkBoxAddVFTE.setText(_translate("MainWindow", "Добавить к данным искусственные FTE"))
        self.checkBoxSaveWithOutFotmulas.setText(_translate("MainWindow", "Сохранить только значения (без формул)"))
        self.checkBoxDelRawData.setText(_translate("MainWindow", "Удалить лист с данными в файле отчета"))
        self.checkBoxOpenExcel.setText(_translate("MainWindow", "Сразу открыть в Excel полученный отчет"))
        self.pushButtonDoIt.setText(_translate("MainWindow", "Сформировать"))
        self.labelProgressStatus.setText(_translate("MainWindow", "Прогресс выполнения обработки данных и подготовки отчета:"))
        self.plainTextEdit.setPlainText(_translate("MainWindow", "> Ожидание выбора пользователя."))
        self.pushButtonOpenLastReport.setText(_translate("MainWindow", "Открыть последний сформированный отчет в Excel"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
