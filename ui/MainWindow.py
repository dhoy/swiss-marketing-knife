# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\MainWindow.ui'
#
# Created by: PyQt5 UI code generator 5.4.1
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(609, 638)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/image/swiss-marketing-knife.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setStyleSheet("QPushButton\n"
"{\n"
"    font-weight: bold;\n"
"}\n"
"QLabel{font-weight: bold;}\n"
"")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.tab_widget = QtWidgets.QTabWidget(self.centralwidget)
        font = QtGui.QFont()
        font.setBold(False)
        font.setWeight(50)
        self.tab_widget.setFont(font)
        self.tab_widget.setContextMenuPolicy(QtCore.Qt.PreventContextMenu)
        self.tab_widget.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.tab_widget.setObjectName("tab_widget")
        self.tab_update = QtWidgets.QWidget()
        self.tab_update.setObjectName("tab_update")
        self.verticalLayout_11 = QtWidgets.QVBoxLayout(self.tab_update)
        self.verticalLayout_11.setObjectName("verticalLayout_11")
        self.verticalLayout_8 = QtWidgets.QVBoxLayout()
        self.verticalLayout_8.setObjectName("verticalLayout_8")
        self.label_3 = QtWidgets.QLabel(self.tab_update)
        self.label_3.setObjectName("label_3")
        self.verticalLayout_8.addWidget(self.label_3)
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.le_match = QtWidgets.QLineEdit(self.tab_update)
        self.le_match.setObjectName("le_match")
        self.horizontalLayout_7.addWidget(self.le_match)
        self.btn_open_dmc = QtWidgets.QPushButton(self.tab_update)
        self.btn_open_dmc.setMaximumSize(QtCore.QSize(25, 16777215))
        self.btn_open_dmc.setObjectName("btn_open_dmc")
        self.horizontalLayout_7.addWidget(self.btn_open_dmc)
        self.verticalLayout_8.addLayout(self.horizontalLayout_7)
        self.btn_import = QtWidgets.QPushButton(self.tab_update)
        self.btn_import.setMaximumSize(QtCore.QSize(75, 16777215))
        self.btn_import.setObjectName("btn_import")
        self.verticalLayout_8.addWidget(self.btn_import)
        self.verticalLayout_11.addLayout(self.verticalLayout_8)
        spacerItem = QtWidgets.QSpacerItem(20, 25, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum)
        self.verticalLayout_11.addItem(spacerItem)
        self.verticalLayout_7 = QtWidgets.QVBoxLayout()
        self.verticalLayout_7.setObjectName("verticalLayout_7")
        self.lbl_match_importing_update = QtWidgets.QLabel(self.tab_update)
        self.lbl_match_importing_update.setObjectName("lbl_match_importing_update")
        self.verticalLayout_7.addWidget(self.lbl_match_importing_update)
        self.pb_import_update = QtWidgets.QProgressBar(self.tab_update)
        self.pb_import_update.setProperty("value", -1)
        self.pb_import_update.setAlignment(QtCore.Qt.AlignCenter)
        self.pb_import_update.setObjectName("pb_import_update")
        self.verticalLayout_7.addWidget(self.pb_import_update)
        self.verticalLayout_11.addLayout(self.verticalLayout_7)
        spacerItem1 = QtWidgets.QSpacerItem(20, 25, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum)
        self.verticalLayout_11.addItem(spacerItem1)
        self.verticalLayout_9 = QtWidgets.QVBoxLayout()
        self.verticalLayout_9.setObjectName("verticalLayout_9")
        self.lbl_matching_update = QtWidgets.QLabel(self.tab_update)
        self.lbl_matching_update.setObjectName("lbl_matching_update")
        self.verticalLayout_9.addWidget(self.lbl_matching_update)
        self.pb_match_update = QtWidgets.QProgressBar(self.tab_update)
        self.pb_match_update.setEnabled(True)
        self.pb_match_update.setProperty("value", -1)
        self.pb_match_update.setAlignment(QtCore.Qt.AlignCenter)
        self.pb_match_update.setObjectName("pb_match_update")
        self.verticalLayout_9.addWidget(self.pb_match_update)
        self.verticalLayout_11.addLayout(self.verticalLayout_9)
        spacerItem2 = QtWidgets.QSpacerItem(20, 25, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum)
        self.verticalLayout_11.addItem(spacerItem2)
        self.verticalLayout_10 = QtWidgets.QVBoxLayout()
        self.verticalLayout_10.setObjectName("verticalLayout_10")
        self.lbl_match = QtWidgets.QLabel(self.tab_update)
        self.lbl_match.setEnabled(True)
        self.lbl_match.setObjectName("lbl_match")
        self.verticalLayout_10.addWidget(self.lbl_match)
        self.pb_match_camp = QtWidgets.QProgressBar(self.tab_update)
        self.pb_match_camp.setEnabled(True)
        self.pb_match_camp.setProperty("value", -1)
        self.pb_match_camp.setAlignment(QtCore.Qt.AlignCenter)
        self.pb_match_camp.setObjectName("pb_match_camp")
        self.verticalLayout_10.addWidget(self.pb_match_camp)
        self.verticalLayout_11.addLayout(self.verticalLayout_10)
        spacerItem3 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_11.addItem(spacerItem3)
        self.tab_widget.addTab(self.tab_update, "")
        self.tab_cross = QtWidgets.QWidget()
        self.tab_cross.setObjectName("tab_cross")
        self.verticalLayout_13 = QtWidgets.QVBoxLayout(self.tab_cross)
        self.verticalLayout_13.setObjectName("verticalLayout_13")
        self.verticalLayout_12 = QtWidgets.QVBoxLayout()
        self.verticalLayout_12.setObjectName("verticalLayout_12")
        self.label_4 = QtWidgets.QLabel(self.tab_cross)
        self.label_4.setObjectName("label_4")
        self.verticalLayout_12.addWidget(self.label_4)
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.le_cross_country = QtWidgets.QLineEdit(self.tab_cross)
        self.le_cross_country.setObjectName("le_cross_country")
        self.horizontalLayout_8.addWidget(self.le_cross_country)
        self.btn_open_cc = QtWidgets.QPushButton(self.tab_cross)
        self.btn_open_cc.setMaximumSize(QtCore.QSize(25, 16777215))
        self.btn_open_cc.setObjectName("btn_open_cc")
        self.horizontalLayout_8.addWidget(self.btn_open_cc)
        self.verticalLayout_12.addLayout(self.horizontalLayout_8)
        self.btn_process_cc = QtWidgets.QPushButton(self.tab_cross)
        self.btn_process_cc.setMaximumSize(QtCore.QSize(75, 16777215))
        self.btn_process_cc.setObjectName("btn_process_cc")
        self.verticalLayout_12.addWidget(self.btn_process_cc)
        self.verticalLayout_13.addLayout(self.verticalLayout_12)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.verticalLayout_13.addLayout(self.verticalLayout_2)
        spacerItem4 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.verticalLayout_13.addItem(spacerItem4)
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.lbl_process_cc = QtWidgets.QLabel(self.tab_cross)
        self.lbl_process_cc.setObjectName("lbl_process_cc")
        self.verticalLayout_3.addWidget(self.lbl_process_cc)
        self.pb_process_cc = QtWidgets.QProgressBar(self.tab_cross)
        self.pb_process_cc.setProperty("value", -1)
        self.pb_process_cc.setObjectName("pb_process_cc")
        self.verticalLayout_3.addWidget(self.pb_process_cc)
        self.verticalLayout_13.addLayout(self.verticalLayout_3)
        spacerItem5 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.verticalLayout_13.addItem(spacerItem5)
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.lbl_upload_cc = QtWidgets.QLabel(self.tab_cross)
        self.lbl_upload_cc.setObjectName("lbl_upload_cc")
        self.verticalLayout_4.addWidget(self.lbl_upload_cc)
        self.pb_upload_cc = QtWidgets.QProgressBar(self.tab_cross)
        self.pb_upload_cc.setProperty("value", -1)
        self.pb_upload_cc.setObjectName("pb_upload_cc")
        self.verticalLayout_4.addWidget(self.pb_upload_cc)
        self.verticalLayout_13.addLayout(self.verticalLayout_4)
        spacerItem6 = QtWidgets.QSpacerItem(20, 320, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum)
        self.verticalLayout_13.addItem(spacerItem6)
        self.tab_widget.addTab(self.tab_cross, "")
        self.tab_autoscale = QtWidgets.QWidget()
        self.tab_autoscale.setObjectName("tab_autoscale")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout(self.tab_autoscale)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.wv_autoscale = QtWebKitWidgets.QWebView(self.tab_autoscale)
        self.wv_autoscale.setStyleSheet("")
        self.wv_autoscale.setUrl(QtCore.QUrl("about:blank"))
        self.wv_autoscale.setObjectName("wv_autoscale")
        self.verticalLayout_5.addWidget(self.wv_autoscale)
        self.tab_widget.addTab(self.tab_autoscale, "")
        self.tab_top_fifty = QtWidgets.QWidget()
        self.tab_top_fifty.setObjectName("tab_top_fifty")
        self.verticalLayout_6 = QtWidgets.QVBoxLayout(self.tab_top_fifty)
        self.verticalLayout_6.setObjectName("verticalLayout_6")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.cb_garments = QtWidgets.QComboBox(self.tab_top_fifty)
        self.cb_garments.setMinimumSize(QtCore.QSize(150, 0))
        self.cb_garments.setObjectName("cb_garments")
        self.horizontalLayout.addWidget(self.cb_garments)
        self.lbl_dt_from = QtWidgets.QLabel(self.tab_top_fifty)
        self.lbl_dt_from.setObjectName("lbl_dt_from")
        self.horizontalLayout.addWidget(self.lbl_dt_from)
        self.de_from = QtWidgets.QDateEdit(self.tab_top_fifty)
        self.de_from.setCalendarPopup(True)
        self.de_from.setObjectName("de_from")
        self.horizontalLayout.addWidget(self.de_from)
        self.label = QtWidgets.QLabel(self.tab_top_fifty)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.de_to = QtWidgets.QDateEdit(self.tab_top_fifty)
        self.de_to.setCalendarPopup(True)
        self.de_to.setObjectName("de_to")
        self.horizontalLayout.addWidget(self.de_to)
        self.btn_top_fifty = QtWidgets.QPushButton(self.tab_top_fifty)
        self.btn_top_fifty.setObjectName("btn_top_fifty")
        self.horizontalLayout.addWidget(self.btn_top_fifty)
        spacerItem7 = QtWidgets.QSpacerItem(37, 17, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem7)
        self.verticalLayout_6.addLayout(self.horizontalLayout)
        self.tbl_top_fifty = QtWidgets.QTableWidget(self.tab_top_fifty)
        self.tbl_top_fifty.setObjectName("tbl_top_fifty")
        self.tbl_top_fifty.setColumnCount(5)
        self.tbl_top_fifty.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.tbl_top_fifty.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tbl_top_fifty.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tbl_top_fifty.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tbl_top_fifty.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tbl_top_fifty.setHorizontalHeaderItem(4, item)
        self.verticalLayout_6.addWidget(self.tbl_top_fifty)
        self.tab_widget.addTab(self.tab_top_fifty, "")
        self.verticalLayout.addWidget(self.tab_widget)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 609, 21))
        self.menubar.setObjectName("menubar")
        self.menuMenu = QtWidgets.QMenu(self.menubar)
        self.menuMenu.setObjectName("menuMenu")
        MainWindow.setMenuBar(self.menubar)
        self.actionExit = QtWidgets.QAction(MainWindow)
        self.actionExit.setObjectName("actionExit")
        self.menuMenu.addAction(self.actionExit)
        self.menubar.addAction(self.menuMenu.menuAction())

        self.retranslateUi(MainWindow)
        self.tab_widget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Swiss Marketing Knife"))
        self.label_3.setText(_translate("MainWindow", "Choose File:"))
        self.btn_open_dmc.setText(_translate("MainWindow", "..."))
        self.btn_import.setText(_translate("MainWindow", "Import"))
        self.lbl_match_importing_update.setText(_translate("MainWindow", "Importing:"))
        self.lbl_matching_update.setText(_translate("MainWindow", "Matching DMC ID\'s with Email:"))
        self.lbl_match.setText(_translate("MainWindow", "Matching DMC ID and Email with Campaigns:"))
        self.tab_widget.setTabText(self.tab_widget.indexOf(self.tab_update), _translate("MainWindow", "DMCID\\Campaign Matching"))
        self.label_4.setText(_translate("MainWindow", "Choose File:"))
        self.btn_open_cc.setText(_translate("MainWindow", "..."))
        self.btn_process_cc.setText(_translate("MainWindow", "Process"))
        self.lbl_process_cc.setText(_translate("MainWindow", "Processing File:"))
        self.lbl_upload_cc.setText(_translate("MainWindow", "Uploading File:"))
        self.tab_widget.setTabText(self.tab_widget.indexOf(self.tab_cross), _translate("MainWindow", "Export for Cross Country"))
        self.tab_widget.setTabText(self.tab_widget.indexOf(self.tab_autoscale), _translate("MainWindow", "Auto Scale"))
        self.lbl_dt_from.setText(_translate("MainWindow", "From:"))
        self.label.setText(_translate("MainWindow", "To:"))
        self.btn_top_fifty.setText(_translate("MainWindow", "View"))
        item = self.tbl_top_fifty.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "SKU"))
        item = self.tbl_top_fifty.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "Design"))
        item = self.tbl_top_fifty.horizontalHeaderItem(2)
        item.setText(_translate("MainWindow", "Orders"))
        item = self.tbl_top_fifty.horizontalHeaderItem(3)
        item.setText(_translate("MainWindow", "Pieces"))
        item = self.tbl_top_fifty.horizontalHeaderItem(4)
        item.setText(_translate("MainWindow", "Sales"))
        self.tab_widget.setTabText(self.tab_widget.indexOf(self.tab_top_fifty), _translate("MainWindow", "Top Fifty"))
        self.menuMenu.setTitle(_translate("MainWindow", "Menu"))
        self.actionExit.setText(_translate("MainWindow", "Exit"))

from PyQt5 import QtWebKitWidgets
import resource_rc
