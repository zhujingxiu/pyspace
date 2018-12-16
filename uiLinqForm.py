# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'uiLinqForm.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!
import os
from PyQt5 import QtCore, QtGui, QtWidgets
from readXLS import read4Linq


class Ui_LinqForm(object):
    def __init__(self):
        self.interval = 20

    def setupUi(self, LinqForm):
        self.selectedDir = False
        self.cwd = os.getcwd()  # 获取当前程序文件位置
        self.output = self.desktopDir()
        LinqForm.setObjectName("LinqForm")
        LinqForm.resize(550, 450)

        self.centralwidget = QtWidgets.QWidget(LinqForm)
        self.centralwidget.setObjectName("centralwidget")

        self.widget = QtWidgets.QWidget(self.centralwidget)
        self.widget.setGeometry(QtCore.QRect(20, 10, 491, 431))
        self.widget.setObjectName("widget")
        # 选择文件夹
        self.dirChooseButton = QtWidgets.QPushButton(self.widget)
        self.dirChooseButton.setGeometry(QtCore.QRect(80, 20, 160, 30))
        self.dirChooseButton.setObjectName("dirChooseButton")
        self.choosedFiles = QtWidgets.QLabel(self.widget)
        self.choosedFiles.setGeometry(QtCore.QRect(260, 20, 160, 20))
        self.choosedFiles.setText("")
        self.choosedFiles.setObjectName("choosedFiles")
        # 选定的文件夹路径
        self.dirLabelInput = QtWidgets.QLabel(self.widget)
        self.dirLabelInput.setGeometry(QtCore.QRect(80, 50, 320, 20))
        self.dirLabelInput.setText("")
        self.dirLabelInput.setObjectName("dirLabelInput")
        # 选定文件夹下的xls文件
        self.scrollArea = QtWidgets.QScrollArea(self.widget)
        self.scrollArea.setGeometry(80, 90, 320, 120)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setObjectName("scrollArea")
        self.scrollArea.setStyleSheet("background-color:white;")
        self.scrollBar = self.scrollArea.verticalScrollBar()
        self.scrollBar.setStyleSheet("background-color:gray;")

        self.scrollContents = QtWidgets.QWidget()
        self.scrollContents.setGeometry(80, 90, 320, 100)
        self.scrollContents.setMinimumSize(280, 1000)

        # 设置sheet大纲关键字
        self.labelKeyword = QtWidgets.QLabel(self.widget)
        self.labelKeyword.setGeometry(QtCore.QRect(80, 220, 80, 20))
        self.labelKeyword.setObjectName("labelKeyword")
        self.sheetKeyword = QtWidgets.QLineEdit(self.widget)
        self.sheetKeyword.setGeometry(QtCore.QRect(180, 220, 140, 20))
        self.sheetKeyword.setObjectName("sheetKeyword")
        # 设置sheet大纲关键字
        self.dirOutputButton = QtWidgets.QPushButton(self.widget)
        self.dirOutputButton.setGeometry(QtCore.QRect(80, 260, 320, 30))
        self.dirOutputButton.setObjectName("dirOutputButton")
        # 选定输出的文件路径
        self.dirLabelOutput = QtWidgets.QLabel(self.widget)
        self.dirLabelOutput.setGeometry(QtCore.QRect(80, 300, 320, 20))
        self.dirLabelOutput.setText("默认:%s" % self.output)
        self.dirLabelOutput.setObjectName("dirLabelOutput")
        # 是否检测文件
        self.checkFile = QtWidgets.QCheckBox(self.widget)
        self.checkFile.setGeometry(QtCore.QRect(80, 350, 111, 20))
        self.checkFile.setObjectName("checkFile")
        self.checkFile.setChecked(True)
        # 生成统计
        self.runButton = QtWidgets.QPushButton(self.widget)
        self.runButton.setGeometry(QtCore.QRect(80, 370, 320, 40))
        self.runButton.setObjectName("runButton")

        self.statusbar = QtWidgets.QStatusBar(LinqForm)
        self.statusbar.setObjectName("statusbar")

        LinqForm.setStatusBar(self.statusbar)

        self.widget.raise_()
        LinqForm.setCentralWidget(self.centralwidget)

        # 设置信号
        self.dirChooseButton.clicked.connect(self.eventInputChoosedDir)
        self.dirOutputButton.clicked.connect(self.eventOutputChoosedDir)
        self.runButton.clicked.connect(self.eventRunning)

        self.retranslateUi(LinqForm)
        QtCore.QMetaObject.connectSlotsByName(LinqForm)


    def retranslateUi(self, LinqForm):
        _translate = QtCore.QCoreApplication.translate
        LinqForm.setWindowTitle(_translate("LinqForm", "林琼有点皮"))
        self.dirChooseButton.setText(_translate("LinqForm", "选择文件夹"))
        self.labelKeyword.setText(_translate("LinqForm", "Sheet关键字"))
        self.dirOutputButton.setText(_translate("LinqForm", "选择输出文件路径"))
        self.checkFile.setText(_translate("LinqForm", "检测文件格式"))
        self.runButton.setText(_translate("LinqForm", "生成统计xls"))
        self.statusbar.showMessage(_translate("LinqForm", "林琼私人专属 版权所有，违者必究"))


    def eventRunning(self):
        if not self.selectedDir:
            QtWidgets.QMessageBox.warning(self.centralwidget, "警告", "请选择目标文件夹", QtWidgets.QMessageBox.Cancel)
            return False
        try:
            linq = read4Linq(self.selectedDir, self.sheetKeyword.text(), self.output)
            checkingFile = self.checkFile.isChecked()
            ignore = True
            if checkingFile:
                try:
                    result = linq.checkingDir()
                except Exception as e:
                    QtWidgets.QMessageBox.warning(self.centralwidget, "警告", e, QtWidgets.QMessageBox.Cancel)
                    return False
                if result:
                    counter = 0
                    msg = []
                    for i, item in enumerate(result):
                        # print(item)
                        _text = "%d.%s----%s" % (counter + 1, item.get('error'), item.get('store'))
                        # self.labelChoosedFile = QtWidgets.QLabel(self.scrollContents)
                        # self.labelChoosedFile.setText()
                        # self.labelChoosedFile.move(10, self.interval * (i+ len(self.scrollContents.children())))
                        msg.append(_text)
                        counter += 1
                    self.scrollArea.setWidget(self.scrollContents)
                    button = QtWidgets.QMessageBox.question(self.centralwidget, "格式确认", "检测到文件中有非正常格式，是否跳过直接运行!\n%s"%("\n".join(msg)),
                                                            QtWidgets.QMessageBox.Ok | QtWidgets.QMessageBox.Cancel, QtWidgets.QMessageBox.Ok)
                    if button == QtWidgets.QMessageBox.Cancel:
                        ignore = False
            if ignore:
                detail_rows = linq.readDir()
                xls_file = linq.genSheets(detail_rows)
                if xls_file:
                    QtWidgets.QMessageBox.information (self.centralwidget, "成功", "文件路径:%s"%(xls_file), QtWidgets.QMessageBox.Cancel)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self.centralwidget, "警告", e, QtWidgets.QMessageBox.Cancel)
            return False

    def eventInputChoosedDir(self):
        try:
            # 起始路径
            dir_choosed_input = QtWidgets.QFileDialog.getExistingDirectory(self.centralwidget, "选取文件夹", self.cwd)
            if dir_choosed_input and os.path.isdir(dir_choosed_input):
                counter = 0
                for i, item in enumerate(os.listdir(dir_choosed_input)):
                    filename, fileext = os.path.splitext(item)
                    self.labelChoosedFile = QtWidgets.QLabel(self.scrollContents)
                    self.labelChoosedFile.setText("%d.%s" % (counter + 1, item))
                    self.labelChoosedFile.move(10, self.interval * i)
                    if filename.find('xls'):
                        counter += 1
                if counter:
                    self.dirLabelInput.setText(dir_choosed_input[-40:])
                    self.selectedDir = self.cwd = dir_choosed_input
                    self.scrollArea.setWidget(self.scrollContents)
                    self.scrollContents.setMinimumSize(300, counter * self.interval)
                    self.choosedFiles.setText("共有%s个xls文件" % counter)
                    self.sheetKeyword.setText(os.path.basename(dir_choosed_input))
                else:
                    QtWidgets.QMessageBox.warning(self.centralwidget, "警告", "该目录下无xls文件", QtWidgets.QMessageBox.Cancel)
                    return False
        except Exception as e:
            QtWidgets.QMessageBox.critical(self.centralwidget, "出错了", e)
            return False

    def eventOutputChoosedDir(self):
        try:
            # 起始路径
            dir_choosed_output = QtWidgets.QFileDialog.getExistingDirectory(self.centralwidget, "选取文件夹", self.desktopDir())
            if dir_choosed_output and os.path.isdir(dir_choosed_output):
                self.dirLabelOutput.setText(dir_choosed_output)
                self.output = dir_choosed_output
        except Exception as e:
            QtWidgets.QMessageBox.critical(self.centralwidget, "出错了", e)
            return False

    def desktopDir(self):
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
        dirPath = winreg.QueryValueEx(key, "Desktop")[0]
        return dirPath if os.path.exists(dirPath) else self.cwd
