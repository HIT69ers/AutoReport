# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'transform.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Transform(object):
    def setupUi(self, Transform):
        Transform.setObjectName("Transform")
        Transform.resize(640, 259)
        self.label_model = QtWidgets.QLabel(Transform)
        self.label_model.setGeometry(QtCore.QRect(30, 50, 131, 21))
        self.label_model.setObjectName("label_model")
        self.label_in = QtWidgets.QLabel(Transform)
        self.label_in.setGeometry(QtCore.QRect(30, 90, 121, 21))
        self.label_in.setObjectName("label_in")
        self.lineEdit_model = QtWidgets.QLineEdit(Transform)
        self.lineEdit_model.setGeometry(QtCore.QRect(170, 50, 321, 20))
        self.lineEdit_model.setObjectName("lineEdit_model")
        self.lineEdit_in = QtWidgets.QLineEdit(Transform)
        self.lineEdit_in.setGeometry(QtCore.QRect(170, 90, 321, 20))
        self.lineEdit_in.setObjectName("lineEdit_in")
        self.pushButton_model = QtWidgets.QPushButton(Transform)
        self.pushButton_model.setGeometry(QtCore.QRect(510, 50, 101, 23))
        self.pushButton_model.setObjectName("pushButton_model")
        self.pushButton_in = QtWidgets.QPushButton(Transform)
        self.pushButton_in.setGeometry(QtCore.QRect(510, 90, 101, 23))
        self.pushButton_in.setObjectName("pushButton_in")
        self.label_start = QtWidgets.QLabel(Transform)
        self.label_start.setGeometry(QtCore.QRect(30, 130, 81, 21))
        self.label_start.setObjectName("label_start")
        self.lineEdit_start = QtWidgets.QLineEdit(Transform)
        self.lineEdit_start.setGeometry(QtCore.QRect(170, 130, 113, 20))
        self.lineEdit_start.setObjectName("lineEdit_start")
        self.lineEdit_end = QtWidgets.QLineEdit(Transform)
        self.lineEdit_end.setGeometry(QtCore.QRect(402, 130, 91, 20))
        self.lineEdit_end.setObjectName("lineEdit_end")
        self.label_end = QtWidgets.QLabel(Transform)
        self.label_end.setGeometry(QtCore.QRect(310, 130, 81, 21))
        self.label_end.setObjectName("label_end")
        self.label_result = QtWidgets.QLabel(Transform)
        self.label_result.setGeometry(QtCore.QRect(30, 210, 321, 21))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setBold(True)
        font.setWeight(75)
        self.label_result.setFont(font)
        self.label_result.setText("")
        self.label_result.setObjectName("label_result")
        self.pushButton_ok = QtWidgets.QPushButton(Transform)
        self.pushButton_ok.setGeometry(QtCore.QRect(364, 210, 101, 23))
        self.pushButton_ok.setObjectName("pushButton_ok")
        self.pushButton_cancel = QtWidgets.QPushButton(Transform)
        self.pushButton_cancel.setGeometry(QtCore.QRect(494, 210, 101, 23))
        self.pushButton_cancel.setObjectName("pushButton_cancel")
        self.label_save = QtWidgets.QLabel(Transform)
        self.label_save.setGeometry(QtCore.QRect(30, 170, 131, 21))
        self.label_save.setObjectName("label_save")
        self.lineEdit_save = QtWidgets.QLineEdit(Transform)
        self.lineEdit_save.setGeometry(QtCore.QRect(170, 170, 321, 20))
        self.lineEdit_save.setObjectName("lineEdit_save")
        self.pushButton_save = QtWidgets.QPushButton(Transform)
        self.pushButton_save.setGeometry(QtCore.QRect(510, 170, 101, 23))
        self.pushButton_save.setObjectName("pushButton_save")

        self.retranslateUi(Transform)
        self.pushButton_cancel.clicked.connect(Transform.close)
        QtCore.QMetaObject.connectSlotsByName(Transform)

    def retranslateUi(self, Transform):
        _translate = QtCore.QCoreApplication.translate
        Transform.setWindowTitle(_translate("Transform", "Transform"))
        self.label_model.setText(_translate("Transform", "导入模板文件："))
        self.label_in.setText(_translate("Transform", "导入报告清单："))
        self.pushButton_model.setText(_translate("Transform", "选择Word"))
        self.pushButton_in.setText(_translate("Transform", "选择Excel"))
        self.label_start.setText(_translate("Transform", "起始序号："))
        self.label_end.setText(_translate("Transform", "终止序号："))
        self.pushButton_ok.setText(_translate("Transform", "转 换"))
        self.pushButton_cancel.setText(_translate("Transform", "取 消"))
        self.label_save.setText(_translate("Transform", "选择保存位置："))
        self.pushButton_save.setText(_translate("Transform", "选择文件夹"))

