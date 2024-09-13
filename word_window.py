from PyQt5 import QtWidgets
from transform import Ui_Transform
from PyQt5.QtWidgets import QFileDialog
from Word_Function import *
import os


class MyWindow(QtWidgets.QWidget, Ui_Transform):

    def __init__(self):
        super(MyWindow, self).__init__()
        self.setupUi(self)
        self.pushButton_model.clicked.connect(self.read_model)
        self.pushButton_in.clicked.connect(self.read_list_file)
        self.pushButton_save.clicked.connect(self.read_folder)
        self.pushButton_ok.clicked.connect(self.process)

    def read_model(self):
        # 读取模板文件
        try:
            filename, filetype = QFileDialog.getOpenFileNames(self,
                                                              '选取文件',
                                                              os.getcwd(),
                                                              'Text Files(*.docx)')
            print(filename, filetype)
            self.lineEdit_model.setText(filename[0])

        except IndexError:
            fail_result = r'模板文件路径不能为空！'
            self.label_result.setText(fail_result)

    def read_list_file(self):
        # 读取报告清单
        try:

            filename, filetype = QFileDialog.getOpenFileNames(self,
                                                              '选取文件',
                                                              os.getcwd(),
                                                              'Text Files(*.xlsx)')
            print(filename, filetype)
            self.lineEdit_in.setText(filename[0])

        except IndexError:

            fail_result = r'报告清单路径不能为空！'
            self.label_result.setText(fail_result)

    def read_folder(self):
        # 读取输出文件保存路径\
        folder_name = QFileDialog.getExistingDirectory(self,
                                                       '选取文件夹',
                                                       os.getcwd())
        print(folder_name)
        self.lineEdit_save.setText(folder_name)

        if folder_name == '':
            fail_result = r'保存位置不能为空！'
            self.label_result.setText(fail_result)

    def process(self):
        model_path = self.lineEdit_model.text()    # 获取模板路径
        file_path = self.lineEdit_in.text()        # 获取报告清单路径
        folder_path = self.lineEdit_save.text()    # 获取输出文件保存路径
        start = self.lineEdit_start.text()         # 获取起始序号
        end = self.lineEdit_end.text()             # 获取终止序号

        if model_path == '':
            fail_result = r'模板文件路径不能为空！'
            self.label_result.setText(fail_result)
            return

        if file_path == '':
            fail_result = r'报告清单路径不能为空！'
            self.label_result.setText(fail_result)
            return

        if folder_path == '':
            fail_result = r'保存位置不能为空！'
            self.label_result.setText(fail_result)
            return

        try:
            start = int(start)
            end = int(end)
        except ValueError:
            fail_result = r'起始、终止序号不能为空且必须为整数！'
            self.label_result.setText(fail_result)
            return

        try:
            ws = open_excel_worksheet_by_filename(file_path)
            transform(model_path, folder_path, ws, start, end)
        except FileNotFoundError as e:
            self.label_result.setText(str(e))
            return
        except TypeError:
            fail_result = r'请导入正确的模板和报告清单！'
            self.label_result.setText(fail_result)
            return

        success_result = r'转换成功！'
        self.label_result.setText(success_result)


if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ui = MyWindow()
    ui.show()
    sys.exit(app.exec_())
