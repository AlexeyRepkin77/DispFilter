import sys
from PyQt5.QtWidgets import *
from PyQt5 import QtWidgets, QtGui
from DispFilterGUI import Ui_DispFilter



global sliceWorkSheetMIS

class Geneneral_Window(QtWidgets.QMainWindow):
    def __init__(self):
        super(Geneneral_Window, self).__init__()
        self.ui = Ui_DispFilter()
        self.ui.setupUi(self)
        # Ссылка на функцию открытия диалогового окна для выбора файла ОМС
        self.ui.TButton_OMS.clicked.connect(self.getFileNameOMS)
        # Ссылка на функцию открытия диалогового окна для выбора файла МИС
        self.ui.TButton_MIS.clicked.connect(self.getFileNameMIS)
        # Ссылка на функцию открытия диалогового окна для сохранения файла ошибок
        self.ui.TButton_Error.clicked.connect(self.getDirectory)
        # Ссылка на функцию нажатия на кнопки фильрации архива МИС
        self.ui.Button_Calculate_OMS.clicked.connect(self.btnClicOMC)
        # Ссылка на функцию нажатия на кнопки фильрации архива МИС
        self.ui.Button_Calculate_MIS.clicked.connect(self.btnClicMIS)
        # Ссылка на функцию нажатия на кнопки сравнения двух файлов
        self.ui.Button_Result.clicked.connect(self.btnResult)
        # Ссылка на функцию открытия QAbstractTableModel для быстрого просмотра ошибок
        self.ui.Button_Open_Result.clicked.connect(self.btnOpenResult)
        # Ссылка на функцию сохранения файла ошибок в формате Excel
        self.ui.Button_Save_Result.clicked.connect(self.btnSaveResult)





    # Функцию открытия диалогового окна выбора файла ОМС ([0]) в конце файла делает так что бы путь выводился корректно в поле QLine_Edit
    def getFileNameOMS(self):
        sliceWorkSheetMIS = QFileDialog.getOpenFileName(self, "Выберите файл", ".", "Книга Excel 97-2003(*.xls);" 
                                                                                    "; Книга Excel(*.xlsx);" 
                                                                                    ";All Files(*)")[0]
        self.ui.lineEdit_OMS.setText(str(sliceWorkSheetMIS))

    # Функцию открытия диалогового окна выбора файла МИС ([0]) в конце файла делает так что бы путь выводился корректно в поле QLine_Edit
    def getFileNameMIS(self):
        omsFrame = QFileDialog.getOpenFileName(self, "Выберите файл", ".", "Книга Excel 97-2003(*.xls);" 
                                                                           "; Книга Excel(*.xlsx);" 
                                                                           ";All Files(*)")[0]
        self.ui.lineEdit_MIS.setText(str(omsFrame))

    # Функцию открытия диалогового окна для сохранения файла ошибок ([0]) в конце файла делает так что бы путь выводился корректно в поле QLine_Edit
    def getDirectory(self):
        generalArchiveSave = QFileDialog.getSaveFileName(self, "Сохраните файл", ".", "Книга Excel(*.xlsx)")[0]
        self.ui.lineEdit_Error.setText(str(generalArchiveSave))


    def btnClicOMC(self):
        pass

    def btnClicMIS(self):
        pass

    def btnResult(self):
        pass

    def btnOpenResult(self):
        pass

    def btnSaveResult(self):
        pass




app = QtWidgets.QApplication(sys.argv)
application = Geneneral_Window()
application.show()

sys.exit(app.exec())