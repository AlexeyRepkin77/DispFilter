import sys
from PyQt5.QtWidgets import *
from PyQt5 import QtWidgets, QtGui
from DispFilterGUI import Ui_DispFilter
#from ParseExcelOMS import omsFrame


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

        self.ui.Button_Calculate_OMS.clicked.connect(self.btnClicOMC)




    # Функцию открытия диалогового окна выбора файла ОМС ([0]) в конце файла делает так что бы путь выводился корректно
    def getFileNameOMS(self):
        sliceWorkSheetMIS = QFileDialog.getOpenFileName(self, "Выберите файл", ".", "Книга Excel 97-2003(*.xls);" ";Книга Excel(*.xlsx);" ";All Files(*)")[0]
        self.ui.lineEdit_OMS.setText(str(sliceWorkSheetMIS))
        print(sliceWorkSheetMIS)

    def getFileNameMIS(self):
        omsFrame = QFileDialog.getOpenFileName(self, "Выберите файл", ".", "Книга Excel 97-2003(*.xls);" ";Книга Excel(*.xlsx);" ";All Files(*)")[0]
        self.ui.lineEdit_MIS.setText(str(omsFrame))
        print(omsFrame)


    def getDirectory(self):
        generalArchiveSave = QFileDialog.getSaveFileName(self, "Сохраните файл", ".", "Книга Excel(*.xlsx)")[0]
        self.ui.lineEdit_Error.setText(str(generalArchiveSave))

    def btnClicOMC(self):



app = QtWidgets.QApplication(sys.argv)
application = Geneneral_Window()
application.show()

sys.exit(app.exec())