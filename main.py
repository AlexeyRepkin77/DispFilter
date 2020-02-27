import sys
from PyQt5.QtWidgets import *
from PyQt5 import QtWidgets, QtGui
from DispFilterGUI import Ui_DispFilter




class Geneneral_Window(QtWidgets.QMainWindow):
    def __init__(self):
        super(Geneneral_Window, self).__init__()
        self.ui = Ui_DispFilter()
        self.ui.setupUi(self)
        # Ссылка на функцию открытия диалогового окна
        self.ui.TButton_OMS.clicked.connect(self.getFileName)

        self.ui.TButton_MIS.clicked.connect(self.getFileName)

        self.ui.TButton_Error.clicked.connect(self.getDirectory)


        # Функцию открытия диалогового окна
    def getFileName(self):
        d = QFileDialog.getOpenFileName(self, "Выберите файл", ".", "Книга Excel 97-2003(*.xls);"
                                                                          ";Книга Excel(*.xlsx);"
                                                                          ";All Files(*)")
        print(d)
        self.ui.lineEdit_OMS.setText(str(d))

    def getDirectory(self):
        b = QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")


app = QtWidgets.QApplication(sys.argv)
application = Geneneral_Window()
application.show()

sys.exit(app.exec())