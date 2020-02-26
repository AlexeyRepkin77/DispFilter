import sys
from PyQt5 import QtWidgets, QtGui
from DispFilterGUI import Ui_DispFilter

class dispProgram(QtWidgets.QMainWindow):
    def __init__(self):
        super(dispProgram, self).__init__()
        self.ui = Ui_DispFilter()
        self.ui.setupUi(self)




app = QtWidgets.QApplication([])
application = dispProgram()
application.show()

sys.exit(app.exec())