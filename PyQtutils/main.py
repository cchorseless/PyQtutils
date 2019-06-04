import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from src.MainPanel import MainPanel

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = MainPanel()
    mainWindow.show()
    sys.exit(app.exec_())
