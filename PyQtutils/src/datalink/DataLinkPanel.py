# -----------------以下由脚本生成----------------
from PyQt5 import QtWidgets
from core.Utils.BasePanel import BasePanel
from view.datalink.UiDataLinkPanel import UiDataLinkPanel
from core.Manage.SignalManage import *


class DataLinkPanel(QtWidgets.QWidget, UiDataLinkPanel, BasePanel):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.eventListener()
        self.subPanel = None

    def eventListener(self):
        self.pushButton.clicked.connect(self.checkresult)

    def checkresult(self):
        tmp = self.tableWidget.selectedItems()
        if tmp is not None:
            _list = []
            for i in tmp:
                _list.append(i.row())
            Worker().parse_triggered.emit(_list)
            self.close()
        else:
            warring_msg = QtWidgets.QMessageBox.information(self, "警告", "请选择挂靠数据表", QtWidgets.QMessageBox.Yes)

    def loaddata(self, sheet_ok):
        self.tableWidget.setRowCount(len(sheet_ok.keys()))
        self.tableWidget.setColumnCount(1)
        self.tableWidget.setHorizontalHeaderLabels(['可挂靠数据表'])
        self.tableWidget.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)  # 整行选中
        self.tableWidget.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)  # 设置表格不可编辑
        for i in range(len(sheet_ok.keys())):
            self.tableWidget.setItem(i, 0, QtWidgets.QTableWidgetItem(sheet_ok[i]))
