"""
@author pipixia
@info
信号管理
用于信号与槽的分离
"""
from PyQt5.QtCore import QObject, pyqtSignal
from core.Utils.SingletonMake import *


@singleton
class Worker(QObject):
    parse_triggered = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
