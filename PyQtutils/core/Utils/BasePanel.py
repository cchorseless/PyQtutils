"""
@author pipixia
@info
所有界面的基类
用于界面与逻辑分离

"""
from abc import ABC, abstractmethod


class BasePanel(object):
    """
    使用抽象类会出现元类冲突，所以这里没有用抽象类，
    所有界面的基类
    相当于接口限定子类要实现的方法
    """

    @abstractmethod
    def eventListener(self):
        """
        界面的事件监听，将信号与槽绑定
        :return:
        """
        pass

    @abstractmethod
    def setupUi(self):
        """
        界面初始化
        :return:
        """
        pass

    @property
    def subPanel(self):
        """
        关联的子界面
        :return:
        """
        return self._subPanel

    @subPanel.setter
    def subPanel(self, val):
        """
        定义抽象方法，为属性设置可以改变
        :param val:
        :return:
        """
        self._subPanel = val
