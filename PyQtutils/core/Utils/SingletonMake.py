"""
@author pipixia
@info
单例模式
用于对象与对象之间的交互，保证唯一性
当作类装饰器，用法 @singleton
"""


def singleton(cls):
    """
    单例模式 by pipixia
    :param cls:
    :return:
    """
    instances = {}

    def getinstance(*args, **kwargs):
        if cls not in instances:
            instances[cls] = cls(*args, **kwargs)
        return instances[cls]

    return getinstance
