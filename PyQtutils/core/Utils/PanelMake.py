"""
@author by pipixia
UI-py-py逻辑文件
实现界面与逻辑分离
"""
import sys
import re
import os


def changeUiClass(filepath, filename):
    """
    修改ui转成PY文件的类名称
    :param filepath:
    :param filename:
    :return: state true:成功
    """
    linecount = None
    _conf = None
    with open(filepath, "r+", encoding="utf-8") as f:
        conf = f.readlines()
        for i in range(len(conf)):
            a = re.search("^class .+\(object\):", conf[i])
            if a:
                linecount = i
                conf[linecount] = "class {}(object):\n".format(filename)
                _conf = conf
                break

    with open(filepath, "w+", encoding="utf-8") as f:
        if linecount:
            f.writelines(_conf)
            print("修改名称成功")
            return True
        else:
            print("修改名称错误")
            return False


def create_dirandfile(filepath, filename):
    """
    创建view对应的逻辑部分，界面与逻辑分离
    :param filepath:
    :param filename:
    :return:
    """
    _filename = filepath.split("\\")
    # print(_filename)
    for i in range(len(_filename)):
        if _filename[i] == "PyQtutils" and _filename[i + 1] == "view":
            _filename[i + 1] = "src"
            break
    _dirname = _filename[-2]
    _filename[-1] = _filename[-1][2:]
    newfile = "\\".join(_filename)
    newdir = "\\".join(_filename[:-1])
    # print(newfile)
    # print(newdir)
    if not os.path.exists(newdir):
        os.makedirs(newdir)
        print("创建文件夹成功")
    if not os.path.exists(newfile):
        with open(newfile, "x", encoding="utf-8") as f:
            print("创建文件成功")

    with open(newfile, "w", encoding="utf-8") as f:
        f.write("# -----------------以下由脚本生成----------------\n")
        f.write("from PyQt5 import QtWidgets\n")
        f.write("from core.Utils.BasePanel import BasePanel\n")
        f.write("from view.{}.{} import {}\n".format(_dirname, filename, filename))
        f.write("\n\n")
        f.write("class {}(QtWidgets.QWidget, {}, BasePanel):\n".format(_filename[-1][:-3], filename))
        f.write("\n")
        f.write("    def __init__(self, parent=None):\n")
        f.write("        super().__init__(parent)\n")
        f.write("        self.setupUi(self)\n")
        f.write("        self.eventListener()\n")
        f.write("        self.subPanel = None\n")
        f.write("\n")
        f.write("    def eventListener(self):\n")
        f.write("        pass\n")
    print("写入文件成功")


if __name__ == "__main__":
    filepath = sys.argv[1]
    filename = sys.argv[2]
    a = changeUiClass(filepath, filename)
    if a:
        create_dirandfile(filepath, filename)
