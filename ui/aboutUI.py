from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5 import QtGui

import sys
import os

from utils import utils

class AboutUI(QtWidgets.QDialog):
    def __init__(self, mainWindow=None):
        super(AboutUI, self).__init__()

        self.mainWindow = mainWindow
        self.height = 100
        self.width = 200


        label1=QtWidgets.QLabel('作者：江苏科技大学 计算机学院 钱强')
        label2 = QtWidgets.QLabel('版本：v1.0 欢迎转发')
        vbox=QtWidgets.QVBoxLayout()
        vbox.addStretch(1)
        vbox.addWidget(label1)
        vbox.addWidget(label2)
        vbox.addStretch(1)
        self.setLayout(vbox)

        self.setGeometry(300, 300, self.width, self.height)
        self.move2center()
        self.setWindowTitle('关于')

    def move2center(self):
        # 获得窗口
        qr = self.frameGeometry()
        # 获得屏幕中心点
        cp = QtWidgets.QDesktopWidget().availableGeometry().center()
        # 显示到屏幕中心
        qr.moveCenter(cp)
        self.move(qr.topLeft())












if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    pgUI = AboutUI()
    sys.exit(app.exec_())
