from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5 import QtGui

import sys
import os

from utils import utils


class ModifyQustionTypeUI(QtWidgets.QDialog):
    def __init__(self, mainWindow=None):
        super(ModifyQustionTypeUI, self).__init__()

        self.mainWindow = mainWindow
        self.height = 400
        self.width = 800
        self.rightIndice = []
        if mainWindow is None:
            self.leftIndice = []
        else:
            self.leftIndice = list(range(len(mainWindow.subject.questionTypes)))

        self.initUI()
        self.setGeometry(300, 300, self.width, self.height)
        self.move2center()
        self.setWindowTitle('请调整出题顺序')

        # self.show()

    def move2center(self):
        # 获得窗口
        qr = self.frameGeometry()
        # 获得屏幕中心点
        cp = QtWidgets.QDesktopWidget().availableGeometry().center()
        # 显示到屏幕中心
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def initUI(self):

        self.leftList = QtWidgets.QListWidget()
        self.rightList = QtWidgets.QListWidget()
        for i, questionTypeName in enumerate(self.mainWindow.subject.questionTypes):
            item = QtWidgets.QListWidgetItem(questionTypeName)
            self.leftList.insertItem(i, item)

        self.toRightBtn = QtWidgets.QPushButton('>>')
        self.toRightBtn.clicked.connect(self.toRightBtnClick)
        self.toLeftBtn = QtWidgets.QPushButton('<<')
        self.toLeftBtn.setEnabled(False)
        self.toLeftBtn.clicked.connect(self.toLeftBtnClick)

        leftVbox = QtWidgets.QVBoxLayout()
        leftVbox.addStretch(1)
        leftVbox.addWidget(self.leftList)
        leftVbox.addStretch(1)

        middleVbox = QtWidgets.QVBoxLayout()
        middleVbox.addStretch(1)
        middleVbox.addWidget(self.toRightBtn)
        middleVbox.addStretch(1)
        middleVbox.addWidget(self.toLeftBtn)
        middleVbox.addStretch(1)

        rightVbox = QtWidgets.QVBoxLayout()
        rightVbox.addStretch(1)
        rightVbox.addWidget(self.rightList)
        rightVbox.addStretch(1)

        hbox = QtWidgets.QHBoxLayout()
        hbox.addStretch(1)
        hbox.addLayout(leftVbox)
        hbox.addLayout(middleVbox)
        hbox.addLayout(rightVbox)
        hbox.addStretch(1)

        self.okBtn = QtWidgets.QPushButton('确定')
        self.okBtn.clicked.connect(self.okBtnClick)
        self.okBtn.setEnabled(False)
        cancelBtn = QtWidgets.QPushButton('取消')
        cancelBtn.clicked.connect(self.close)

        hbox2 = QtWidgets.QHBoxLayout()
        hbox2.addStretch(1)
        hbox2.addWidget(self.okBtn)
        hbox2.addWidget(cancelBtn)

        vbox = QtWidgets.QVBoxLayout()
        vbox.addLayout(hbox)
        vbox.addLayout(hbox2)
        self.setLayout(vbox)

    def okBtnClick(self):
        self.mainWindow.isModified = True
        self.close()

    def toRightBtnClick(self):
        n = self.rightList.count()
        leftItem = self.leftList.currentItem()
        if leftItem == None:
            return
        rightItem = QtWidgets.QListWidgetItem(leftItem.text())

        self.rightList.insertItem(n, rightItem)
        self.leftList.removeItemWidget(leftItem)

        leftRow = self.leftList.currentRow()
        self.rightIndice.append(self.leftIndice[leftRow])
        self.leftIndice.pop(leftRow)

        self.leftList.takeItem(leftRow)

        if self.leftList.count() == 0:
            self.toRightBtn.setEnabled(False)
            self.okBtn.setEnabled(True)
        if self.rightList.count() > 0:
            self.toLeftBtn.setEnabled(True)

    def toLeftBtnClick(self):
        n = self.leftList.count()
        rightItem = self.rightList.currentItem()
        if rightItem == None:
            return
        leftItem = QtWidgets.QListWidgetItem(rightItem.text())

        self.leftList.insertItem(n, leftItem)
        self.rightList.removeItemWidget(rightItem)
        rightRow = self.rightList.currentRow()
        self.leftIndice.append(self.rightIndice[rightRow])
        self.rightIndice.pop(rightRow)

        self.rightList.takeItem(rightRow)

        if self.rightList.count() == 0:
            self.toLeftBtn.setEnabled(False)
        if self.leftList.count() > 0:
            self.toRightBtn.setEnabled(True)
            self.okBtn.setEnabled(False)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    pgUI = ModifyQustionTypeUI()
    sys.exit(app.exec_())
