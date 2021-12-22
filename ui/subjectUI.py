from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5 import QtGui

import sys
import os

from utils import utils

class SubjectUI(QtWidgets.QDialog):
    def __init__(self, mainWindow=None):
        super(SubjectUI, self).__init__()

        self.mainWindow = mainWindow
        self.height = 800
        self.width = 400
        self.treeViewSelectIndex = None
        self.setGeometry(300, 300, self.width, self.height)
        self.move2center()
        self.setWindowTitle('请选择课程试题库所在目录')
        self.showSubjectView()
        # self.show()

    def move2center(self):
        # 获得窗口
        qr = self.frameGeometry()
        # 获得屏幕中心点
        cp = QtWidgets.QDesktopWidget().availableGeometry().center()
        # 显示到屏幕中心
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def showSubjectView(self):

        self.model = QtWidgets.QDirModel()
        self.model.setFilter(QtCore.QDir.AllDirs | QtCore.QDir.Dirs | QtCore.QDir.Drives | QtCore.QDir.NoDotAndDotDot)
        self.treeView = QtWidgets.QTreeView(self)
        self.treeView.setModel(self.model)
        self.treeView.clicked.connect(self.tree_cilcked)
        # self.treeView.doubleClicked.connect(self.tree_cilcked)
        self.treeView.setGeometry(0, 0, self.width, self.height - 40)
        for i in [1, 2, 3]:
            self.treeView.setColumnHidden(i, True)
        self.treeView.setHeaderHidden(True)
        self.treeView.setIndentation(10)

        hbox = QtWidgets.QHBoxLayout()
        self.label = QtWidgets.QLabel()
        hbox.addWidget(self.label)
        hbox.addStretch(1)
        self.okButton = QtWidgets.QPushButton("确定")
        self.okButton.setEnabled(False)
        self.okButton.clicked.connect(self.okButtonClicked)
        self.cancelButton = QtWidgets.QPushButton("取消")
        self.cancelButton.clicked.connect(self.close)
        hbox.addWidget(self.okButton)
        hbox.addWidget(self.cancelButton)

        vbox = QtWidgets.QVBoxLayout()
        vbox.addWidget(self.treeView)
        vbox.addLayout(hbox)
        self.setLayout(vbox)



    def okButtonClicked(self):
        subjectPath = self.model.filePath(self.treeViewSelectIndex)
        subjectName = self.model.fileName(self.treeViewSelectIndex)
        if os.path.isfile(subjectPath):
            self.label.setText('请选择文件夹')
            # self.label.repaint()
            print('请选择文件夹')
        else:
            isSubjectPath=utils.checkSubjectPath(subjectPath)
            if isSubjectPath:
                self.mainWindow.subjectPath = subjectPath
                self.mainWindow.subjectName = subjectName

                self.mainWindow.clearLayoutBox(self.mainWindow.vbox)
                self.mainWindow.subjectLabel = QtWidgets.QLabel()

                self.mainWindow.subjectLabel.setText('已经选择课程为：'+subjectName+'。      位于'+subjectPath)
                self.mainWindow.statusBar.showMessage('已经选择课程为：'+subjectName+'。      位于'+subjectPath)

                self.mainWindow.vbox.addWidget(self.mainWindow.subjectLabel)
                self.mainWindow.vbox.addStretch(1)

                self.mainWindow.isSelectSubject=True
                self.mainWindow.showSubjectInfoAction.setEnabled(True)
                self.mainWindow.mergeSubjectAction.setEnabled(True)
                #self.label.setText(subjectPath)
                self.close()
            else:
                self.label.setStyleSheet("font-size:14px;color:red")
                self.label.setText('请选择正确的课程试题库目录')




    def tree_cilcked(self, Qmodelidx):
        self.treeViewSelectIndex = Qmodelidx
        self.okButton.setEnabled(True)
        # print(self.model.filePath(Qmodelidx))
        # print(self.model.fileName(Qmodelidx))
        # print(self.model.fileInfo(Qmodelidx))


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    pgUI = SubjectUI()
    sys.exit(app.exec_())
