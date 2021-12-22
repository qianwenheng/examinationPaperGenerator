from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5 import QtGui

import sys
import os

from utils import utils


class InstructionUI(QtWidgets.QDialog):
    def __init__(self, mainWindow=None):
        super(InstructionUI, self).__init__()

        self.mainWindow = mainWindow
        self.height = 400
        self.width = 800

        edit = QtWidgets.QTextEdit()
        instructionText = '使用说明：\n'
        instructionText += '1.出卷流程：【课程|选择课程】->【课程|统计课程信息】->【试卷|设置试卷信息】->【试卷|出卷】/出卷\n'
        instructionText += '2.默认试卷模板为江苏科技大学试卷模板，可以通过【课程|设置试卷模板】选择其他试卷模板\n'
        instructionText += '3.试题库请按照："课程名|章节|题型|xxx.docx"的格式准备;只支持docx文档，不支持doc文档\n'
        instructionText += '4.可以通过【课程|合并试题库】将所有试题库试题合并在一个word文档中，如果题库题目数量较多，合并过程需要等待较长时间\n'
        instructionText += '5.支持图像、表格和公式。不支持图形和MathType公式。对于图形，请转换成图像；对于MathType公式请在word界面中转换成OfficeMath格式\n'
        instructionText += '6.如果试卷中中文字体不支持，请保证试题库中题目文档中的中文字体设置正确（即不要用英文字体去设置中文字符）\n'
        instructionText += '7.欢迎转发\n'

        edit.setPlainText(instructionText)
        vbox = QtWidgets.QVBoxLayout()

        vbox.addWidget(edit)


        self.setLayout(vbox)

        self.setGeometry(300, 300, self.width, self.height)
        self.move2center()
        self.setWindowTitle('使用说明')

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
    pgUI = InstructionUI()
    sys.exit(app.exec_())
