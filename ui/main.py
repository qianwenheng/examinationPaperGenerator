from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5 import QtGui
import PyQt5

import sys
import copy

from ui import subjectUI, modifyQuestionTypeUI, aboutUI, instructionUI
from statistics import subject, questionMerging
from generating import paperGenerating
from utils import utils


class PaperGeneratorUI(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        self.subject = subject.SubjectInfo()
        self.height = 768
        self.width = 1024
        self.setGeometry(0, 0, self.width, self.height)
        self.move2center()
        self.setWindowTitle('江苏科技大学考卷生成器')
        self.setWindowIcon(QtGui.QIcon('JUSTlogo.png'))

        self.subjectName = None
        self.subjectPath = None
        self.isModified = False

        self.subjectUI = subjectUI.SubjectUI(self)

        self.paperTemplate = ''
        self.initUI()
        self.show()

    def move2center(self):
        # 获得窗口
        qr = self.frameGeometry()
        # 获得屏幕中心点
        cp = QtWidgets.QDesktopWidget().availableGeometry().center()
        # 显示到屏幕中心
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def initUI(self):
        self.centerWidget = QtWidgets.QWidget(self)
        self.setCentralWidget(self.centerWidget)

        self.selectSubjectAction = QtWidgets.QAction('选择课程', self)
        self.selectSubjectAction.setToolTip('选择课程试题库所在的文件夹')
        self.selectSubjectAction.triggered.connect(self.showSubjectUI)

        self.showSubjectInfoAction = QtWidgets.QAction('统计课程信息', self)
        self.showSubjectInfoAction.setToolTip('统计课程试题库各章节题型题数')
        self.showSubjectInfoAction.triggered.connect(self.showSubjectInfo)
        self.showSubjectInfoAction.setEnabled(False)

        self.mergeSubjectAction = QtWidgets.QAction('合并试题库', self)
        self.mergeSubjectAction.triggered.connect(self.mergeSubject)
        self.mergeSubjectAction.setEnabled(False)

        self.setPaperInfoAction = QtWidgets.QAction('设置试卷信息', self)
        self.setPaperInfoAction.setToolTip('设置试卷各章节题型题数')
        self.setPaperInfoAction.triggered.connect(self.setPaperInfo)
        self.setPaperInfoAction.setEnabled(False)

        self.setPaperTemplateAction = QtWidgets.QAction('设置试卷模板', self)
        self.setPaperTemplateAction.triggered.connect(self.setPaperTemplate)
        self.setPaperTemplateAction.setEnabled(False)

        self.generatePaperInfoAction = QtWidgets.QAction('出卷', self)
        self.generatePaperInfoAction.triggered.connect(self.paperOkBtnClicked)
        self.generatePaperInfoAction.setEnabled(False)

        self.instructionAction = QtWidgets.QAction('使用说明', self)
        self.instructionAction.triggered.connect(self.showInstruction)
        self.aboutAction = QtWidgets.QAction('关于', self)
        self.aboutAction.triggered.connect(self.showAbout)

        self.statusBar = QtWidgets.QStatusBar()
        self.statusBar.setStyleSheet(
            "QStatusBar{padding-left:8px;background:rgba(255,255,255,128);color:red;font-weight:bold;}")
        self.setStatusBar(self.statusBar)

        menuBar = self.menuBar()
        subjectMenu = menuBar.addMenu('课程')
        paperMenu = menuBar.addMenu('试卷')
        helpMenu = menuBar.addMenu('帮助')

        subjectMenu.addAction(self.selectSubjectAction)
        subjectMenu.addAction(self.showSubjectInfoAction)
        subjectMenu.addAction(self.mergeSubjectAction)

        paperMenu.addAction(self.setPaperInfoAction)
        paperMenu.addAction(self.setPaperTemplateAction)
        paperMenu.addAction(self.generatePaperInfoAction)

        helpMenu.addAction(self.instructionAction)
        helpMenu.addAction(self.aboutAction)

        self.vbox = QtWidgets.QVBoxLayout()
        self.subjectLabel = QtWidgets.QLabel()
        self.vboxPaper = QtWidgets.QVBoxLayout()

        self.vbox.addWidget(self.subjectLabel)
        self.vbox.addStretch(1)
        self.vbox.addLayout(self.vboxPaper)

        self.centerWidget.setLayout(self.vbox)

    def showAbout(self):
        aboutUI.AboutUI().exec()

    def showInstruction(self):
        instructionUI.InstructionUI().exec()

    def showSubjectUI(self):
        self.setPaperInfoAction.setEnabled(False)
        self.setPaperTemplateAction.setEnabled(False)
        self.generatePaperInfoAction.setEnabled(False)
        self.subjectUI.exec()

    def mergeSubject(self):
        self.questionMerging = questionMerging.QuestionMerging()
        mergedDoc = self.questionMerging.merge(self.subjectName, self.subjectPath)

        fileDialog = QtWidgets.QFileDialog()
        filePath, _ = fileDialog.getSaveFileName(caption='保存文件', directory='C:\\', filter='WORD文档(*.docx)')
        mergedDoc.save(filePath)

    def showSubjectInfo(self):

        if not self.isModified:
            self.subject.subjectName = self.subjectName
            self.subject.subjectPath = self.subjectPath
            self.subject.getSubjectStatics()
        else:
            self.isModified = False

        self.inserTable(tableContent='subject')

        self.setPaperInfoAction.setEnabled(True)
        self.generatePaperInfoAction.setEnabled(False)
        self.setPaperTemplateAction.setEnabled(False)
        self.clearLayoutBox(self.vbox)

        self.subjectLabel = QtWidgets.QLabel()
        self.subjectLabel.setText(self.subjectName + ' 课程信息(' + self.subjectPath + ')：')
        self.vbox.addWidget(self.subjectLabel)
        self.vbox.addWidget(self.subjectTable)
        self.vbox.addStretch(1)
        # self.centerWidget.setLayout(self.vbox)

    def setPaperTemplate(self):
        fileDialog = QtWidgets.QFileDialog()
        filePath, _ = fileDialog.getOpenFileName(caption='选择试卷模板文件', directory='C:\\', filter='WORD文档(*.docx)')
        if filePath is not '':
            self.paperTemplate = filePath
            self.paperGenerator.paperTemplate = filePath
            self.paperTemplateLabel.setText('试卷模板已选择为：' + filePath)

    def setPaperInfo(self):
        self.setPaperInfoAction.setEnabled(False)
        self.setPaperTemplateAction.setEnabled(True)
        self.generatePaperInfoAction.setEnabled(True)

        self.paperInfo = paperGenerating.PaperInfo(self.subject)

        self.inserTable(tableContent='paper')

        self.generatePaperInfoAction.setEnabled(True)

        self.paperLabel = QtWidgets.QLabel()
        self.paperLabel.setText('请编辑下面表格中各章节的题型数量：')

        self.modifyQuestionTypeOrderBtn = QtWidgets.QPushButton('调整题序')
        self.setPaperOkBtn = QtWidgets.QPushButton('出卷')
        self.setPaperResetBtn = QtWidgets.QPushButton('清空')
        self.modifyQuestionTypeOrderBtn.clicked.connect(self.modifyBtnClicked)
        self.setPaperOkBtn.clicked.connect(self.paperOkBtnClicked)
        self.setPaperResetBtn.clicked.connect(self.paperResetBtnClicked)

        self.paperTemplateLabel = QtWidgets.QLabel('')
        self.hboxPaper = QtWidgets.QHBoxLayout()
        self.hboxPaper.addWidget(self.paperTemplateLabel)
        self.hboxPaper.addStretch(1)

        self.hboxPaper.addWidget(self.modifyQuestionTypeOrderBtn)
        self.hboxPaper.addWidget(self.setPaperOkBtn)
        self.hboxPaper.addWidget(self.setPaperResetBtn)

        self.vboxPaper = QtWidgets.QVBoxLayout()
        self.vboxPaper.addWidget(self.paperLabel)
        self.vboxPaper.addWidget(self.paperTable)
        self.vboxPaper.addLayout(self.hboxPaper)

        self.vbox.addLayout(self.vboxPaper)
        self.vbox.addStretch(1)

    def modifyBtnClicked(self):
        self.modifyUI = modifyQuestionTypeUI.ModifyQustionTypeUI(self)
        self.modifyUI.exec()

        if self.isModified:
            self.subject.questionTypesOrder = self.modifyUI.rightIndice

            self.clearLayoutBox(self.vbox)
            self.showSubjectInfo()
            self.setPaperInfo()

    def clearLayoutBox(self, L=False):
        if not L:
            L = self.vbox
        if L is not None:
            while L.count():
                item = L.takeAt(0)
                widget = item.widget()

                if widget is not None:
                    widget.deleteLater()
                else:
                    self.clearLayoutBox(item.layout())

    def inserTable(self, tableContent='subject'):
        if tableContent == 'subject':
            self.subjectTable = QtWidgets.QTableWidget()
            self.subjectTable.setGeometry(10, 50, self.width - 20, int(self.height * 0.4))

            tableCols = len(self.subject.questionTypes)
            tableRows = len(self.subject.chapters) + 1

            self.subjectTable.setColumnCount(tableCols)
            self.subjectTable.setRowCount(tableRows)

            self.subjectTable.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

            chapters = copy.deepcopy(self.subject.chapters)
            chapters.append('合计')

            self.questionTypesTemp = []
            for i in self.subject.questionTypesOrder:
                self.questionTypesTemp.append(self.subject.questionTypes[i])
            self.subjectTable.setHorizontalHeaderLabels(self.questionTypesTemp)
            self.subjectTable.setVerticalHeaderLabels(chapters)

            for i in range(tableRows):
                self.subjectTable.setRowHeight(i, self.height / (5 * tableRows))
                for j in range(tableCols):
                    questionTypeIndex = self.subject.questionTypesOrder[j]
                    if i < tableRows - 1:

                        item = QtWidgets.QTableWidgetItem(str(self.subject.chapterQuestionNums[i][questionTypeIndex]))
                    else:
                        item = QtWidgets.QTableWidgetItem(str(self.subject.questionNums[questionTypeIndex]))
                    item.setFlags(QtCore.Qt.ItemFlag(0))
                    item.setFlags(QtCore.Qt.ItemIsEnabled)
                    item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    self.subjectTable.setItem(i, j, item)
        elif tableContent == 'paper':
            self.paperTable = QtWidgets.QTableWidget()
            self.paperTable.setGeometry(10, int(self.height * 0.4) + 50, self.width - 20, int(self.height * 0.4))
            self.paperTable.itemChanged.connect(self.changePaperTableItem)

            tableCols = len(self.subject.questionTypes)
            tableRows = len(self.subject.chapters) + 1

            self.paperTable.setColumnCount(tableCols)
            self.paperTable.setRowCount(tableRows)

            self.paperTable.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

            chapters = copy.deepcopy(self.subject.chapters)
            chapters.append('合计')
            self.paperTable.setHorizontalHeaderLabels(self.questionTypesTemp)
            self.paperTable.setVerticalHeaderLabels(chapters)

            self.isTableInit = True
            for i in range(tableRows):
                self.paperTable.setRowHeight(i, self.height / (5 * tableRows))
                for j in range(tableCols):
                    if i < tableRows - 1:
                        item = QtWidgets.QTableWidgetItem('0')
                    else:
                        item = QtWidgets.QTableWidgetItem('0')
                        item.setFlags(QtCore.Qt.ItemFlag(0))
                        item.setFlags(QtCore.Qt.ItemIsEnabled)
                    item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    self.paperTable.setItem(i, j, item)
            self.isTableInit = False
            self.paperGenerator = paperGenerating.PaperInfo(self.subject)

            if self.paperTemplate is not '':
                self.paperGenerator.paperTemplate = self.paperTemplate

            for i, chapterQuestionNumItem in enumerate(self.subject.chapterQuestionNums):
                self.paperGenerator.chapterQuestionNums.append([])
                for questionNum in chapterQuestionNumItem:
                    self.paperGenerator.chapterQuestionNums[i].append(0)
            self.paperGenerator.convert2QuestionChapterNums()

    def changePaperTableItem(self, item):
        row = item.row()
        col = item.column()
        if not utils.isInt(item.text()):
            print('输入的数据不是正整数')
            item.setText(str(self.paperGenerator.chapterQuestionNums[row][col]))
        else:
            questionNum = int(item.text())
            if row < len(self.subject.chapters):
                questionTypeIndex = self.subject.questionTypesOrder[col]
                if questionNum > self.subject.chapterQuestionNums[row][questionTypeIndex]:
                    print('本章节该题型数量已经超过试题库的题型总数量')
                    item.setText(str(self.paperGenerator.chapterQuestionNums[row][questionTypeIndex]))
                else:
                    if not self.isTableInit:
                        self.paperGenerator.chapterQuestionNums[row][questionTypeIndex] = questionNum
                        self.paperGenerator.convert2QuestionChapterNums()
                        itemTotal = self.paperTable.item(len(self.subject.chapters), col)
                        itemTotal.setText(str(self.paperGenerator.questionNums[questionTypeIndex]))

    def paperOkBtnClicked(self):
        self.paperGenerator.questionTypes = self.subject.questionTypes
        self.paperGenerator.selectQuestionNo()
        self.paperGenerator.generatePaper()
        fileDialog = QtWidgets.QFileDialog()
        filePath, _ = fileDialog.getSaveFileName(caption='保存试卷', directory='C:\\', filter='WORD文档(*.docx)')
        self.paperGenerator.paperDoc.save(filePath)
        filePathList=filePath.split('.')
        filePathTemp=filePath[0:len(filePath)-len(filePathList[-1])-1]
        filePathTemp+='_题目列表.txt'
        with open(filePathTemp, "w") as f:
            for outputPath in self.paperGenerator.outputQuestionPath:
                f.write(outputPath+'\n')

    def paperResetBtnClicked(self):
        tableCols = len(self.subject.questionTypes)
        tableRows = len(self.subject.chapters) + 1
        for i in range(tableRows):
            for j in range(tableCols):
                item = self.paperTable.item(i, j)
                item.setText('0')


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    pgUI = PaperGeneratorUI()
    sys.exit(app.exec_())
