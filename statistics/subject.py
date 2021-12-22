import os


class SubjectInfo:
    def __init__(self, subjectName=None, subjectPath=None):
        self.subjectName = subjectName
        self.subjectPath = subjectPath

        self.chapters = []  # 课程所含的所有章节
        self.questionTypes = []  # 课程包含的所有题型
        self.chapterQuestionNums = []  # [ listCh1, listCh2,...] 其中listChN也是list，里面本章各种题型对应的数量
        self.questionChpaterNums = []
        self.questionNums = []
        self.questionTypesOrder=[]

    def getQuestionNums(self):
        if self.chapters == []:
            self.getSubjectStatics()
        questionNums = [0 for x in self.questionTypes]
        for questionChapterNums in self.chapterQuestionNums:
            questionNums = [x + y for x, y in zip(questionNums, questionChapterNums)]
        self.questionNums = questionNums
        return questionNums

    def getQuestionTypes(self):
        if self.subjectPath is None:
            self.subjectPath = os.path.join('..\\subjects', self.subjectName)
        questionTypes = []
        if len(self.chapters) == 0:
            self.chapters = os.listdir(self.subjectPath)
        for chapterItem in self.chapters:
            chapterPath = os.path.join(self.subjectPath, chapterItem)
            questionChapterTypes = os.listdir(chapterPath)
            questionTypes.extend(questionChapterTypes)

        questionTypes = list(set(questionTypes))
        self.questionTypes = questionTypes
        self.questionTypesOrder=list(range(len(self.questionTypes)))
        return questionTypes

    def getSubjectStatics(self):
        if self.subjectPath is None:
            self.subjectPath = os.path.join('..\\subjects', self.subjectName)

        self.chapters = os.listdir(self.subjectPath)

        self.getQuestionTypes()

        chapterQuestionNums = []
        for chapterItem in self.chapters:
            chapterQusetionNum = []
            chapterPath = os.path.join(self.subjectPath, chapterItem)

            for questionTypeItem in self.questionTypes:
                questionTypePath = os.path.join(chapterPath, questionTypeItem)
                if os.path.exists(questionTypePath):
                    questions = os.listdir(questionTypePath)
                    chapterQusetionNum.append(len(questions))
                else:
                    chapterQusetionNum.append(0)
            chapterQuestionNums.append(chapterQusetionNum)

        self.chapterQuestionNums = chapterQuestionNums
        self.convert2QuestionChapterNums()

        self.questionNums=[]
        for questionChapterNum in self.questionChpaterNums:
            self.questionNums.append(sum(questionChapterNum))


    def convert2QuestionChapterNums(self):
        questionChapterNums = []

        for chapterQuestionNum in self.chapterQuestionNums:
            index = 0
            for questionTypeNum in chapterQuestionNum:
                if index + 1 > len(questionChapterNums):
                    questionChapterNums.append([])
                questionChapterNums[index].append(questionTypeNum)
                index = index + 1
        self.questionChpaterNums = questionChapterNums


if __name__ == '__main__':
    subjectInfo = SubjectInfo('信息论')
    subjectInfo.getSubjectStatics()
    subjectInfo.getQuestionNums()
    print(subjectInfo)
