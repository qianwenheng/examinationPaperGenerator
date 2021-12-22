import os


def moveNonDocxFile(questionFiles):
    i = 0
    while i < len(questionFiles):
        fileName, extName = os.path.splitext(questionFiles[i])
        if not extName == '.docx':
            questionFiles.remove(questionFiles[i])
            i = i - 1
        i = i + 1
    return questionFiles


def checkSubjectPath(subjectPath):
    chaptersTemp = os.listdir(subjectPath)
    chapters=[]
    chapterFiles=[]
    for chapterTempItem in chaptersTemp:
        chapterTempItemPath=os.path.join(subjectPath,chapterTempItem)
        if os.path.isdir(chapterTempItemPath):
            chapters.append(chapterTempItem)
        elif os.path.isfile(chapterTempItemPath):
            chapterFiles.append(chapterTempItem)

    if len(chapters) == 0:
        print("试题库目录下没有章节")
        return False
    elif len(chapterFiles)>0:
        print("试题库目录下不应有文件")
        return False
    else:
        questionNum = 0
        for chapter in chapters:
            chapterPath = os.path.join(subjectPath, chapter)
            questionTypesTemp = os.listdir(chapterPath)
            questionTypes = []
            questionTypesFiles=[]

            for questionTypeTempItem in questionTypesTemp:
                questionTypeTempItemPath=os.path.join(chapterPath,questionTypeTempItem)
                if os.path.isdir(questionTypeTempItemPath):
                    questionTypes.append(questionTypeTempItem)
                elif os.path.isfile(questionTypeTempItemPath):
                    questionTypesFiles.append(questionTypeTempItem)

            if len(questionTypesFiles)>0:
                print(chapterPath+"：该章节目录下不应有文件")
                return False

            for questionType in questionTypes:
                questionTypePath = os.path.join(chapterPath, questionType)
                questionsTemp = os.listdir(questionTypePath)
                questions=[]
                questionsDir=[]
                for questionTempItem in questionsTemp:
                    questionTempItemPath=os.path.join(questionTypePath,questionTempItem)
                    if os.path.isdir(questionTempItemPath):
                        questionsDir.append(questionTempItem)
                    elif os.path.isfile(questionTempItemPath):
                        questions.append(questionTempItem)

                if len(questionsDir) > 0:
                    print(questionTypePath + ': 该题型目录下不应再有子目录')
                    return False

                questions = moveNonDocxFile(questions)
                questionNum += len(questions)
        if questionNum==0:
            print('试题库目录下总题数为0')
            return False
    return True


def isInt(s):
    try:
        temp=int(s)
        if temp<0:
            return False
        else:
            return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False
