import sys
import PyPDF2
import docx
import Levenshtein as lev
import string
from nltk import text
from nltk.corpus import wordnet
from PyQt5 import QtGui
from PyQt5.QtCore import QFileInfo
from PyQt5.QtGui import QStaticText, QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPlainTextEdit, QTextEdit, QFileDialog, QPushButton, QComboBox, QGridLayout


class MainWindow(QMainWindow):
    globalFileName = None
    def __init__(self):
        #set up actual window
        super().__init__()
        self.setWindowTitle("ResuMate")
        self.setStyleSheet("QMainWindow {background: 'white';}")
        self.icon = QIcon('/Users/clairethompson/Desktop/EngHacks2021/Logo.png')
        self.setWindowIcon(self.icon)
        self.height = 800
        self.width = 800
        self.left = 10
        self.top = 10
        self.setGeometry(self.left, self.top,self.width, self.height)
        

        #set up widgets
        self.textbox = QTextEdit(self)
        self.textbox.move(350, 80)
        self.textbox.resize(400,220)
        self.textbox.verticalScrollBar().minimum()
        self.resultsbox = QTextEdit(self)
        self.button1 = QPushButton("Choose File", self)
        self.button1.clicked.connect(self.getOpenFileName)
        self.button2 = QPushButton("Load New Point", self)
        self.button2.clicked.connect(self.loadText)
        self.button3 = QPushButton("Generate Resume", self)
        self.button3.clicked.connect(self.backEnd)
        self.button3.setEnabled(False)
        self.button4 = QPushButton("Clear", self)
        self.button4.hide()
        self.button4.clicked.connect(self.clear)
        self.label1 = QLabel("Load Job Posting",self)
        self.label2 = QLabel("Add a new qualification point",self)
        self.label3 = QLabel("Loaded Job Posting:",self)
        self.label4 = QLabel(self)
        self.label5=QLabel(self)
        self.label6 = QLabel("Select prefered number of qualification points", self)
        self.comboBox = QComboBox(self)
        self.comboBox.addItems(['1','2','3','4','5'])
        self.comboBox.currentIndexChanged.connect(self.selectionChange)


        #arrange widgets
        self.button1.move(30,80)
        self.button2.move(349,320)
        self.button2.adjustSize() 
        self.button3.move(30, 325) 
        self.button3.adjustSize()
        self.label1.move(30,40)
        self.label1.adjustSize()
        self.label2.move(350, 40)
        self.label2.adjustSize()
        self.label3.move(30,140)
        self.label3.adjustSize()
        self.label6.move(30, 250)
        self.label6.adjustSize()
        self.comboBox.move(25, 270)

        #Style Widgets
        self.label1.setStyleSheet('color: #229396;')
        self.label2.setStyleSheet('color: #229396;')
        self.label3.setStyleSheet('color: #229396;')
        self.label4.setStyleSheet('color: #229396;')
        self.label5.setStyleSheet('color: #229396;')
        self.label6.setStyleSheet('color: #229396;')
        self.button1.setStyleSheet("background-color: #adb7b7; color: #229396;")
        self.button2.setStyleSheet("background-color: #adb7b7; color: #229396;")
        self.button3.setStyleSheet("background-color: #adb7b7; color: #229396;")

    #Gets job description pdf and displays file name once loaded
    def getOpenFileName(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,'Open file','PDF files (*.pdf)', options = options)
        if fileName:
            filename2 = QFileInfo(fileName).fileName()
            self.label4.setText(filename2) 
            self.label4.move(30,160)
            self.label4.setFixedWidth(300)
            self.globalFileName = fileName
        if self.globalFileName:
           self.button3.setEnabled(True)
            
    #Saves text from textbox to variable 
    def loadText(self):
        textValue = self.textbox.toPlainText()
        if textValue == "":
            self.label5.setText("No points entered")
        else: 
            self.label5.setText("Loaded Point: " + textValue)
            textValue
            document = docx.Document('/Users/clairethompson/Desktop/EngHacks2021/ResumePoints.docx')
            if textValue:
                document.add_paragraph(textValue)
        self.label5.move(355,355)
        self.label5.adjustSize()
        self.textbox.clear()   

    def selectionChange(self):
            self.num_items = self.comboBox.currentText()
            self.num_items = int(self.num_items)
            return  self.num_items
    
    def backEnd(self):
        pdfFileObj = open(self.globalFileName, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        pageObj = pdfReader.getPage(0)
        pageObj1 = pdfReader.getPage(1)
        text = (pageObj.extractText()) + (pageObj1.extractText())
        post = text.split('Required Skills')
        Require = post[1].split('Application Information')
        Require = Require[0].split('Transportation and Housing')
        numberOfPoints = self.selectionChange()
        doc = docx.Document('/Users/clairethompson/Desktop/EngHacks2021/ResumePoints.docx')
        pa = doc.paragraphs
        point = []
        for i in range(len(pa)):
            if (i%2 ==0):
                 point.append(pa[i].text)

        keyPoints = self.textClean(point)
        Require = self.textClean(Require) 
        sim = []
        bestScore = 0
        newScore = 0
        sim = [0]*len(keyPoints)
        for i in range(len(keyPoints)):
            keyPoints[i] = keyPoints[i].lower()
            sentence = keyPoints[i].split()
            for u in range(len(sentence)):
                word = sentence[u]
                bestWord = word
                bestScore = 0
                for synonymSet in wordnet.synsets(word):
                    for lemma in synonymSet.lemmas():
                        tempSentence = " ".join(sentence)
                        newTempSentence = tempSentence.replace(word,lemma.name())
                        newScore = self.compareSentences(newTempSentence, Require[0].lower())
                        if newScore > bestScore:
                            bestScore = newScore
                            bestWord = lemma.name()
                sentence[u]=bestWord
            if bestScore > sim[i]:
                sim[i] = bestScore
                for i in range(len(pa)):
                    if (i%2 ==0):
                        point.append(pa[i].text)
        topPoints = sorted(range(len(sim)), key=lambda i: sim[i], reverse=True)[:numberOfPoints]
        for k in range(len(topPoints)):
            #This is where you change the bottom text box
            self.resultsbox.insertPlainText(f'{sim[topPoints[k]]*100:2.2f} ' + point[topPoints[k]] + '\n')
            self.resultsbox.move(100,400)
            self.resultsbox.resize(550,300)
            self.resultsbox.setReadOnly(True)
            self.resultsbox.verticalScrollBar().minimum()
            self.button4.move(100, 720) 
            self.button4.show()

            #print(f'{sim[topPoints[k]]*100:2.2f}')
            #print(point[topPoints[k]])

        pdfFileObj.close()  

    def textClean (self,point):
        keyP = []
        filler = ['good','to', 'and', 'a', 'for', 'with', 'by', 'in', 'the', 'as', 'such', 'of', 'experience', 'through', 'competencies', 'attained','completion', 'would','further', 'implemented', 'ensure', 'executing', 'at', 'previous', 'based','are', 'an', 'be', 'or', 'will', 'is', 'both', 'but', 'not', 'should', 'have', 'proficiency', 'related', 'taking', 'required', 'applicant', 'applicants','4-month', '8-month', 'one', 'two', 'three', 'pursuing']

        remov = string.punctuation
        remov = remov.replace("-", "")
        remov = remov.replace("/", "")

        for i in range(len(point)):
            words = point[i].split()
            for p in range(len(words)):
                words[p] = words[p].lower()
                table = str.maketrans(dict.fromkeys(remov))
                words[p] = words[p].translate(table)
                #words[p]= words[p].split("-", 1)[0]
                words[p]= words[p].split("/", 1)[0]
                words[p]= words[p].replace("Ł", " ")
                words[p]= words[p].replace("ł", " ")
                words[p]= words[p].replace("€€", " ")
    
            for j in range(len(filler)):
                while (filler[j] in words):
                    words.remove(filler[j])
                # print(words[0])

            text = " ".join(words)
            keyP.append(text)
            words.clear()
        return keyP 

    def compareSentences(self,app, point):
        score = 0

        appSet = app.split(' ')
        pointSet = point.split(' ')

        #lev, difflab SequenceMatcher
        bestConnectScore = 0
        connectScore = 0
        scoreArray = [0]*len(pointSet)
        for i in range(len(pointSet)):
            bestConnectScore = 0
            connectScore = 0
            for j in range(len(appSet)):
                connectScore = lev.ratio(pointSet[i], appSet[j])
                if connectScore > bestConnectScore:
                    bestConnectScore = connectScore
            scoreArray[i] = bestConnectScore

        score = sum(scoreArray) / len(scoreArray)

        return score

    def clear(self):
        self.resultsbox.clear()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

