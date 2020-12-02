# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'plotto.ui'
#
# Created by: PyQt5 UI code generator 5.13.2
#
# WARNING! All changes made in this file will be lost!

import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QColor
import random
from datetime import datetime
from bs4 import BeautifulSoup
import requests
import operator
import tkinter as tk
import tkinter.filedialog
from urllib.request import urlopen
import openpyxl
import win32com.client
from tkinter import *


class WidgetsDemo:
    def __init__(self):
        window = Tk()
        window.title("로또세금계산기")
        window.geometry("380x200")
        frame0 = Frame(window)
        frame0.pack()

        label1 = Label(frame0, text="로또세금계산기")
        label1.pack()

        frame1 = Frame(window)
        frame1.pack()

        self.name = StringVar(window, value='')
        self.entryName = Entry(frame1, textvariable=self.name, width=45)
        self.entryName.grid(row=1, column=1)

        frame2 = Frame(window)
        frame2.pack()

        Button(frame2, text="1", command=lambda: self.test('1'), width=10, height=2).grid(row=0, column=0)
        Button(frame2, text="2", command=lambda: self.test('2'), width=10, height=2).grid(row=0, column=1)
        Button(frame2, text="3", command=lambda: self.test('3'), width=10, height=2).grid(row=0, column=2)
        Button(frame2, text="4", command=lambda: self.test('4'), width=10, height=2).grid(row=1, column=0)
        Button(frame2, text="5", command=lambda: self.test('5'), width=10, height=2).grid(row=1, column=1)
        Button(frame2, text="6", command=lambda: self.test('6'), width=10, height=2).grid(row=1, column=2)
        Button(frame2, text="7", command=lambda: self.test('7'), width=10, height=2).grid(row=2, column=0)
        Button(frame2, text="8", command=lambda: self.test('8'), width=10, height=2).grid(row=2, column=1)
        Button(frame2, text="9", command=lambda: self.test('9'), width=10, height=2).grid(row=2, column=2)
        Button(frame2, text="0", command=lambda: self.test('0'), width=10, height=2).grid(row=3, column=1)

        Button(frame2, text="=", command=self.processbutton, width=10, height=2).grid(row=3, column=2)
        Button(frame2, text="AC", command=self.deletebutton, width=10, height=2).grid(row=3, column=0)
        window.mainloop()

    def test(self, value):
        self.entryName.insert("end", value)

    def deletebutton(self):
        self.entryName.delete(0, "end")

    def processbutton(self):
        self.save_r = self.name.get()
        self.entryName.delete(0, "end")
        # result = self.cal()
        tmp = int(self.save_r)
        if tmp < 300000000:
            result = tmp * 0.78
        else:
            result = tmp * 0.67

        self.entryName.insert("end", result)

    def cal(self):
        if self.val == '+':
            return int(self.save_f) + int(self.save_r)
        elif self.val == '-':
            return int(self.save_f) - int(self.save_r)
        elif self.val == '/':
            return int(self.save_f) / int(self.save_r)
        else:
            return int(self.save_f) * int(self.save_r)

    def symbol(self, value):
        self.save_f = self.name.get()
        self.entryName.delete(0, "end")
        self.val = value


Numberlist = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24,
              25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45]

# 엑셀 파일 읽고 숫자 별 횟수, 총합 구하기
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('C:\\Users\\ciko1\\OneDrive\\바탕 화면\\로또\\lotto.xlsx')
ws = wb.ActiveSheet

i = 1
j = 0
add = 0
k = 1
a = [0, ]

for i in range(45):
    i += 1
    j += 2
    a.insert(i, ws.Cells(1 + j, 2).Value)

excel.Quit()

a = list(map(int, a))

for k in range(45):
    add += a[k]

tnum1 = a[1]
tnum2 = a[2]
tnum3 = a[3]
tnum4 = a[4]
tnum5 = a[5]
tnum6 = a[6]
tnum7 = a[7]
tnum8 = a[8]
tnum9 = a[9]
tnum10 = a[10]
tnum11 = a[11]
tnum12 = a[12]
tnum13 = a[13]
tnum14 = a[14]
tnum15 = a[15]
tnum16 = a[16]
tnum17 = a[17]
tnum18 = a[18]
tnum19 = a[19]
tnum20 = a[20]
tnum21 = a[21]
tnum22 = a[22]
tnum23 = a[23]
tnum24 = a[24]
tnum25 = a[25]
tnum26 = a[26]
tnum27 = a[27]
tnum28 = a[28]
tnum29 = a[29]
tnum30 = a[30]
tnum31 = a[31]
tnum32 = a[32]
tnum33 = a[33]
tnum34 = a[34]
tnum35 = a[35]
tnum36 = a[36]
tnum37 = a[37]
tnum38 = a[38]
tnum39 = a[39]
tnum40 = a[40]
tnum41 = a[41]
tnum42 = a[42]
tnum43 = a[43]
tnum44 = a[44]
tnum45 = a[45]


# GUI
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(470, 592)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(180, 10, 111, 21))
        self.label.setObjectName("label")

        # 로또 번호 생성
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(40, 120, 111, 31))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.onRandomClick)

        # 확률 적용
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(180, 120, 111, 31))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.clicked.connect(self.onProbability)

        # 로또 세금 계산기
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(320, 160, 111, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.showLottoCalc)

        # 보너스 번호 포함
        self.pushButton_5 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_5.setGeometry(QtCore.QRect(320, 120, 111, 31))
        self.pushButton_5.setObjectName("pushButton_5")
        self.pushButton_5.clicked.connect(self.onnoBonus)

        # 출력물
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(40, 210, 391, 271))
        self.textEdit.setObjectName("textEdit")
        self.textEdit.setText(''' 이 프로그램은 로또 번호 출력 프로그램 입니다.\n\n 우선 상단에 이번 주 당첨 번호를 표시하고 있습니다.
\n 그리고 아래에 버튼 4개가 있습니다. \n1. 로또 번호 생성 버튼은 무작위로 로또 번호를 생성해 줍니다.
2. 확률 적용은 1회부터 현재 까지 1등, 2등 당첨번호의 빈도수를 계산하여 번호별로 확률을 구해 적용시켜서 번호를 생성해 줍니다.
3. 보너스 번호 미포함 버튼은 보너스 번호를 미포함한 번호의 빈도수를 확률에 적용하고 번호를 생성해 줍니다.
4. 로또 세금 계산기는 로또 당첨 금액을 입력 했을 때 실수령액을 확인 할 수 있는 기능입니다.\n
번호 생성을 하게되면 이곳 텍스트 박스에 번호를 생성받습니다.\n\n 마지막으로 아래에 번호 저장하기 버튼을 누르면 생성 받은 번호를 .txt 파일로 저장 할 수 있습니다''')

        # 번호 저장하기
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(190, 500, 101, 31))
        self.pushButton_4.setObjectName("pushButton_4 ")
        self.pushButton_4.clicked.connect(self.save)

        # 이번 주 당첨 번호
        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser.setGeometry(QtCore.QRect(40, 40, 391, 51))
        self.textBrowser.setObjectName("textBrowser")
        self.textBrowser.setFontPointSize(25)
        self.textBrowser.setTextColor(QColor(107, 102, 255))
        self.textBrowser.append(result2)
        MainWindow.setCentralWidget(self.centralwidget)

        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 470, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)

        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "확률 적용 로또 프로그램"))
        self.pushButton.setText(_translate("MainWindow", "로또 번호 생성"))
        self.pushButton_3.setText(_translate("MainWindow", "확률 적용"))
        self.pushButton_2.setText(_translate("MainWindow", "로또 세금 계산기"))
        self.pushButton_5.setText(_translate("MainWindow", "보너스 번호 미포함"))
        self.pushButton_4.setText(_translate("MainWindow", "번호 저장하기"))
        self.label.setText(_translate("MainWindow", "이번 주 당첨 번호"))

    # 메모장에 저장
    def save(self):
        save_text = filedialog.asksaveasfile(title='다른 이름으로 저장', initialfile='lotto.txt')

        txt = self.textEdit.toPlainText()
        save_text.write(txt)
        save_text.close()

    # 이온유수정
    def showLottoCalc(self):
        print("showLottoCalc")
        WidgetsDemo()

    # 랜덤 생성 함수
    def onRandomClick(self):
        Alphabet = ['A', 'B', 'C', 'D', 'E']

        Label = ''
        Label += '로또 번호 랜덤 추첨\n\n'
        Label += '발행 시간 : ' + str(datetime.now()) + '\n\n'
        Label += '-------------------------\n\n'

        for i in range(5):
            arr = []
            use = [False] * 46
            random.seed(random.randrange(1, 1000000))
            for j in range(6):

                num = random.randrange(1, 46)
                if not use[num]:
                    use[num] = True
                    if num < 10:
                        string = '0'
                        string += str(num)
                        arr.append(string)
                    else:
                        arr.append(str(num))
                else:
                    while 1:
                        num = random.randrange(1, 46)
                        if not use[num]:
                            use[num] = True
                            if num < 10:
                                string = '0'
                                string += str(num)
                                arr.append(string)
                            else:
                                arr.append(str(num))
                            break

            arr.sort()
            Label += '  ' + str(Alphabet[i]) + '  자  동 ' + str(arr[0]) + ' ' + str(arr[1]) + ' ' + str(
                arr[2]) + ' ' + str(arr[3]) + ' ' + str(arr[4]) + ' ' + str(arr[5]) + '\n\n'
        Label += '-------------------------'
        self.textEdit.setText(Label)

    # 확률 적용 함수
    def onProbability(self):
        Alphabet = ['A', 'B', 'C', 'D', 'E']

        Label = ''
        Label += '로또 번호 랜덤 추첨\n\n'
        Label += '발행 시간 : ' + str(datetime.now()) + '\n\n'
        Label += '-------------------------\n\n'

        for i in range(5):
            arr = []
            use = [False] * 46
            random.seed(random.randrange(1, 1000000))
            for j in range(6):
                num = random.choices(Numberlist, weights=(
                tnum1 / add, tnum2 / add, tnum3 / add, tnum4 / add, tnum5 / add, tnum6 / add, tnum7 / add, tnum8 / add,
                tnum9 / add, tnum10 / add, tnum11 / add, tnum12 / add, tnum13 / add, tnum14 / add, tnum15 / add,
                tnum16 / add, tnum17 / add, tnum18 / add, tnum19 / add,
                tnum20 / add, tnum21 / add, tnum22 / add, tnum23 / add, tnum24 / add, tnum25 / add, tnum26 / add,
                tnum27 / add, tnum28 / add, tnum29 / add, tnum30 / add, tnum31 / add, tnum32 / add, tnum33 / add,
                tnum34 / add, tnum35 / add, tnum36 / add, tnum37 / add,
                tnum38 / add, tnum39 / add, tnum40 / add, tnum41 / add, tnum42 / add, tnum43 / add, tnum44 / add,
                tnum45 / add), k=1)
                rnum = str(num)
                rnum = rnum.strip('[')
                rnum = rnum.strip(']')
                num = int(rnum)

                if not use[num]:
                    use[num] = True
                    if num < 10:
                        string = '0'
                        string += str(num)
                        arr.append(string)
                    else:
                        arr.append(str(num))
                else:
                    while 1:
                        num = random.choices(Numberlist, weights=(
                        tnum1 / add, tnum2 / add, tnum3 / add, tnum4 / add, tnum5 / add, tnum6 / add, tnum7 / add,
                        tnum8 / add, tnum9 / add, tnum10 / add, tnum11 / add, tnum12 / add, tnum13 / add, tnum14 / add,
                        tnum15 / add, tnum16 / add, tnum17 / add, tnum18 / add, tnum19 / add,
                        tnum20 / add, tnum21 / add, tnum22 / add, tnum23 / add, tnum24 / add, tnum25 / add,
                        tnum26 / add, tnum27 / add, tnum28 / add, tnum29 / add, tnum30 / add, tnum31 / add,
                        tnum32 / add, tnum33 / add, tnum34 / add, tnum35 / add, tnum36 / add, tnum37 / add,
                        tnum38 / add, tnum39 / add, tnum40 / add, tnum41 / add, tnum42 / add, tnum43 / add,
                        tnum44 / add, tnum45 / add), k=1)
                        rnum = str(num)
                        rnum = rnum.strip('[')
                        rnum = rnum.strip(']')
                        num = int(rnum)
                        if not use[num]:
                            use[num] = True
                            if num < 10:
                                string = '0'
                                string += str(num)
                                arr.append(string)
                            else:
                                arr.append(str(num))
                            break

            arr.sort()
            Label += '  ' + str(Alphabet[i]) + '  자  동 ' + str(arr[0]) + ' ' + str(arr[1]) + ' ' + str(
                arr[2]) + ' ' + str(arr[3]) + ' ' + str(arr[4]) + ' ' + str(arr[5]) + '\n\n'
        Label += '-------------------------'
        self.textEdit.setText(Label)

    # 보너스 번호 미포함
    def onnoBonus(self):
        Alphabet = ['A', 'B', 'C', 'D', 'E']

        Label = ''
        Label += '로또 번호 랜덤 추첨\n\n'
        Label += '발행 시간 : ' + str(datetime.now()) + '\n\n'
        Label += '-------------------------\n\n'

        for i in range(5):
            arr = []
            use = [False] * 46
            random.seed(random.randrange(1, 1000000))
            for j in range(6):
                num = random.choices(Numberlist, weights=(
                tnum1 / add, tnum2 / add, tnum3 / add, tnum4 / add, tnum5 / add, tnum6 / add, tnum7 / add, tnum8 / add,
                tnum9 / add, tnum10 / add, tnum11 / add, tnum12 / add, tnum13 / add, tnum14 / add, tnum15 / add,
                tnum16 / add, tnum17 / add, tnum18 / add, tnum19 / add,
                tnum20 / add, tnum21 / add, tnum22 / add, tnum23 / add, tnum24 / add, tnum25 / add, tnum26 / add,
                tnum27 / add, tnum28 / add, tnum29 / add, tnum30 / add, tnum31 / add, tnum32 / add, tnum33 / add,
                tnum34 / add, tnum35 / add, tnum36 / add, tnum37 / add,
                tnum38 / add, tnum39 / add, tnum40 / add, tnum41 / add, tnum42 / add, tnum43 / add, tnum44 / add,
                tnum45 / add), k=1)
                rnum = str(num)
                rnum = rnum.strip('[')
                rnum = rnum.strip(']')
                num = int(rnum)

                if not use[num]:
                    use[num] = True
                    if num < 10:
                        string = '0'
                        string += str(num)
                        arr.append(string)
                    else:
                        arr.append(str(num))
                else:
                    while 1:
                        num = random.choices(Numberlist, weights=(
                        tnum1 / add, tnum2 / add, tnum3 / add, tnum4 / add, tnum5 / add, tnum6 / add, tnum7 / add,
                        tnum8 / add, tnum9 / add, tnum10 / add, tnum11 / add, tnum12 / add, tnum13 / add, tnum14 / add,
                        tnum15 / add, tnum16 / add, tnum17 / add, tnum18 / add, tnum19 / add,
                        tnum20 / add, tnum21 / add, tnum22 / add, tnum23 / add, tnum24 / add, tnum25 / add,
                        tnum26 / add, tnum27 / add, tnum28 / add, tnum29 / add, tnum30 / add, tnum31 / add,
                        tnum32 / add, tnum33 / add, tnum34 / add, tnum35 / add, tnum36 / add, tnum37 / add,
                        tnum38 / add, tnum39 / add, tnum40 / add, tnum41 / add, tnum42 / add, tnum43 / add,
                        tnum44 / add, tnum45 / add), k=1)
                        rnum = str(num)
                        rnum = rnum.strip('[')
                        rnum = rnum.strip(']')
                        num = int(rnum)
                        if not use[num]:
                            use[num] = True
                            if num < 10:
                                string = '0'
                                string += str(num)
                                arr.append(string)
                            else:
                                arr.append(str(num))
                            break

            arr.sort()
            Label += '  ' + str(Alphabet[i]) + '  자  동 ' + str(arr[0]) + ' ' + str(arr[1]) + ' ' + str(
                arr[2]) + ' ' + str(arr[3]) + ' ' + str(arr[4]) + ' ' + str(arr[5]) + '\n\n'
        Label += '-------------------------'
        self.textEdit.setText(Label)


# 상단 이번 주 당첨 번호 크롤링
thisweek = requests.get(
    'https://search.naver.com/search.naver?sm=top_sug.pre&fbm=0&acr=3&acq=%EB%A1%9C%EB%98%90+%EB%8B%B9%EC%B2%A8&qdt=0&ie=utf8&query=%EB%A1%9C%EB%98%90+%EB%8B%B9%EC%B2%A8%EB%B2%88%ED%98%B8')
soup = BeautifulSoup(thisweek.text, "html.parser")
result = str(soup.find('div', class_='num_box').text)
result1 = result.replace('보너스번호', ' ')
result2 = result1.replace('내 번호 당첨조회', ' ')

# 크롤링 후 엑셀에 저장
excel_file = openpyxl.Workbook()
excel_sheet = excel_file.active
excel_sheet.column_dimensions['B'].width = 100
excel_sheet.append(['수', '횟수'])

session = requests.session()
r = session.get("https://dhlottery.co.kr/gameResult.do?method=statByNumber")
soup = BeautifulSoup(r.text, "html.parser")

times = soup.find_all('table', id='printTarget')[0].find_all('td', {'class': ''})

num = 0

for td in times:
    num += 1
    excel_sheet.append([num, td.get_text()])

cell_A1 = excel_sheet['A1']
cell_A1.alignment = openpyxl.styles.Alignment(horizontal="center")

cell_B1 = excel_sheet['B1']
cell_B1.alignment = openpyxl.styles.Alignment(horizontal="center")

excel_file.save('lotto.xlsx')
excel_file.close()

if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
