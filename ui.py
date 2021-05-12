'''
지수선물옵션 트레이딩을 위하여 다음과 같은 그래프를 그림
1) 옵션월물별시세표를 이용하여 등가격의 옵션시세(프리미엄) 추이를 나타내는 그래프(콜, 풋, 양합)
2) 등가격의 내재변동성 추이를 나타내는 그래프, 역사적변동성과 비교하여 표시.
3) 등가격의 델타, 세타, 베타 값 추이
4) 위 등가격외 5포인트, 10포인트, 15포인트 외가격에 대하여도 그래프 그림.
5) 위 그래프를 이용하여 월중 어느시점이 내재변동성이 가장 높은지 및 하루중 어느 시간대가 가장 매수적기인지를 판단.
** 별도로 후일 참고하기 위하여 30분(또는 10분) 단위로 월물별 시세표를 액셀에 자동 저장함.
'''


import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic

form_class = uic.loadUiType("main_window.ui")[0]

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.btn1_clicked)
        self.pushButton_2.clicked.connect(self.btn2_clicked)

    def btn1_clicked(self):
        self.label_2.setText("버튼이 클릭되었습니다.")

    def btn2_clicked(self):
        self.label_2.clear()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mywindow = MyWindow()
    mywindow.show()
    app.exec_()

