import sys
import os
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from pandas import *
import datetime
import logging
import matplotlib.pyplot as plt
import numpy as np

# from PyQt5.QtTest import QTest
# import openpyxl
# import pybithumb

# logging.basicConfig(filename="log.txt", level=logging.ERROR)    # 테스트가 끝나면 이조건을 설정
logging.basicConfig(level=logging.INFO)

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        self.login_event_loop = QEventLoop()
        self.option_data_loop = QEventLoop()

        ## 키움API 인스턴스 생성
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

        ## 이벤트에 슬롯 연결
        self.set_event_slot()

        ## 키움 API 로그인
        self.login()

        ## 옵션월물별 시세표 가져오기
        # 콜종목 결제월별시세 요청
        self.dynamicCall("SetInputValue(QString,QString)","만기년월", "202104")
        self.dynamicCall("CommRqData(QString, QString, Int, QString)", "콜종목결제월별시세요청", "opt50021", 0, "7000")
        self.option_data_loop.exec_()

        # 풋종목 결제월별시세 요청
        self.dynamicCall("SetInputValue(QString,QString)","만기년월", "202104")
        self.dynamicCall("CommRqData(QString, QString, Int, QString)", "풋종목결제월별시세요청", "opt50022", 0, "7000")
        self.option_data_loop.exec_()

        #일일변동성차트 그리기
        self.dynamicCall("SetInputValue(QString,QString)","역사적변동성1", "30")
        self.dynamicCall("SetInputValue(QString,QString)","역사적변동성2", "60")
        self.dynamicCall("SetInputValue(QString,QString)","역사적변동성3", "100")
        self.dynamicCall("CommRqData(QString, QString, Int, QString)", "일별변동성분석그래프요청", "opt50024", 0, "7000")
        self.option_data_loop.exec_()


    def set_event_slot(self):
        self.OnEventConnect.connect(self.handler_login)
        self.OnReceiveTrData.connect(self.handler_tr_data)

    def login(self):
        self.dynamicCall("CommConnect()")
        self.login_event_loop.exec_()

    def handler_login(self, errCode):
        logging.info(f"handler login {errCode}")
        if errCode == 0:
            print("연결되었습니다")
        else:
            print("연결되지 않았습니다.")
        self.login_event_loop.exit()

    def handler_tr_data(self, scrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        logging.info(f"OnReceiveTrData {scrNo}, {sRQName}, {sTrCode}, {sRecordName}")

        if sRQName == "콜종목결제월별시세요청":
            print("콜종목결제월별시세요청입니다.")
            data_dict = {}
            cnt = self.dynamicCall("GetRepeatCnt(Qstring, Qstring)", sTrCode, sRecordName)
            print("cnt", cnt)
            for i in range(cnt):
                kospi_conversion = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "지수환산")
                atm = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "ATM구분")
                code = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "종목코드")
                exe_price = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "행사가")
                current_price = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "현재가")
                standard_price = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "기준가")
                theory_price = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "이론가")
                i_v = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "내재변동성")
                delta = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "델타")
                gamma= self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "감마")
                theta = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "세타")
                vega = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "베가")

                code = code.strip()

                data_dict['지수환산'] = kospi_conversion.strip()
                data_dict['행사가'] = exe_price.strip()
                data_dict['현재가'] = current_price.strip()
                data_dict['기준가'] = standard_price.strip()
                data_dict['이론가'] = theory_price.strip()
                data_dict['내재변동성'] = i_v.strip()
                data_dict['델타'] = delta.strip()
                data_dict['감마'] = gamma.strip()
                data_dict['세타'] = theta.strip()
                data_dict['베가'] = vega.strip()













                # print("data_dict", data_dict)

                if i == 0:
                    df = DataFrame(data_dict, index=[code])
                else:
                    df.loc[code]=data_dict

            print('df\n', df)

            # 콜시세표를 엑셀파일에 저장하기 ---> 파일명은 월물을 표시 sheet이름은 조회하는 시간을 표시.
            file_name = 'data/call_202104.xlsx'  # 월물표시
            now = str(datetime.datetime.now())     # 2021-04-07 17:22:39.244263

            sheet_name = now[2:4]+now[5:7]+now[8:10] + now[11:13] +now[14:16]       # '2104071035' 형태로 년월일시간분 표시
            print(sheet_name)

            if not os.path.exists(file_name):
                with ExcelWriter(file_name, mode='w', engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=True)
            else:
                with ExcelWriter(file_name, mode='a', engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=True)

            self.option_data_loop.exit()

            '''  
            if not self.calculator_event_loop.isRunning():
                self.calculator_event_loop.exec_()
            '''

        elif sRQName == "풋종목결제월별시세요청":
            print("풋종목결제월별시세요청입니다.")
            data_dict = {}
            cnt = self.dynamicCall("GetRepeatCnt(Qstring, Qstring)", sTrCode, sRecordName)
            print("cnt", cnt)
            for i in range(cnt):
                kospi_conversion = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i,
                                                    "지수환산")
                atm = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "ATM구분")
                code = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "종목코드")
                exe_price = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "행사가")
                current_price = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i,
                                                 "현재가")
                standard_price = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i,
                                                  "기준가")
                theory_price = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i,
                                                "이론가")
                i_v = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "내재변동성")
                delta = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "델타")
                gamma = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "감마")
                theta = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "세타")
                vega = self.dynamicCall("GetCommData(QString, QString, Int, QString)", sTrCode, sRQName, i, "베가")

                code = code.strip()

                data_dict['지수환산'] = kospi_conversion.strip()
                data_dict['행사가'] = exe_price.strip()
                data_dict['현재가'] = current_price.strip()
                data_dict['기준가'] = standard_price.strip()
                data_dict['이론가'] = theory_price.strip()
                data_dict['내재변동성'] = i_v.strip()
                data_dict['델타'] = delta.strip()
                data_dict['감마'] = gamma.strip()
                data_dict['세타'] = theta.strip()
                data_dict['베가'] = vega.strip()

                # print("data_dict", data_dict)

                if i == 0:
                    df = DataFrame(data_dict, index=[code])
                else:
                    df.loc[code] = data_dict

            print('df\n', df)

            # 풋시세표를 엑셀파일에 저장하기 ---> 파일명은 월물을 표시 sheet이름은 조회하는 시간을 표시.
            file_name = 'data/put_202104.xlsx'  # 월물표시
            now = str(datetime.datetime.now())  # 2021-04-07 17:22:39.244263

            sheet_name = now[2:4] + now[5:7] + now[8:10] + now[11:13] + now[14:16]  # '202104071035' 형태로 년월일시간분 표시
            print(sheet_name)

            if not os.path.exists(file_name):
                with ExcelWriter(file_name, mode='w', engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=True)
            else:
                with ExcelWriter(file_name, mode='a', engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=True)

            self.option_data_loop.exit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    app.exec_()