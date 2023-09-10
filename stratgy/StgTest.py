import sys
from PyQt5.QtWidgets import *
import win32com.client
from moneyCopy.common.MarketConnectTest import *

isSB = False
objCur = []

# 복수 종목 실시간 조회 샘플 (조회는 없고 실시간만 있음)
class CpEvent:
    def set_params(self, client):
        self.client = client

    def OnReceived(self):
        code = self.client.GetHeaderValue(0)  # 초
        name = self.client.GetHeaderValue(1)  # 초
        timess = self.client.GetHeaderValue(18)  # 초
        exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
        cprice = self.client.GetHeaderValue(13)  # 현재가
        diff = self.client.GetHeaderValue(2)  # 대비
        cVol = self.client.GetHeaderValue(17)  # 순간체결수량
        vol = self.client.GetHeaderValue(9)  # 거래량

        if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            print("실시간(예상체결)", name, timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        elif (exFlag == ord('2')):  # 장중(체결)
            print("실시간(장중 체결)", name, timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)

        print("로직추가여기서")


        stopAllSubscribe()



class CpStockCur(MarketConnectTest):

    def Subscribe(self, code):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        handler = win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        handler.set_params(self.objStockCur)
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()



def stgStart(self):

    self.stopAllSubscribe()

    marketConnectTest = MarketConnectTest()

    #itemSerach 에서 받아와야함
    codes = ["A003540", "A000660", "A005930", "A035420", "A069500", "Q530031"]
    # 요청 필드 배열 - 종목코드, 시간, 대비부호 대비, 현재가, 거래량, 종목명
    rqField = [0, 1, 2, 3, 4, 10, 17]  # 요청 필드

    #연결 췍
    if(marketConnectTest.Request(codes, rqField) == False):
        exit()

    cnt = len(codes)
    for i in range(cnt):
        self.objCur.append(CpStockCur())
        self.objCur[i].Subscribe(codes[i])

    print("빼기빼기================-")
    print(cnt, "종목 실시간 현재가 요청 시작")
    isSB = True


def stopAllSubscribe(self):
    if self.isSB:
        cnt = len(self.objCur)
        for i in range(cnt):
            self.objCur[i].Unsubscribe()
        print(cnt, "종목 실시간 해지되었음")
    self.isSB = False

    self.objCur = []

def stopOneSubscribe(self, idx):
    self.objCur[idx].Unsubscribe()





