import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5.QtTest import *
from collections import defaultdict


class Kiwoom:
    def __init__(self):
        super().__init__()

        # Event Loop
        self.ocx = None                                 # 사용할 OCX
        self.loginEventLoop = QEventLoop()              # 로그인을 위한 이벤트 루프
        self.accEventLoop = QEventLoop()                # 계좌 관련을 위한 이벤트 루프
        self.calculatorEventLoop = QEventLoop()

        # Variables
        self.account_num = None                         # 계좌번호
        self.username = None                            # 사용자명
        self.deposit = None                             # 예수금
        self.withdraw_amount = None                     # 출금가능금액
        self.order_amount = None                        # 주문가능금액
        self.tBuyAmount = None                          # 총 매입금액: total buy amount
        self.tEvalAmount = None                         # 총 평가금액: total evaluation amount
        self.tProfit = None                             # 총 평가손익금액: total profit
        self.tYield = None                              # 총 수익률(%): total yield
        self.stock_account = defaultdict()              # 체결된 계좌: trade-completed
        self.not_ordered_account = defaultdict()        # 체결 안 된 계좌: open-orders
        self.portfolio_account = defaultdict()          # 조건을 성립한 종목 정보 ex) Granvile

        self.scrAccNum = '2000'                         # 계좌관련 화면 번호
        self.scrCalculationStock = '4000'               # 계산용 화면 번호

        self.calculatorList = []                        # 종목 분석 변수

        # Setting
        self.createInstance()                           # OCX를 파이썬에서 활용 / API 모듈 불러옴
        self.startEvent()                               # 키움과 연결하기 위한 시그널/이벤트
        self.startLogin()                               # 로그인 요청 함수
        self.getAccNum()                                # 계좌번호 요청 함수
        self.getUserName()                              # 사용자명 요청 함수
        self.getDepositInfo()                           # 예수금 상세 현황 요청
        self.getAccEvalBalance()                        # 계좌평가 잔고내역 요청
        QTimer.singleShot(5000, self.getNotSignedAcc)   # 5초 후 미체결 종목들 가져오기

        # Read Portfolio
        QTest.qWait(10000)                              # 이전 세팅 완료를 위한 10초 대기
        self.readPortfolio()

    def createInstance(self):
        ''' 
            To use OCX, we need CLSID / ProgID
            Kiwoom's CLSID: KHOPENAPI.KHOpenAPICtrl.1
        '''
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

    def startEvent(self):
        '''
            Start event related with
            1. login 2. transition
        '''
        self.ocx.OnEventConnect.connect(self.onEventConnect)
        self.ocx.OnReceiveTrData.connect(self.onReceiveTrData)

    def onEventConnect(self, err_code):
        '''
            Check if logged in successfully
        '''
        if not err_code:
            print('Login Successfully')
        else:
            print('ERR: ', err_code)
            sys.exit(0)
        self.loginEventLoop.exit()

    def startLogin(self):
        '''
            Ask for login Kiwoom
        '''
        self.ocx.dynamicCall("CommConnect()")
        self.loginEventLoop.exec_()

    def onReceiveTrData(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
            Notify when the server is ready for a transition request
            1. SetInputValue   -> set input before asking transition
            2. CommRqData      -> when ready to ask for transition, send request to the server
            3. OnReceiveTrData -> notifying client that server's ready
            4. GetCommData     -> client receiving data from the server
        '''
        print('1: ', sScrNo, sRQName, sTrCode, sRecordName, sPrevNext)

        if sRQName == '예수금상세현황요청':
            my_deposit = self.getCommData(sTrCode, sRQName, 0, "예수금")
            my_withdraw = self.getCommData(sTrCode, sRQName, 0, "출금가능금액")
            my_order = self.getCommData(sTrCode, sRQName, 0, "주문가능금액")

            self.deposit = int(my_deposit)
            self.withdraw_amount = int(my_withdraw)
            self.order_amount = int(my_order)

            self.disconnectRealData(self.scrAccNum)
            self.accEventLoop.exit()

        elif sRQName == '계좌평가잔고내역요청':
            setting_complete = (self.tBuyAmount or self.tEvalAmount
                                or self.tProfit or self.tYield)

            if not setting_complete:
                tBuyAmount = self.getCommData(sTrCode, sRQName, 0, "총매입금액")
                tEvalAmount = self.getCommData(sTrCode, sRQName, 0, "총평가금액")
                tProfit = self.getCommData(sTrCode, sRQName, 0, "총평가손익금액")
                tYield = self.getCommData(sTrCode, sRQName, 0, "총수익률(%)")

                self.tBuyAmount = int(tBuyAmount)
                self.tEvalAmount = int(tEvalAmount)
                self.tProfit = int(tProfit)
                self.tYield = float(tYield)

            count = self.getRepeatCnt(sTrCode, sRQName)

            for i in range(count):
                sCode = self.getCommData(sTrCode, sRQName, i, "종목번호")
                sName = self.getCommData(sTrCode, sRQName, i, "종목명")
                sEvalProfit = self.getCommData(sTrCode, sRQName, i, "평가손익")
                sYield = self.getCommData(sTrCode, sRQName, i, "수익률(%)")
                sBuyPrice = self.getCommData(sTrCode, sRQName, i, "매입가")
                sQuantity = self.getCommData(sTrCode, sRQName, i, "보유수량")
                sCurrPrice = self.getCommData(sTrCode, sRQName, i, "현재가")
                sAvailQuantity = self.getCommData(
                    sTrCode, sRQName, i, "매매가능수량")

                sCode = sCode[1:]
                sEvalProfit = int(sEvalProfit)
                sYield = float(sYield)
                sBuyPrice = int(sBuyPrice)
                sQuantity = int(sQuantity)
                sCurrPrice = int(sCurrPrice)
                sAvailQuantity = int(sAvailQuantity)

                self.stock_account[sCode].update({'종목명': sName})
                self.stock_account[sCode].update({'평가손익': sEvalProfit})
                self.stock_account[sCode].update({'수익률(%)': sYield})
                self.stock_account[sCode].update({'매입가': sBuyPrice})
                self.stock_account[sCode].update({'보유수량': sQuantity})
                self.stock_account[sCode].update({'매매가능수량': sAvailQuantity})
                self.stock_account[sCode].update({'현재가': sCurrPrice})

            if sPrevNext == '2':
                self.getAccEvalBalance(nPrevNext=2)
            else:
                self.disconnectRealData(self.scrAccNum)
                self.accEventLoop.exit()

        elif sRQName == '실시간미체결요청':
            count = self.getRepeatCnt(sTrCode, sRQName)      # <= 600 days

            for i in range(count):
                sCode = self.getCommData(sTrCode, sRQName, i, "종목코드")
                sOrdNum = self.getCommData(sTrCode, sRQName, i, "주문번호")
                sName = self.getCommData(sTrCode, sRQName, i, "종목명")
                sOrdType = self.getCommData(sTrCode, sRQName, i, "주문구분")
                sOrdPrice = self.getCommData(sTrCode, sRQName, i, "주문가격")
                sOrdStat = self.getCommData(sTrCode, sRQName, i, "주문상태")
                sOrdQuantity = self.getCommData(sTrCode, sRQName, i, "주문수량")
                notOrdQuantity = self.getCommData(sTrCode, sRQName, i, "미체결수량")
                orderedQuantity = self.getCommData(sTrCode, sRQName, i, "체결량")

                sOrdNum = int(sOrdNum)
                # +매수, +매수정정, -매도, -매도정정
                sOrdType = sOrdType.lstrip('+').lstrip('-')
                sOrdPrice = int(sOrdPrice)
                sOrdQuantity = int(sOrdQuantity)
                notOrdQuantity = int(notOrdQuantity)
                orderedQuantity = int(orderedQuantity)

                self.not_ordered_account[sOrdNum].update({'종목코드': sCode})
                self.not_ordered_account[sOrdNum].update({'종목명': sName})
                self.not_ordered_account[sOrdNum].update({'주문구분': sOrdType})
                self.not_ordered_account[sOrdNum].update({'주문가격': sOrdPrice})
                self.not_ordered_account[sOrdNum].update({'주문상태': sOrdStat})
                self.not_ordered_account[sOrdNum].update(
                    {'주문수량': sOrdQuantity})
                self.not_ordered_account[sOrdNum].update(
                    {'미체결수량': notOrdQuantity})
                self.not_ordered_account[sOrdNum].update(
                    {'체결량': orderedQuantity})

            if sPrevNext == '2':
                self.getNotSignedAcc(nPrevNext=2)
            else:
                self.disconnectRealData(sScrNo)
                self.accEventLoop.exit()

        elif sRQName == '주식일봉차트조회요청':
            # getCommDataEx -> Cannot get data exceeding 600 days
            # stockData = self.getCommDataEx(sTrCode, sRecordName)
            stockCode = self.getCommData(sTrCode, sRQName, 0, "종목코드")

            count = self.getRepeatCnt(sTrCode, sRQName)      # <= 600 days

            for i in range(count):
                calculatorList = []

                sCurrPrice = self.getCommData(sTrCode, sRQName, i, "현재가")
                sVolume = self.getCommData(sTrCode, sRQName, i, "거래량")
                sTrPrice = self.getCommData(sTrCode, sRQName, i, "거래대금")
                date = self.getCommData(sTrCode, sRQName, i, "일자")
                sStartPrice = self.getCommData(sTrCode, sRQName, i, "시가")
                sHighPrice = self.getCommData(sTrCode, sRQName, i, "고가")
                sLowPrice = self.getCommData(sTrCode, sRQName, i, "저가")

                calculatorList.append("")
                calculatorList.append(int(sCurrPrice))
                calculatorList.append(int(sVolume))
                calculatorList.append(int(sTrPrice))
                calculatorList.append(int(date))
                calculatorList.append(int(sStartPrice))
                calculatorList.append(int(sHighPrice))
                calculatorList.append(int(sLowPrice))

                self.calculatorList.append(calculatorList.copy())

            if sPrevNext == '2':
                self.checkAboveMA(120, stockCode, None, 2)      # 120일 이동평균선 확인
            else:

                pass
                enoughData = False
                ma = 120

                if (not self.calculatorList) or (len(self.calculatorList) < ma):
                    enoughData = False
                else:
                    totalPrice = sum(int(val[1])
                                     for val in self.calculatorList[:ma])
                    todayPrice_MA = totalPrice / ma

    def getCommData(self, trcode, rqname, index, item):
        '''
            Get received data from the server
        '''
        data = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)",
                                    trcode, rqname, index, item)
        return data.strip()

    def getCommDataEx(self, trcode, sRecordName):
        '''
            Get large amount of data from the server as an array (<=600 days)
        '''
        return self.ocx.dynamicCall("GetCommDataEx(QString, QString)", trcode, sRecordName)

    def setInputValue(self, id, value):
        '''
            Set input value when requesting transition
        '''
        self.ocx.dynamicCall("SetInputValue(QString, QString)", id, value)

    def CommRqData(self, rqname, trcode, next, screen):
        '''
            Use when ready to send the request to the server
        '''
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)",
                             rqname, trcode, next, screen)

    def getMasterCodeName(self, code):
        '''
            Change code to name    ex) '005930' -> 'Samsung'
        '''
        name = self.ocx.dynamicCall("GetMasterCodeName(QString)", code)
        return name

    def getCodeListByMarket(self, marketCode):
        '''
            Market Code {'0': KOSPI, '3': 'ELW', '4': 'Mutual Fund, '8': ETF, '10': KOSDAQ}
        '''
        code_list = self.ocx.dynamicCall(
            "GetCodeListByMarket(QString)", marketCode)
        # [005930;112351;1235523;...;423530;''] -> last one is empty
        code_list = code_list.split(";")[:-1]
        return code_list

    def getAccNum(self):
        # a;b;c -> [a, b, c]
        accounts = self.ocx.dynamicCall("GetLoginInfo(QString)", "ACCNO")
        self.account_num = accounts.split(';')[0]

    def getUserName(self):
        self.username = self.ocx.dynamicCall(
            "GetLoginInfo(QString)", "USER_NAME")

    def getRepeatCnt(self, sTrCode, sRQName):
        return self.ocx.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)

    def disconnectRealData(self, sScrNo):
        # After receiving data, need to disconnect to avoid unnecessary request
        self.ocx.dynamicCall("DisconnectRealData(QString)", sScrNo)

    def getDepositInfo(self, nPrevNext=0):
        '''
            Deposit: deposit, withdraw amount, order available amount
        '''
        self.setInputValue("계좌번호", self.account_num)
        self.setInputValue("비밀번호", " ")
        self.setInputValue("비밀번호입력매체구분", "00")
        self.setInputValue("조회구분", "2")
        self.CommRqData("예수금상세현황요청", "opw00001", nPrevNext, self.scrAccNum)

        self.accEventLoop.exec_()

    def getAccEvalBalance(self, nPrevNext=0):
        '''
            Account Evaluation: total purchase amount, total eval amount, total yield, ...
        '''
        self.setInputValue("계좌번호", self.account_num)
        self.setInputValue("비밀번호", " ")
        self.setInputValue("비밀번호입력매체구분", "00")
        self.setInputValue("조회구분", "1")
        self.CommRqData("계좌평가잔고내역요청", "opw00018", nPrevNext, self.scrAccNum)

        if not self.accEventLoop.isRunning():
            self.accEventLoop.exec_()

    def getNotSignedAcc(self, nPrevNext=0):
        '''
            Get Open Orders
        '''
        self.setInputValue("계좌번호", self.account_num)
        self.setInputValue("전체종목구분", "0")
        self.setInputValue("매매구분", "0")
        self.setInputValue("체결구분", "1")
        self.CommRqData("실시간미체결요청", "opt10075", nPrevNext, self.scrAccNum)

        if not self.accEventLoop.isRunning():
            self.accEventLoop.exec_()

    def checkEachCode(self):
        '''
            Get the code by code list and analyze using Granvile Theory
        '''
        code_list = self.getCodeListByMarket('10')

        for idx, code in enumerate(code_list):
            self.disconnectRealData(self.scrCalculationStock)

        self.checkAboveMA(ma=120, stockCode=code)

    def checkAboveMA(self, ma, stockCode=None, date=None, nPrevNext=0):
        '''
            Write stock code that is going above moving average given [Theory of Granvile]

            Args:
                ma: the number of days of moving average we want to find
        '''
        # At least 3.6s needed to analyze each sector => "10h ~" needed to analyze all stock in KOSDAQ
        QTest.qWait(3600)

        self.setInputValue("종목코드", stockCode)
        self.setInputValue("수정주가구분", 1)

        # If date is None, today is the date
        if date:
            self.setInputValue("기준일자", date)

        self.CommRqData("주식일봉차트조회요청", "opt10081",
                        nPrevNext, self.scrCalculationStock)

        if not self.calculatorEventLoop.isRunning():
            self.calculatorEventLoop.exec_()
        else:
            # Avg Price for the moving average given
            totalPrice = sum(int(val[1]) for val in self.calculatorList[:ma])
            todayPrice_MA = totalPrice / ma

            # check if today's "lowPrice <= todayPrice_MA <= highPrice"
            isTodayBetweenMA = False
            todayPrice = None
            lowPrice = int(self.calculatorList[0][7])
            highPrice = int(self.calculatorList[0][6])
            if (lowPrice <= todayPrice_MA <= highPrice):
                isTodayBetweenMA = True
                todayPrice = highPrice

            # check daily data for the past 20 days to check the day below MA
            prevPrice = None
            targetExist = False
            if isTodayBetweenMA:
                prevPrice_MA = 0
                isStockRising = False
                idx = 1

                while True:
                    if len(self.calculatorList[idx:]) < ma:
                        break

                    totalPrice = sum(int(val[1])
                                     for val in self.calculatorList[idx:idx+ma])
                    prevPrice_MA = totalPrice / ma

                    lowPrice = int(self.calculatorList[idx][7])
                    highPrice = int(self.calculatorList[idx][6])

                    if (prevPrice_MA <= highPrice) and (idx <= 20):
                        # 'high price' for 20 days should be lower than MA for 120 days to pass the condition
                        break

                    if (prevPrice_MA < lowPrice) and (idx > 20):
                        # find that 'low price', at least once, was higher than MA for 120 days
                        isStockRising = True
                        prevPrice = lowPrice
                        break

                    idx += 1

                targetExist = (isStockRising) and \
                    (todayPrice_MA > prevPrice_MA) and (todayPrice > prevPrice)

            if targetExist:
                stockName = self.getMasterCodeName(stockCode)
                f = open('files/targetStock.txt', 'a', encoding='UTF8')
                # calculaotrList[0][1] => current price
                f.write(
                    f'{stockCode}\t{stockName}\t{str(self.calculatorList[0][1])}\n')
                f.close()

            self.calculatorList.clear()
            self.calculatorEventLoop.exit()

    def readPortfolio(self):
        pass
