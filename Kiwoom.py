import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from collections import defaultdict


class Kiwoom:
    def __init__(self):
        super().__init__()

        # Event Loop
        self.ocx = None
        self.login_eventLoop = QEventLoop()
        self.account_eventLoop = QEventLoop()

        # variables
        self.account_num = None
        self.username = None
        self.deposit = None                     # 예수금
        self.withdraw_amount = None             # 출금가능금액
        self.order_amount = None                # 주문가능금액
        self.total_buy_amount = None            # 총 매입금액
        self.total_eval_amount = None           # 총 평가금액
        self.total_profit = None                # 총 평가손익금액
        self.total_yield = None                 # 총 수익률 (%)
        self.stock_account = defaultdict()

        self.screen_num = '1000'

        # Before running
        self.create_instance()
        self.start_event()
        self.start_login()
        self.get_account_num()
        self.get_username()
        self.get_deposit_info()                 # 예수금 정보

        # Setting

    def create_instance(self):
        ''' 
            To use OCX, we need CLSID / ProgID
            Kiwoom's CLSID: KHOPENAPI.KHOpenAPICtrl.1
        '''
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

    def start_event(self):
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
        self.login_eventLoop.exit()

    def start_login(self):
        '''
            Ask for login Kiwoom
        '''
        self.ocx.dynamicCall("CommConnect()")
        self.login_eventLoop.exec_()

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
            self.deposit = int(my_deposit)

            my_withdraw = self.getCommData(sTrCode, sRQName, 0, "출금가능금액")
            self.withdraw_amount = int(my_withdraw)

            my_order = self.getCommData(sTrCode, sRQName, 0, "주문가능금액")
            self.order_amount = int(my_order)

            self.cancel_realData(self.screen_num)
            self.account_eventLoop.exit()

        elif sRQName == '계좌평가잔고내역요청':
            setting_complete = (self.total_buy_amount or self.total_eval_amount
                                or self.total_profit or self.total_yield)
            if not setting_complete:
                total_buy_amount = self.getCommData(
                    sTrCode, sRQName, 0, "총매입금액")
                self.total_buy_amount = int(total_buy_amount)

                total_eval_amount = self.getCommData(
                    sTrCode, sRQName, 0, "총평가금액")
                self.total_eval_amount = int(total_eval_amount)

                total_profit = self.getCommData(sTrCode, sRQName, 0, "총평가손익금액")
                self.total_profit = int(total_profit)

                total_yield = self.getCommData(sTrCode, sRQName, 0, "총수익률(%)")
                self.total_yield = float(total_yield)

            count = self.getRepeatCnt(sTrCode, sRQName)

            for i in range(count):
                stock_code = self.getCommData(sTrCode, sRQName, i, "종목번호")
                stock_code = stock_code[1:]

                stock_name = self.getCommData(sTrCode, sRQName, i, "종목명")

                stock_eval_profit = self.getCommData(
                    sTrCode, sRQName, i, "평가손익")
                stock_eval_profit = int(stock_eval_profit)

                stock_yield = self.getCommData(sTrCode, sRQName, i, "수익률(%)")
                stock_yield = float(stock_yield)

                stock_purchase_price = self.getCommData(
                    sTrCode, sRQName, i, "매입가")
                stock_purchase_price = int(stock_purchase_price)

                stock_quantity = self.getCommData(sTrCode, sRQName, i, "보유수량")
                stock_quantity = int(stock_quantity)

                stock_trade_avail_quantity = self.getCommData(
                    sTrCode, sRQName, i, "매매가능수량")
                stock_trade_avail_quantity = int(stock_trade_avail_quantity)

                stock_present_price = self.getCommData(
                    sTrCode, sRQName, i, "현재가")
                stock_present_price = int(stock_present_price)

                self.stock_account[stock_code].update(
                    {'종목명': stock_name})
                self.stock_account[stock_code].update(
                    {'평가손익': stock_eval_profit})
                self.stock_account[stock_code].update(
                    {'수익률(%)': stock_yield})
                self.stock_account[stock_code].update(
                    {'매입가': stock_purchase_price})
                self.stock_account[stock_code].update(
                    {'보유수량': stock_quantity})
                self.stock_account[stock_code].update(
                    {'매매가능수량': stock_trade_avail_quantity})
                self.stock_account[stock_code].update(
                    {'현재가': stock_present_price})

            if sPrevNext == '2':
                self.get_account_eval_balance(2)
            else:
                self.cancel_realData(self.screen_num)
                self.account_eventLoop.exit()

    def getCommData(self, trcode, rqname, index, item):
        '''
            Get received data from the server
        '''
        data = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)",
                                    trcode, rqname, index, item)
        return data.strip()

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

    def get_mastercode_name(self, code):
        '''
            Change code to name    ex) '005930' -> 'Samsung'
        '''
        name = self.ocx.dynamicCall("GetMasterCodeName(QString)", code)
        return name

    def get_codelist_by_market(self, market_code):
        '''
            Market Code {'0': KOSPI, '3': 'ELW', '4': 'Mutual Fund, '8': ETF, '10': KOSDAQ}
        '''
        code_list = self.ocx.dynamicCall(
            "GetCodeListByMarket(QString)", market_code)
        code_list = code_list.split(";")[:-1]
        return code_list

    def get_account_num(self):
        accounts = self.ocx.dynamicCall("GetLoginInfo(QString)", "ACCNO")
        self.account_num = accounts.split(';')[0]

    def get_username(self):
        username = self.ocx.dynamicCall("GetLoginInfo(QString)", "USER_NAME")
        self.username = username

    def getRepeatCnt(self, sTrCode, sRQName):
        return self.ocx.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)

    def cancel_realData(self, sScrNo):
        self.ocx.dynamicCall("DisconnectRealData(QString)", sScrNo)

    def get_deposit_info(self, nPrevNext=0):
        '''
            Deposit: deposit, withdraw amount, order available amount
        '''
        self.setInputValue("계좌번호", self.account_num)
        self.setInputValue("비밀번호", " ")
        self.setInputValue("비밀번호입력매체구분", "00")
        self.setInputValue("조회구분", "2")
        self.setInputValue("예수금상세현황요청", "opw00001", nPrevNext, self.screen_num)

        self.account_eventLoop.exec_()

    def get_account_eval_balance(self, nPrevNext=0):
        '''
            Account Evaluation: total purchase amount, total eval amount, total yield, ...
        '''
        self.setInputValue("계좌번호", self.account_num)
        self.setInputValue("비밀번호", " ")
        self.setInputValue("비밀번호입력매체구분", "00")
        self.setInputValue("조회구분", "1")
        self.setInputValue("계좌평가잔고내역요청", "opw00018",
                           nPrevNext, self.screen_num)

        if not self.account_eventLoop.isRunning():
            self.account_eventLoop.exec_()
