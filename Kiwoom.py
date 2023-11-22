import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *


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
        if sRQName == '예수금상세현황요청':
            my_deposit = self.ocx.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "예수금")
            self.deposit = int(my_deposit)

            my_withdraw = self.ocx.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "출금가능금액")
            self.withdraw_amount = int(my_withdraw)

            my_order = self.ocx.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "주문가능금액")
            self.order_amount = int(my_order)

            self.cancel_realData(self.screen_num)
            self.account_eventLoop.exit()

        print(sScrNo, sRQName, sTrCode, sRecordName, sPrevNext)

    def getCommData(self, trcode, rqname, index, item):
        '''
            Get received data from the server
        '''
        data = self.ocx.dynamicCall(
            "GetCommData(QString, QString, int, QString)", trcode, rqname, index, item)
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
        code_list = self.ocx.dynamicCall("GetCodeListByMarket(QString)",
                                         market_code)
        code_list = code_list.split(";")[:-1]
        return code_list

    def get_account_num(self):
        accounts = self.ocx.dynamicCall("GetLoginInfo(QString)", "ACCNO")
        self.account_num = accounts.split(';')[0]

    def get_username(self):
        username = self.ocx.dynamicCall("GetLoginInfo(QString)", "USER_NAME")
        self.username = username

    def cancel_realData(self, sScrNo):
        self.ocx.dynamicCall("DisconnectRealData(QString)", sScrNo)

    def get_deposit_info(self, nPrevNext=0):
        self.ocx.dynamicCall("SetInputValue(QString, QString)",
                             "계좌번호", self.account_num)
        self.ocx.dynamicCall("SetInputValue(QString, QString)",
                             "비밀번호", " ")
        self.ocx.dynamicCall("SetInputValue(QString, QString)",
                             "비밀번호입력매체구분", "00")
        self.ocx.dynamicCall("SetInputValue(QString, QString)",
                             "조회구분", "2")
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)",
                             "예수금상세현황요청", "opw00001", nPrevNext, self.screen_num)

        self.account_eventLoop.exec_()
