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

        # variables
        self.account_num = None
        self.username = None

        # Before running
        self.create_instance()
        self.start_event()
        self.start_login()

        # Setting
        self.account_num = self.get_account_num()
        self.username = self.get_username()

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
        '''
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
        self.ocx.dynamicCall(
            "CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen)

    def get_mastercode_name(self, code):
        '''
            Change code to name    ex) '005930' -> 'Samsung'
        '''
        name = self.ocx.dynamicCall("GetMasterCodeName(QString)", code)
        return name

    def get_codelist_by_market(self, market_code):
        '''
            Market Code
            {'0': KOSPI, '3': 'ELW', '4': 'Mutual Fund,
             '8': ETF, '10': KOSDAQ}
        '''
        code_list = self.ocx.dynamicCall(
            "GetCodeListByMarket(QString)", market_code)
        code_list = code_list.split(";")[:-1]
        return code_list

    def get_account_num(self):
        acc_list = self.ocx.dynamicCall("GetLoginInfo(QString)", "ACCNO")
        return acc_list.split(';')[0]

    def get_username(self):
        username = self.ocx.dynamicCall("GetLoginInfo(QString)", "USER_NAME")
        return username
