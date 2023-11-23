import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from Kiwoom import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.kiwoom = Kiwoom()

        # 005930: Samsung
        self.kiwoom.setInputValue("종목코드", "005930")

        market_code = self.kiwoom.get_codelist_by_market('0')
        for i, code in enumerate(market_code[:10]):
            print(i+1, self.kiwoom.get_mastercode_name(code))

        username = self.kiwoom.get_username()
        print(username)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()
