from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from config.config import *
from PyQt5.QtTest import *

import pickle
import os

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__() # super == 부모의 init함수를 사용하겠다
        
        print("Kiwoom 클래스 시작합니다.")
        
        ##### 이벤트 루프 모음 #####
        self.login_event_loop = None
        self.detail_account_info_event_loop= QEventLoop()
        self.calculator_event_loop = QEventLoop()
        
        ##### 스크린 번호 모음 #####
        self.screen_my_info = "2000"
        self.screen_calculation_stock = "4000"
        
        ##### 종목 정보용 초기 빈 딕셔너리 모음 #####
        self.account_stock_dict = {}
        
        ##### 종목 분석 용 초기 빈 리스트 모음 ##### 
        self.calcul_data = []
        self.day_stock = []
        
        ##### 계좌 정보 관련 변수 #####
        self.account_num = None # 계좌번호
        self.use_money = 0 #실제 투자에 사용할 금액
        self.use_money_percent = 0.5 #예수금에서 실제 사용할 비율

        
        ##### 지정 함수 호출 모음 #####
        self.get_ocx_instance() # OCX 방식을 파이썬에 사용할 수 있게 변환해주는 함수
        self.event_slot() # 키움과 연결하기 위한 시그널/슬롯 모음
        self.signal_login_commConnect() # 로그인 요청 시그널 포함
        self.get_account_info() # 계좌번호 정보 가져오기
        
        self.detail_account_info() # 증권 계좌 예수금 정보 가져오기
        self.detail_account_mystock() # 증권 계좌 잔고내역 정보 가져오기

        # self.calculator_fnc()
        
        self.OPT10023()
        #==========================================================================================================================================================
        #==========================================================================================================================================================


        
    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") # 레지스트리에 저장된 api 모듈 불러오기
        
        
    def event_slot(self):
        self.OnEventConnect.connect(self.login_slot) # 로그인 관련 이벤트
        self.OnReceiveTrData.connect(self.trdata_slot) # TRdata 요청 이벤트
        
        
    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()") # 로그인 요청 시그널
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()
        
        
    def login_slot(self, errCode):
        print(errors(errCode))
        
        self.login_event_loop.exit()
    
    
    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        
        self.account_num = account_list.split(';')[1] # 0: 선물옵션 계좌번호 / 1번 8042899611 : 모의투자 상시용 계좌 / 2번 8042899711 : 모의투자 대회용 계좌
        
        print("\n나의 보유 계좌번호: %s" % self.account_num)
        
        
    def detail_account_info(self):
        print("\n----예수금 요청 부분----")
        
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num) # 계좌번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", 0000) # 비밀번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", 00)
        self.dynamicCall("SetInputValue(String, String)", "조회구분", 2)
        
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청", "opw00001", "0", self.screen_my_info) # 마지막 화면번호 1000: 잔고조회 / 2000: 실시간 데이터 조회(종목 1~100) / 2001: 실시간 데이터 조회(종목 101~200) / 3000: 주문요청 / 4000: 일봉조회
    
        self.detail_account_info_event_loop.exec_()
    
    
    def detail_account_mystock(self,sPrevNext="0"):
        print("\n----계좌평가 잔고내역 요청 연속조회----")
        
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num) # 계좌번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", 0000) # 비밀번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", 00)
        self.dynamicCall("SetInputValue(String, String)", "조회구분", 2)
        
        self.dynamicCall("CommRqData(String, String, int, String)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.screen_my_info)
        
        self.detail_account_info_event_loop.exec_()
    
    
    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        -TR요청을 하는 slot-
        
        sScrNo: 스크린 번호
        sRQName: 내가 요청할 때 지은 이름
        sTrCode: 요청한 TR Code
        sRecordName: 사용 안함
        sPrevNext: 다음 페이지가 있는지 알려줌
        
        '''
        
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")
            print("예수금 : %s" % int(deposit), "원")
            
            self.use_money = int(deposit) * self.use_money_percent
            self.use_money = self.use_money / 4
            
            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "출금가능금액")
            print("출금가능금액: %s" % int(ok_deposit), "원")
            
            self.detail_account_info_event_loop.exit()
        #====================================================================================================================
        #====================================================================================================================   
        elif sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총매입금액")
            print("총매입금액 : %s" % int(total_buy_money), "원")
            
            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)")
            print("총수익률 : %s" % float(total_profit_loss_rate), "%")
            
            
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            cnt = 0
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:]
                
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")
                
                if code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict[code] = {}
                
                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())
                
                self.account_stock_dict[code].update({"종목명" : code_nm})
                self.account_stock_dict[code].update({"보유수량" : stock_quantity})
                self.account_stock_dict[code].update({"매입가" : buy_price})
                self.account_stock_dict[code].update({"수익률(%)" : learn_rate})
                self.account_stock_dict[code].update({"현재가" : current_price})
                self.account_stock_dict[code].update({"매입금액" : total_chegual_price})
                self.account_stock_dict[code].update({"매매가능수량" : possible_quantity})
                
                cnt += 1
            
            print("계좌 보유 종목 개수: %s" % cnt, "개")            
            
            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()
        #====================================================================================================================
        #====================================================================================================================
        elif sRQName == "거래량급증요청":
            
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName) # 1회 조회당 600틱까지 데이터를 받을 수 있음
            print("====  Data updating ... ====")
            print("거래량 급증 종목 조회 :  데이터 종목 수 %s" % cnt)
            
            for i in range(cnt):
                data = [] # 종목별.. 1일마다 빈 리스트로 만들어짐..
                stock_code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목코드")
                code_name = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                today_updown = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "전일대비")
                updown_ratio = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "등락률")
                before_volume = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "이전거래량")
                now_volume = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재거래량")
                increase_ratio = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "급증률") 
                
                # data.append("")
                data.append(stock_code.strip())
                data.append(code_name.strip())
                data.append(current_price.strip())
                data.append(today_updown.strip())
                data.append(updown_ratio.strip())
                data.append(before_volume.strip())
                data.append(now_volume.strip())
                data.append(increase_ratio.strip())
                # data.append("")
                
                self.calcul_data.append(data.copy()) # 종목별 1분씩 생성된 데이터를 self.calcul_data에 append 
                
            if sPrevNext == "2":
                print("==>> 다음 페이지 데이터 조회 ==>>\n")
                self.OPT10023(sPrevNext = sPrevNext)
            else:
                print("Stock Code : [ 총 %s 개 ] 수집 완료" % len(self.calcul_data))
                
                # 데이터 저장하는 부분
                path = './increase_stock_code'
                
                if not os.path.exists(path):
                    os.makedirs(path)
                kiwoom_day = os.path.join(path, 'increase_stock_code_data.pkl')
                
                with open(kiwoom_day, 'wb') as f:
                    pickle.dump(self.calcul_data, f)
                    
                self.calcul_data.clear() # 특정 종목 모든 일수 조회결과 저장 후 초기화
                self.calculator_event_loop.exit()

    
    def OPT10023(self, sPrevNext="0"):
        
        QTest.qWait(3600)
               
        self.dynamicCall("SetInputValue(QString, QString)", "시장구분", "001")
        self.dynamicCall("SetInputValue(QString, QString)", "정렬구분", "2") #급증률
        self.dynamicCall("SetInputValue(QString, QString)", "시간구분", "2") # 전일 대비
        self.dynamicCall("SetInputValue(QString, QString)", "거래량구분", "5") # 5천주 이상 
        self.dynamicCall("SetInputValue(QString, QString)", "시간", "분")
        self.dynamicCall("SetInputValue(QString, QString)", "종목조건", "1") # 관리종목 제외
        self.dynamicCall("SetInputValue(QString, QString)", "가격구분", "0")
            
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "거래량급증요청", "OPT10023", sPrevNext, "self.screen_calculation_stock")
        
        self.calculator_event_loop.exec_()