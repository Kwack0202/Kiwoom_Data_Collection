from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from PyQt5.QtTest import *

import datetime
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
        
        self.portfolio_stock_dict = {}
        
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
        
    
        self.OPT50029_101T6000()
        self.OPT50029_105T4000()

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
        
        self.account_num = account_list.split(';')[0] # 0: 선물옵션 계좌번호 / 1번 8042899611 : 모의투자 상시용 계좌 / 2번 8042899711 : 모의투자 대회용 계좌
        
        
        print("\n나의 보유 계좌번호: %s" % self.account_num)
        
        
    def detail_account_info(self):
        print("\n----예탁금 및 증거금 요청 부분----")
        
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num) # 계좌번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", 0000) # 비밀번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", 00)
        
        self.dynamicCall("CommRqData(String, String, int, String)", "선옵예탁금및증거금조회요청", "OPW20010", "0", self.screen_my_info) # 마지막 화면번호 1000: 잔고조회 / 2000: 실시간 데이터 조회(종목 1~100) / 2001: 실시간 데이터 조회(종목 101~200) / 3000: 주문요청 / 4000: 일봉조회
    
        self.detail_account_info_event_loop.exec_()
    
    
    def detail_account_mystock(self,sPrevNext="0"):
        print("\n----선옵 잔고 현황 정산가 기준 요청 부분----")
        
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num) # 계좌번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", 0000) # 비밀번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", 00)
        
        self.dynamicCall("CommRqData(String, String, int, String)", "선옵잔고현황정산가기준요청", "opw20007", sPrevNext, self.screen_my_info)
        
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
        
        if sRQName == "선옵예탁금및증거금조회요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예탁총액")
            print("예탁총액 : %s" % int(deposit), "원")
            
            
            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "인출가능총액")
            print("인출가능총액: %s" % int(ok_deposit), "원")
            
            self.detail_account_info_event_loop.exit()
        #====================================================================================================================
        #====================================================================================================================   
        elif sRQName == "선옵잔고현황정산가기준요청":
            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "약정금액합계")
            print("약정금액합계 : %s" % int(total_buy_money), "원")
            
            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "평가손익합계")
            print("평가손익합계 : %s" % int(total_profit_loss_rate), "원")
            
            self.detail_account_info_event_loop.exit()
            
        #====================================================================================================================
        #====================================================================================================================
        elif sRQName == "선물옵션분차트요청_101T6000":
            print("101T6000 분봉데이터 요청") 
            
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName) # 1회 조회당 900틱까지 데이터를 받을 수 있음
            print("조회 데이터 분봉 수 %s" % cnt)
                    
            for i in range(cnt):
                data = [] # 종목별.. 1일마다 빈 리스트로 만들어짐..
                
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래량")
                trading_time = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결시간")
                start_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "시가")
                high_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "고가")
                low_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "저가")
                
                # data.append("")
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_time.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                # data.append("")
                
                self.calcul_data.append(data.copy()) # 종목별 1분씩 생성된 데이터를 self.calcul_data에 append 
                
            if sPrevNext == "2":
                print("==>> 다음 페이지 데이터 조회 ==>>\n")
                self.OPT50029_101T6000(sPrevNext = sPrevNext)
            else:
                print("Futures Options Code : [ 101T6000 ] 수집 완료 총 %s" % len(self.calcul_data))
                
                # 데이터 저장하는 부분
                path = './saved_options_time_data'
                
                if not os.path.exists(path):
                    os.makedirs(path)
                today = datetime.datetime.today().strftime('%Y%m%d')
                kiwoom_day = os.path.join(path, '{}_101T6000_data.pkl'.format(today))
                
                with open(kiwoom_day, 'wb') as f:
                    pickle.dump(self.calcul_data, f)
    
                self.calcul_data.clear() # 특정 종목 모든 일수 조회결과 저장 후 초기화
                self.calculator_event_loop.exit()
        #====================================================================================================================
        #====================================================================================================================
        elif sRQName == "선물옵션분차트요청_105T4000":
            print("105T4000 분봉데이터 요청") 
            
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName) # 1회 조회당 900틱까지 데이터를 받을 수 있음
            print("조회 데이터 분봉 수 %s" % cnt)
                    
            for i in range(cnt):
                data = [] # 종목별.. 1일마다 빈 리스트로 만들어짐..
                
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래량")
                trading_time = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결시간")
                start_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "시가")
                high_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "고가")
                low_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "저가")
                
                # data.append("")
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_time.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                # data.append("")
                
                self.calcul_data.append(data.copy()) # 종목별 1분씩 생성된 데이터를 self.calcul_data에 append 
                
            if sPrevNext == "2":
                print("==>> 다음 페이지 데이터 조회 ==>>\n")
                self.OPT50029_105T4000(sPrevNext = sPrevNext)
            else:
                print("Futures Options Code : [ 105T4000 ] 수집 완료 총 %s" % len(self.calcul_data))
                
                # 데이터 저장하는 부분
                path = './saved_options_time_data'
                
                if not os.path.exists(path):
                    os.makedirs(path)
                today = datetime.datetime.today().strftime('%Y%m%d')
                kiwoom_day = os.path.join(path, '{}_105T4000_data.pkl'.format(today)) 
                
                with open(kiwoom_day, 'wb') as f:
                    pickle.dump(self.calcul_data, f)
    
                self.calcul_data.clear() # 특정 종목 모든 일수 조회결과 저장 후 초기화
                self.calculator_event_loop.exit() 
    
    
    def OPT50029_101T6000(self, sPrevNext="0"): # 코스피 200 선물
        
        QTest.qWait(3600)
               
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", '101T6000')
        self.dynamicCall("SetInputValue(QString, QString)", "시간단위", "1")
            
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "선물옵션분차트요청_101T6000", "OPT50029", sPrevNext, "self.screen_calculation_stock")
        
        self.calculator_event_loop.exec_()
     
       
    def OPT50029_105T4000(self, sPrevNext="0"): # 코스피 200 선물
        
        QTest.qWait(3600)
               
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", '105T4000')
        self.dynamicCall("SetInputValue(QString, QString)", "시간단위", "1")
            
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "선물옵션분차트요청_105T4000", "OPT50029", sPrevNext, "self.screen_calculation_stock")
        
        self.calculator_event_loop.exec_()
        