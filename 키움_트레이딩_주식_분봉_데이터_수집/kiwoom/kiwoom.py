from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
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
        
        self.calculator_fnc() # 종목 분석용(임시용)

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
        elif sRQName == "주식분봉차트조회":
            code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "종목코드")
            code = code.strip()
            print("%s 분봉데이터 요청" % code) 
            
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName) # 1회 조회당 600틱까지 데이터를 받을 수 있음
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
                data.append(code)
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_time.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                # data.append("")
                
                self.calcul_data.append(data.copy()) # 종목별 1분씩 생성된 데이터를 self.calcul_data에 append 
                
            if sPrevNext == "2":
                self.day_kiwoom_db(code = code, sPrevNext = sPrevNext)
            else:
                print("Stock Code : [ %s ] 수집 완료 총 %s" % (code, len(self.calcul_data))) 

                # 데이터 저장하는 부분
                path = './saved_code_time_data'
                
                if not os.path.exists(path):
                    os.makedirs(path)
                kiwoom_day = os.path.join(path, f'{code}_data.pkl')
                
                with open(kiwoom_day, 'wb') as f:
                    pickle.dump(self.calcul_data, f)
                    
                print("==>> 다음 종목 데이터 조회 ==>>")
                self.calcul_data.clear() # 특정 종목 모든 일수 조회결과 저장 후 초기화
                self.calculator_event_loop.exit()
     
    
    def calculator_fnc(self): # 종목 분석 실행용 함수
        code_list = ['005930', '373220', '000660', '207940', '006400', '051910', '005935', '005380', '035420', '000270',
                     '005490', '035720', '068270', '028260', '012330', '105560', '003670', '066570', '055550', '096770',
                     '003550', '032830', '034730', '086790', '033780', '034020', '323410', '015760', '009150', '010130',
                     '017670', '000810', '011200', '051900', '018260', '010950', '036570', '003490', '329180', '316140',
                     '259960', '009830', '377300', '024110', '030200', '352820', '138040', '011170', '090430', '011070',
                     '086280', '028050', '402340', '005830', '383220', '302440', '034220', '009540', '088980', '271560',
                     '326030', '012450', '251270', '018880', '097950', '032640', '004020', '161390', '361610', '267250',
                     '010140', '011780', '008560', '047810', '006800', '000720', '035250', '241560', '000100', '011790',
                     '047050', '078930', '021240', '029780', '005387', '071050', '139480', '128940', '028670', '001450',
                     '307950', '282330', '004990', '002790', '112610', '005940', '180640', '007070', '008770', '005070']
        
        print("\n===== Kospi Stock Data ======")
        print("====  Data update start  ====\n")

        for idx, code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(Qstring)", self.screen_calculation_stock)
            
            print("==========================================================")
            print("%s / %s KOSPi TOP Stock Code : [ %s ] is updating ..." % (idx+1, len(code_list), code))
            print("==========================================================")
        
            self.day_kiwoom_db(code = code)
        
          
  
    def day_kiwoom_db(self, code = None, date = None, sPrevNext="0"):
        
        QTest.qWait(3600) #인위적으로 일정 시간을 기다림 --> 해당 설정을 해주지 않으면 과대조회로 api가 일시적으로 차단당함...ㅜ
        
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "틱범위", "1")
        
        if date != None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)
            
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "주식분봉차트조회", "opt10080", sPrevNext, self.screen_calculation_stock)
        self.calculator_event_loop.exec_()  