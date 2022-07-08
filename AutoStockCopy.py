# https://py-son.tistory.com/8 텔레그램파이썬
from subprocess import list2cmdline
import win32com.client
 
 
# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos") #win32com을통해 CybosPlus연결
bConnect = objCpCybos.IsConnect
if (bConnect == 0): # 크레온플러스 실행시켜야 연결 됨
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()
 
# 종목코드 리스트 구하기
objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
codeList = objCpCodeMgr.GetStockListByMarket(1) #거래소 codelist불러옴
codeList2 = objCpCodeMgr.GetStockListByMarket(2) #코스닥 codelist불러옴
 
 
print("거래소 종목코드", len(codeList))
for i, code in enumerate(codeList):
    secondCode = objCpCodeMgr.GetStockSectionKind(code)
    name = objCpCodeMgr.CodeToName(code)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)
    print(i, code, secondCode, stdPrice, name)
 
print("코스닥 종목코드", len(codeList2))
for i, code in enumerate(codeList2):
    secondCode = objCpCodeMgr.GetStockSectionKind(code)
    name = objCpCodeMgr.CodeToName(code)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)
    print(i, code, secondCode, stdPrice, name)
 
print("거래소 + 코스닥 종목코드 ",len(codeList) + len(codeList2))


import telegram
from telegram.ext import Updater
from telegram.ext import MessageHandler, Filters
 
token = '5531235803:AAFk6WLU7lltNUr8xkxymhHwZpR2-K9p338'
id='5347123098'
bot = telegram.Bot(token=token)
bot.sendMessage(chat_id=id,text='어서오십시오, 주식 자동매매를 시작합니다. ')

# updater
updater = Updater(token=token, use_context=True)
dispatcher = updater.dispatcher
updater.start_polling()




# 사용자가 보낸 메세지를 읽어들이고, 답장을 보내줍니다.
# 아래 함수만 입맛에 맞게 수정해주면 됩니다. 다른 것은 건들 필요없어요.

def handler(update, context):
    user_text = update.message.text # 사용자가 보낸 메세지를 user_text 변수에 저장합니다.
    
    
    
    ################ if user_text == "안녕": # 사용자가 보낸 메세지가 "안녕"이면?

    # if user_text =='구매':
    #         bot.send_message(chat_id=id, text='구매하고 싶은 종목이름을 입력하여 종목코드를 복사한 후 종목코드, 구매하고 싶은 수량을 "ex),50"형식으로 적어주세요.') # 답장 보내기
    #         global buyList
    #         buyList=user_text
    #         Lists=buyList.split(',')
    #         global list1
    #         list1 = buyList[0]
    #         global list2
    #         list2 = buyList[1]
    

    for i, code in enumerate(codeList):
        secondCode = objCpCodeMgr.GetStockSectionKind(code)
        name = objCpCodeMgr.CodeToName(code)
        stdPrice = objCpCodeMgr.GetStockStdPrice(code)

        if user_text in name:
            bot.send_message(chat_id=id, text=str(name)+"의 현재가(List1) : "+str(stdPrice)) # 답장 보내기  
            global saveCode
            saveCode=code[0:7]
            bot.send_message(chat_id=id, text=" "+str(name)+'의 종목 코드명 : '+saveCode) # 답장 보내기


    
    for i, code in enumerate(codeList2):
        secondCode = objCpCodeMgr.GetStockSectionKind(code)
        name = objCpCodeMgr.CodeToName(code)
        stdPrice = objCpCodeMgr.GetStockStdPrice(code)
        
        if user_text in name:
            bot.send_message(chat_id=id, text=str(name)+"의 현재가(List2) : "+str(stdPrice)) # 답장 보내기
            
            saveCode=code[0:7]                # code = 코드번호 + 가격 A0000003456 그래서 7자리까지
            bot.send_message(chat_id=id, text=" "+str(name)+'의 종목 코드명 : '+saveCode) # 답장 보내기
        #     bot.send_message(chat_id=id, text='해당 종목 매수 원하시면 "매수"를 입력해주세요.')
        # if user_text == "매수": # 사용자가 보낸 메세지가 "매수"면?
        #     bot.send_message(chat_id=id, text=str(name)+'종목 : '+saveCode+'매수합니다.') # 답장 보내기
        #     buy_etf(saveCode)
        #     print("buy_etf 체크")

        if 'A' in user_text and ',' in user_text :
            
            buyList=user_text
            lists=buyList.split(',')
            
            list1 = lists[0]
            
            list2 = lists[1]
            print(list1,list2+'리스트확인')
            buy_stoo(list1,list2)
            
            break


            # global codeA
            # codeA=user_text
            # bot.send_message(chat_id=id, text='해당'+codeA+'매수합니다.')
            # print(codeA)
            # break
        if user_text == "매수": # 사용자가 보낸 메세지가 "매수"면?
                bot.send_message(chat_id=id, text='구매하고 싶은 종목이름을 입력하여 종목코드를 복사한 후 종목코드, 구매하고 싶은 수량을 "ex)A066910,37"형식으로 적어주세요.') # 답장 보내기
                
                
                break
            

    
        
        
 
echo_handler = MessageHandler(Filters.text, handler)
dispatcher.add_handler(echo_handler)







import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime
# from slacker import Slacker
import time, calendar
 

# slack = Slacker('xoxb-3738243281616-3707921255558-QAcokDsK5uMFJ6JHMJFvMZ50')
def dbgout(message):
    """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    bot.send_message(chat_id=id, text="strbuf 99번라인"+strbuf)  #텔레그램 전송
    # slack.chat.post_message('#stock', strbuf)  #메세지 보내는 채널

def printlog(message, *args):
    """인자로 받은 문자열을 파이썬 셸에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)
 
# 크레온 플러스 공통 OBJECT
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')  

def check_creon_system():
    """크레온 플러스 시스템 연결 상태를 점검한다."""
    # 관리자 권한으로 프로세스 실행 여부
    if not ctypes.windll.shell32.IsUserAnAdmin():
        printlog('check_creon_system() : admin user -> FAILED')
        return False
 
    # 연결 여부 체크
    if (cpStatus.IsConnect == 0):
        printlog('check_creon_system() : connect to server -> FAILED')
        return False
 
    # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    if (cpTradeUtil.TradeInit(0) != 0):
        printlog('check_creon_system() : init trade -> FAILED')
        return False
    return True

def get_current_price(code):
    """인자로 받은 종목의 현재가, 매도호가, 매수호가를 반환한다."""
    cpStock.SetInputValue(0, code)  # 종목코드에 대한 가격 정보
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)   # 현재가
    item['ask'] =  cpStock.GetHeaderValue(16)        # 매도호가
    item['bid'] =  cpStock.GetHeaderValue(17)        # 매수호가    
    return item['cur_price'], item['ask'], item['bid']

def get_ohlc(code, qty):
    """인자로 받은 종목의 OHLC 가격 정보를 qty 개수만큼 반환한다."""
    cpOhlc.SetInputValue(0, code)           # 종목코드
    cpOhlc.SetInputValue(1, ord('2'))        # 1:기간, 2:개수
    cpOhlc.SetInputValue(4, qty)             # 요청개수
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5]) # 0:날짜, 2~5:OHLC
    cpOhlc.SetInputValue(6, ord('D'))        # D:일단위
    cpOhlc.SetInputValue(9, ord('1'))        # 0:무수정주가, 1:수정주가
    cpOhlc.BlockRequest()
    count = cpOhlc.GetHeaderValue(3)   # 3:수신개수
    columns = ['open', 'high', 'low', 'close']
    index = []
    rows = []
    for i in range(count): 
        index.append(cpOhlc.GetDataValue(0, i)) 
        rows.append([cpOhlc.GetDataValue(1, i), cpOhlc.GetDataValue(2, i),
            cpOhlc.GetDataValue(3, i), cpOhlc.GetDataValue(4, i)]) 
    df = pd.DataFrame(rows, columns=columns, index=index) 
    return df

def get_stock_balance(code):
    """인자로 받은 종목의 종목명과 수량을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpBalance.SetInputValue(0, acc)         # 계좌번호
    cpBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpBalance.SetInputValue(2, 50)          # 요청 건수(최대 50)
    cpBalance.BlockRequest()     
    if code == 'ALL':
        dbgout('계좌명: ' + str(cpBalance.GetHeaderValue(0)))
        dbgout('결제잔고수량 : ' + str(cpBalance.GetHeaderValue(1)))
        dbgout('평가금액: ' + str(cpBalance.GetHeaderValue(3)))
        dbgout('평가손익: ' + str(cpBalance.GetHeaderValue(4)))
        dbgout('종목수: ' + str(cpBalance.GetHeaderValue(7)))
    stocks = []
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)  # 종목코드
        stock_name = cpBalance.GetDataValue(0, i)   # 종목명
        stock_qty = cpBalance.GetDataValue(15, i)   # 수량
        if code == 'ALL':
            dbgout(str(i+1) + ' ' + stock_code + '(' + stock_name + ')' 
                + ':' + str(stock_qty))
            stocks.append({'code': stock_code, 'name': stock_name, 
                'qty': stock_qty})
        if stock_code == code:  
            return stock_name, stock_qty
    if code == 'ALL':
        return stocks
    else:
        stock_name = cpCodeMgr.CodeToName(code)
        return stock_name, 0

def get_current_cash():
    """증거금 100% 주문 가능 금액을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]    # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpCash.SetInputValue(0, acc)              # 계좌번호
    cpCash.SetInputValue(1, accFlag[0])      # 상품구분 - 주식 상품 중 첫번째
    cpCash.BlockRequest() 
    return cpCash.GetHeaderValue(9) # 증거금 100% 주문 가능 금액

def get_target_price(code):
    """매수 목표가를 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 10)
        if str_today == str(ohlc.iloc[0].name):
            today_open = ohlc.iloc[0].open 
            lastday = ohlc.iloc[1]
        else:
            lastday = ohlc.iloc[0]                                      
            today_open = lastday[3]
        lastday_high = lastday[1]
        lastday_low = lastday[2]

        target_price = today_open + (lastday_high - lastday_low) * 0.5
        return target_price
    except Exception as ex:
        dbgout("`get_target_price() -> exception! " + str(ex) + "`")
        return None
    
def get_movingaverage(code, window):
    """인자로 받은 종목에 대한 이동평균가격을 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 20)
        if str_today == str(ohlc.iloc[0].name):
            lastday = ohlc.iloc[1].name
        else:
            lastday = ohlc.iloc[0].name
        closes = ohlc['close'].sort_index()         
        ma = closes.rolling(window=window).mean()
        return ma.loc[lastday]
    except Exception as ex:
        dbgout('get_movingavrg(' + str(window) + ') -> exception! ' + str(ex))
        return None    

#################################
def buy_stoo(code,count):
    time_now=datetime.now()
    printlog(' meets the buy condition!`')
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션                
    # 최유리 FOK 매수 주문 설정
    cpOrder.SetInputValue(0, "2")        # 2: 매수                          #최유리 : 당장 가장 유리하게 매매할 수 있는 가격,
    cpOrder.SetInputValue(1, acc)        # 계좌번호                         #최우선 : 우선 대기하는 가격
    cpOrder.SetInputValue(2, accFlag[0]) # 상품구분 - 주식 상품 중 첫번째   #IOC거래방식 - 체결 후 남은 수량 취소 (5000개 구매원하면 가능한 갯수만큼만 구매)
    cpOrder.SetInputValue(3, code)       # 종목코드                         #FOK거래방식 - 전량 체결되지 않으면 주문 자체를 취소(5000개구매원하면 5000개구매 or 0)
    cpOrder.SetInputValue(4, count)    # 매수할 수량
    cpOrder.SetInputValue(7, "2")        # 주문조건 0:기본, 1:IOC, 2:FOK
    cpOrder.SetInputValue(8, "12")       # 주문호가 1:보통, 3:시장가
                                                 # 5:조건부, 12:최유리, 13:최우선 
    # 매수 주문 요청
    ret = cpOrder.BlockRequest() 

    ##########################################################

def buy_etf(code):
    """인자로 받은 종목을 최유리 지정가 FOK 조건으로 매수한다."""
    try:
        global bought_list      # 함수 내에서 값 변경을 하기 위해 global로 지정
        if code in bought_list: # 매수 완료 종목이면 더 이상 안 사도록 함수 종료
            #printlog('code:', code, 'in', bought_list)
            return False
        time_now = datetime.now()
        current_price, ask_price, bid_price = get_current_price(code)  #현재가격 구해줌
        target_price = get_target_price(code)    # 매수 목표가 @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        ma5_price = get_movingaverage(code, 5)   # 5일 이동평균가    주가의 이동 평균을 구해서 평균값을 이은 선
        ma10_price = get_movingaverage(code, 10) # 10일 이동평균가   둘다 위에 있는지 검사
        buy_qty = 0        # 매수할 수량 초기화
        if ask_price > 0:  # 매도호가가 존재하면   
            buy_qty = buy_amount // ask_price   # //은 몫

    except Exception as ex:
        dbgout("`buy_etf("+ str(code) + ") -> exception! " + str(ex) + "`")

############################################################ 매수들어갈때 물어보고 매수하기? 
    
        

############################################################

        stock_name, stock_qty = get_stock_balance(code)  # 종목명과 보유수량 조회
        #printlog('bought_list:', bought_list, 'len(bought_list):',
        #    len(bought_list), 'target_buy_count:', target_buy_count)     
        if current_price > target_price and current_price > ma5_price \
            and current_price > ma10_price:         #위의 조건이 일치하면,(변동성 돌파전략으로 구해준 타겟보다 높은지,이동평균가보다 높고) 매수
            printlog(stock_name + '(' + str(code) + ') ' + str(buy_qty) +
                'EA : ' + str(current_price) + ' meets the buy condition!`')            
            cpTradeUtil.TradeInit()
            acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
            accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션                
            # 최유리 FOK 매수 주문 설정
            cpOrder.SetInputValue(0, "2")        # 2: 매수                          #최유리 : 당장 가장 유리하게 매매할 수 있는 가격,
            cpOrder.SetInputValue(1, acc)        # 계좌번호                         #최우선 : 우선 대기하는 가격
            cpOrder.SetInputValue(2, accFlag[0]) # 상품구분 - 주식 상품 중 첫번째   #IOC거래방식 - 체결 후 남은 수량 취소 (5000개 구매원하면 가능한 갯수만큼만 구매)
            cpOrder.SetInputValue(3, code)       # 종목코드                         #FOK거래방식 - 전량 체결되지 않으면 주문 자체를 취소(5000개구매원하면 5000개구매 or 0)
            cpOrder.SetInputValue(4, buy_qty)    # 매수할 수량
            cpOrder.SetInputValue(7, "2")        # 주문조건 0:기본, 1:IOC, 2:FOK
            cpOrder.SetInputValue(8, "12")       # 주문호가 1:보통, 3:시장가
                                                 # 5:조건부, 12:최유리, 13:최우선 
            # 매수 주문 요청
            ret = cpOrder.BlockRequest() 
            printlog('최유리 FoK 매수 ->', stock_name, code, buy_qty, '->', ret)
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time/1000)
                time.sleep(remain_time/1000) 
                return False
            time.sleep(2)
            printlog('현금주문 가능금액 :', buy_amount)
            stock_name, bought_qty = get_stock_balance(code)
            printlog('get_stock_balance :', stock_name, stock_qty)
            if bought_qty > 0:
                bought_list.append(code)
                dbgout("`buy_etf("+ str(stock_name) + ' : ' + str(code) + 
                    ") -> " + str(bought_qty) + "EA bought!" + "`")         #현재 어떤 주식을 얼마나 샀는지 텔레그램(슬랙)으로 보냄
    except Exception as ex:
        dbgout("`buy_etf("+ str(code) + ") -> exception! " + str(ex) + "`")

def sell_all():
    """보유한 모든 종목을 최유리 지정가 IOC 조건으로 매도한다."""   #@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    try:
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]       # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션   
        while True:    
            stocks = get_stock_balance('ALL') 
            total_qty = 0 
            for s in stocks:
                total_qty += s['qty'] 
            if total_qty == 0:
                return True
            for s in stocks:
                if s['qty'] != 0:                  
                    cpOrder.SetInputValue(0, "1")         # 1:매도, 2:매수
                    cpOrder.SetInputValue(1, acc)         # 계좌번호
                    cpOrder.SetInputValue(2, accFlag[0])  # 주식상품 중 첫번째
                    cpOrder.SetInputValue(3, s['code'])   # 종목코드
                    cpOrder.SetInputValue(4, s['qty'])    # 매도수량
                    cpOrder.SetInputValue(7, "1")   # 조건 0:기본, 1:IOC, 2:FOK
                    cpOrder.SetInputValue(8, "12")  # 호가 12:최유리, 13:최우선 
                    # 최유리 IOC 매도 주문 요청
                    ret = cpOrder.BlockRequest()
                    printlog('최유리 IOC 매도', s['code'], s['name'], s['qty'], 
                        '-> cpOrder.BlockRequest() -> returned', ret)
                    if ret == 4:
                        remain_time = cpStatus.LimitRequestRemainTime
                        printlog('주의: 연속 주문 제한, 대기시간:', remain_time/1000)
                time.sleep(1)
            time.sleep(30)
    except Exception as ex:
        dbgout("sell_all() -> exception! " + str(ex))

if __name__ == '__main__': 
    try:
        symbol_list = ['A027710', 'A008970', 'A001440', 'A208710']   # ETF 종목 코드 입력 #A027710 팜스토리 A008970동양철관 A001440대한전선  A208710바이오로그디바이스
        bought_list = []     # 매수 완료된 종목 리스트
        target_buy_count = 4 # 매수할 종목 수 리스트 중 최대 몇 종목 까지 매수?
        buy_percent = 0.25     #전체 가용 금액에서 몇 퍼센트를 살것인지
        printlog('check_creon_system() :', check_creon_system())  # 크레온 접속 점검
        stocks = get_stock_balance('ALL')      # 보유한 모든 종목 조회
        total_cash = int(get_current_cash())   # 100% 증거금 주문 가능 금액 조회
        buy_amount = total_cash * buy_percent  # 종목별 주문 금액 계산
        printlog('100% 증거금 주문 가능 금액 :', total_cash)
        printlog('종목별 주문 비율 :', buy_percent)
        printlog('종목별 주문 금액 :', buy_amount)
        printlog('시작 시간 :', datetime.now().strftime('%m/%d %H:%M:%S'))
        soldout = False

        while True:
            t_now = datetime.now() #현재시간 저장
            t_9 = t_now.replace(hour=9, minute=0, second=0, microsecond=0)         #주식 시장 정규시간09:00~15:30
            t_start = t_now.replace(hour=9, minute=5, second=0, microsecond=0)   #LP(유동성 공급자)활동시간09:05~15:20
            t_sell = t_now.replace(hour=15, minute=10, second=0, microsecond=0)  #팔기 시작하는 시간 #자동매매시간 09:05~15:15 / 15:20 (종료)
            t_exit = t_now.replace(hour=15, minute=20, second=0,microsecond=0)  #프로그램종료
            today = datetime.today().weekday()
            if today == 5 or today == 6:  # 토요일이나 일요일이면 자동 종료
                printlog('Today is', 'Saturday.' if today == 5 else 'Sunday.')  
                sys.exit(0)
            if t_9 < t_now < t_start and soldout == False: #혹시 남아있는 주식이 있으면 09시 ~09시05분에 팔기
                soldout = True
                sell_all()
            if t_start < t_now < t_sell :  # AM 09:05 ~ PM 03:15 : 매수
                for sym in symbol_list:         #자동매매를 위해 고른 종목 코드 후보들
                    if len(bought_list) < target_buy_count:  # 목표한 갯수 다 구매했는지 검사
                        buy_etf(sym)                        #다 구매 안했으면 조건 검사 후 매수
                        time.sleep(1)
                if t_now.minute == 30 and 0 <= t_now.second <= 5: #30분마다 현재 잔고 알림
                    get_stock_balance('ALL')
                    time.sleep(5)
            if t_sell < t_now < t_exit:  # PM 03:15 ~ PM 03:20 : 일괄 매도
                if sell_all() == True:
                    dbgout('`sell_all() returned True -> self-destructed!`')
                    sys.exit(0)
            if t_exit < t_now:  # PM 03:20 ~ :프로그램 종료
                dbgout('`self-destructed!`')
                sys.exit(0)
            time.sleep(3)
    except Exception as ex:
        dbgout('`main -> exception! ' + str(ex) + '`')

