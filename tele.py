
# https://py-son.tistory.com/8
import win32com.client
 
# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()
 
# 현재가 객체 구하기
objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
objStockMst.SetInputValue(0, 'A005930')   #종목 코드 - 삼성전자
objStockMst.BlockRequest() # 삼성전자 현재가 가져옴
 
# 현재가 통신 및 통신 에러 처리 
rqStatus = objStockMst.GetDibStatus()
rqRet = objStockMst.GetDibMsg1()
print("통신상태", rqStatus, rqRet)
if rqStatus != 0:
    exit()
 
# 현재가 정보 조회
# code = objStockMst.GetHeaderValue(0)  #종목코드
# name= objStockMst.GetHeaderValue(1)  # 종목명
# time= objStockMst.GetHeaderValue(4)  # 시간
# cprice= objStockMst.GetHeaderValue(11) # 종가
# diff= objStockMst.GetHeaderValue(12)  # 대비
# open= objStockMst.GetHeaderValue(13)  # 시가
# high= objStockMst.GetHeaderValue(14)  # 고가
# low= objStockMst.GetHeaderValue(15)   # 저가
offer = objStockMst.GetHeaderValue(16)  #매도호가
# bid = objStockMst.GetHeaderValue(17)   #매수호가
# vol= objStockMst.GetHeaderValue(18)   #거래량
# vol_value= objStockMst.GetHeaderValue(19)  #거래대금
 
# 예상 체결관련 정보
exFlag = objStockMst.GetHeaderValue(58) #예상체결가 구분 플래그
exPrice = objStockMst.GetHeaderValue(55) #예상체결가
exDiff = objStockMst.GetHeaderValue(56) #예상체결가 전일대비
exVol = objStockMst.GetHeaderValue(57) #예상체결수량
 
 
# print("코드", code)
# print("이름", name)
# print("시간", time)
# print("종가", cprice)
# print("대비", diff)
# print("시가", open)
# print("고가", high)
# print("저가", low)
print("매도호가", offer)
# print("매수호가", bid)
# print("거래량", vol)
# print("거래대금", vol_value)
 
 
if (exFlag == ord('0')):
    print("장 구분값: 동시호가와 장중 이외의 시간")
elif (exFlag == ord('1')) :
    print("장 구분값: 동시호가 시간")
elif (exFlag == ord('2')):
    print("장 구분값: 장중 또는 장종료")
 
print("예상체결가 대비 수량")
print("예상체결가", exPrice)
print("예상체결가 대비", exDiff)
print("예상체결수량", exVol)



import telegram
from telegram.ext import Updater
from telegram.ext import MessageHandler, Filters
 
token = '5531235803:AAFk6WLU7lltNUr8xkxymhHwZpR2-K9p338'
id='5347123098'
bot = telegram.Bot(token=token)
bot.sendMessage(chat_id=id,text='삼성전자 : '+str(offer))

# updater
updater = Updater(token=token, use_context=True)
dispatcher = updater.dispatcher
updater.start_polling()

# 사용자가 보낸 메세지를 읽어들이고, 답장을 보내줍니다.
# 아래 함수만 입맛에 맞게 수정해주면 됩니다. 다른 것은 건들 필요없어요.
def handler(update, context):
    user_text = update.message.text # 사용자가 보낸 메세지를 user_text 변수에 저장합니다.
    
    if user_text == "안녕": # 사용자가 보낸 메세지가 "안녕"이면?
        bot.send_message(chat_id=id, text="어 그래 안녕") # 답장 보내기
    elif user_text == "뭐해": # 사용자가 보낸 메세지가 "뭐해"면?
        bot.send_message(chat_id=id, text="그냥 있어") # 답장 보내기
 
echo_handler = MessageHandler(Filters.text, handler)
dispatcher.add_handler(echo_handler)
