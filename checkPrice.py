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


    for i, code in enumerate(codeList):
        secondCode = objCpCodeMgr.GetStockSectionKind(code)
        name = objCpCodeMgr.CodeToName(code)
        stdPrice = objCpCodeMgr.GetStockStdPrice(code)
        if user_text in name:
            bot.send_message(chat_id=id, text="현재가 : "+str(stdPrice)) # 답장 보내기


    
    for i, code in enumerate(codeList2):
        secondCode = objCpCodeMgr.GetStockSectionKind(code)
        name = objCpCodeMgr.CodeToName(code)
        stdPrice = objCpCodeMgr.GetStockStdPrice(code)
        if user_text in name:
            bot.send_message(chat_id=id, text="현재가 : "+str(stdPrice)) # 답장 보내기



    if user_text == "뭐햄": # 사용자가 보낸 메세지가 "뭐해"면?
        bot.send_message(chat_id=id, text="그냥 있어") # 답장 보내기
 
echo_handler = MessageHandler(Filters.text, handler)
dispatcher.add_handler(echo_handler)