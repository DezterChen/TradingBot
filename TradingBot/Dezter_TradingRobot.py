# -*- coding: utf-8 -*-
"""
Created on Sat Jun 29 16:07:41 2019
Update on Sun Nov 24 2019 --> Fix: Switch day will have possibility to judge "No Open"
                          --> Order Function: Turn On
Update on Sun Feb 02 2020 --> Delete unused msg
@author: Dezter
Setup Trading Robot Notice:
    Lib need to install:
        comtypes(1.1.4), pip(9.0.1), pytz(2018.3), pywin32(222)
        Lib need to check install or not
        time, math, os, datetime, numpy, pandas
    file need to have and put in the same address:
        Des_TradingRobot.py, Des_CapitalAPI.py, Des_Transaction.py, Des_MsgInformTool.py
"Print" *1, "Error" clean
"""
#%% Initial Parameter Setting %%#
ErrorMsg = "No Error"; Command = []; FutureOrderBack = []; ErrorList = []
StockTSEABack=[]; StockTXBack=[]; FutureRightBack=[]; FutureOpenBack=[]; StockTodayBack=[]; StockListBack=[];
Login="Login ungo";EnterCall="EnterCall ungo";Ini="Ini ungo";ReadCert="ReadCert ungo";StockListCall="StockList ungo"
StockTSEACall="StockTSEACall ungo";StockTXCall="StockTXCall ungo";FutureRightCall="FutureRightCall ungo"
FutureOpenCall="FutureOpenCall ungo";FutureOrderCall="FutureOrderCall ungo"#;LeaveCall="LeaveCall ungo"

#%% import Lib %%#
try:
    import comtypes.client
    import time
    import datetime
except Exception as e:
    ErrorMsg = "Basic Lib import Error, "+ str(e)
StartTime = time.time()
#%% Capital API Activation %%#
try:
    import Des_CapitalAPI as API
    QuoteCall = API.QuoteCall()
    QuoteBack = API.QuoteBack()
    OrderCall = API.OrderCall()
    OrderBack = API.OrderBack()
    OrderForm = API.OrderForm()
    QuoteBackActive = comtypes.client.GetEvents(API.skQ, QuoteBack)
    OrderBackActive = comtypes.client.GetEvents(API.skO, OrderBack)
except Exception as e:
    ErrorMsg = "Des Capital API Error, "+ str(e)
#%% Transcation Model type A(Fall Point Variety with Stop Loss Once) Import%%#
try:
    import Des_TransactionModel as Model
except Exception as e:
    ErrorMsg = "Des Transcation Model Import Error, "+ str(e)
#%% MSG Inform Tool Import %%#
try:
    import Des_MsgInformTool as MsgInform
except Exception as e:
    ErrorMsg = "Des Message Inform Tool Import Error, "+ str(e)
#%% Trading Robot Parameter %%#
Id = "*********"; Pw = "*********"; Tsea = "TSEA"; Tx = "TX00"; Account = "*********"; ConnStatusBack = []
Scale = 4
StocksToday = Tsea + "," + Tx
#%% Trading Robot %%#
try:
    # Today Date
    TodayTime = datetime.datetime.today()
    Weekday = TodayTime.weekday() # Monday = 0; Tuesday = 1; Wednesday = 2; Thursday = 3; Friday = 4; Saturday = 5; Sunday = 6
    # Login, Connect Data Base, Order initialization, Cert check
    Login =API.Login(Id,Pw)
    EnterCall = QuoteCall.Enter()
    for t in range(0,5):
        API.pumpwait2(4)
        ConnStatusBack = QuoteBack.ConnResult
        if ConnStatusBack[-1] == "Stocks ready!":
            break
        else:
            continue
    Ini = OrderCall.Initialize()
    ReadCert = OrderCall.ReadCert(Id)
    
    # Stock History Data - Call TSEA data & MTX/TX data
    QuoteBack.StockData = []
    StockTSEACall = QuoteCall.StockData(Tsea,Scale,1)
    for t in range(0,20):
        API.pumpwait2(1)
        StockTSEABack = QuoteBack.StockData
        if StockTSEABack != []:
            break
        else:
            continue    
    QuoteBack.StockData = []
    StockTXCall = QuoteCall.StockData(Tx,Scale,1)
    for t in range(0,20):
        API.pumpwait2(1)
        StockTXBack = QuoteBack.StockData
        if StockTXBack != []:
            break
        else:
            continue
    # Call Future Right Report/Open Interest
    """ --> FutureRightCall/FutureRightBack is real time report """
    FutureRightCall = OrderCall.FutureRight(Id,Account)
    FutureOpenCall = OrderCall.FutureOpen(Id,Account)
    # Today Data - Call Current Stock Price
    StockTodayCall = QuoteCall.StockPrice(StocksToday)
    # Call Stock List
    StockListCall = QuoteCall.StockList(2)
    
    # Data Back
    API.pumpwait2(8)
    FutureRightBack = OrderBack.FutureRight
    FutureOpenBack = OrderBack.FutureOpen
    StockTodayBack = QuoteBack.StockPrice
    #ConnStatusBack = QuoteBack.ConnResult
    for t in range(0,20):
        API.pumpwait2(1)
        StockListBack = QuoteBack.StockList
        if StockListBack != []:
            break
        else:
            continue

    # Opening Check --> need to fix the problem: switch day will judge to no open day 
    Opendaycheck = len(StockTodayBack)
    #StockTodayBack = QuoteBack.StockPrice
    API.pumpwait2(10) # need to check while is not opening day
    Opendaycheck1 = len(StockTodayBack)
    
    DueDayInfo = StockListBack[0][0].split(',')
    DueDay = datetime.datetime.strptime(DueDayInfo[-1],'%Y%m%d').date()
    if DueDay == TodayTime.date():
        Opencheck = "Open"
    elif Opendaycheck1 != Opendaycheck:
        Opencheck = "Open"
    elif Opendaycheck1 == Opendaycheck:
        Opencheck = "No Open"
    # Today data back test
    """
    StockTodayCalltest = QuoteCall.StockPrice(StocksToday)
    API.pumpwait2(8)
    StockTodayBacktest = QuoteBack.StockPrice
    print(StockTodayBacktest)    
    time.sleep(30)
    StockTodayBacktest = QuoteBack.StockPrice
    print(StockTodayBacktest)
    """
    # Transcation Modal - import Today Date, Today Data, History Data, Future Right Report(Property), Open Interest Report(B or S check)
    Command = Model.TransactionModel(Opencheck,TodayTime,Tsea,StockTSEABack,Tx,StockTXBack,StockTodayBack,FutureOpenBack,FutureRightBack,StockListBack,1500,Scale)
    
    # Execution Transcation Model Command
    Execution = Command[2]
    #Execution[0][0] = "Keep" #--> test use
    #FutureOrderCall = "System test, Inform Only" #--> test use
    if Execution[0][0] == "Keep" or Execution[0][0] == "NoSwitch":
        FutureOrderBack = ["Order: None"]
    elif Execution[0][0] == "Order" or Execution[0][0] == "Offset" or Execution[0][0] == "Switch":
        for i in range(len(Execution)):
            if Execution[i][0] == "Offset":
                FutureOrderForm = OrderForm.Future('*********',Execution[i][1],Execution[i][2],"M",Execution[i][3],'IOC',1)# reserve "1", instant trading "0"
                FutureOrderCall = OrderCall.FutureOrder(Id,True,FutureOrderForm) #--> Error Msg Return Setting
            elif Execution[i][0] == "Order":
                FutureOrderForm = OrderForm.Future('*********',Execution[i][1],Execution[i][2],"M",Execution[i][3],'IOC',1)# reserve "1", instant trading "0"
                FutureOrderCall = OrderCall.FutureOrder(Id,True,FutureOrderForm) #--> Error Msg Return Setting
            elif Execution[i][0] == "Switch":
                FutureOrderForm = OrderForm.Future('*********',Execution[i][1],Execution[i][2],"M",Execution[i][3],'IOC',1)# reserve "1", instant trading "0"
                FutureOrderCall = OrderCall.FutureOrder(Id,True,FutureOrderForm) #--> Error Msg Return Setting
        for t in range(0,20):
            API.pumpwait2(1)
            FutureOrderBack = OrderBack.OrderFeedback
            if FutureOrderBack != []:
                break
            else:
                continue
except Exception as e:
    ErrorMsg = "Des Trading Robot Error, "+ str(e)
    
#%% API Connecting Error Collection %%#
try:
    CallError = [Login,EnterCall,Ini,ReadCert,StockTSEACall,StockTXCall,FutureRightCall,FutureOpenCall,FutureOrderCall,ConnStatusBack[-1]]
    for i in range(len(CallError)):
        if "Success" in CallError[i]:
            pass
        else:
            ErrorList.append(CallError[i])
    BackError = [StockTSEABack,StockTXBack,FutureRightBack,FutureOpenBack,StockTodayBack]
    BackErrorStr = ['StockTSEABack','StockTXBack','FutureRightBack','FutureOpenBack','StockTodayBack']
    for i in range(len(BackErrorStr)):
        if len(BackError[i]) == 0:
            BackErrorMsg = BackErrorStr[i] + " no data"
            ErrorList.append(BackErrorMsg)
        else:
            pass
except Exception as e:
    ErrorMsg = "Error Collection Error, "+ str(e)
if ErrorMsg == "No Error":
    if ErrorList == []:
        ErrorMsg = "No Error"
    else:
        ErrorMsg = str(ErrorList)
else:
    ErrorMsg = ErrorMsg + '\n'+ str(ErrorList)

#%% Message/Mail Inform %%#
MsgInform.InformTool(Command,FutureOrderBack,ErrorMsg) #--> test use (block)

EndTime = time.time()
DeltaTime = EndTime - StartTime
print(DeltaTime)
