# -*- coding: utf-8 -*-
"""
Created on Sun Sep 22 18:25:59 2019

@author: Dezter
"""
#%% import Lib %%#
import comtypes.client
import time

#%% Capital API Activation %%#
import Des_CapitalAPI as API
QuoteCall = API.QuoteCall()
QuoteBack = API.QuoteBack()
OrderCall = API.OrderCall()
OrderBack = API.OrderBack()
OrderForm = API.OrderForm()
QuoteBackActive = comtypes.client.GetEvents(API.skQ, QuoteBack)
OrderBackActive = comtypes.client.GetEvents(API.skO, OrderBack)

#%% Capital API Function Check %%#
RunAPI = 11111
if RunAPI == 1:
    """ Login """
    Id = "*********"; Pw = "*********" # --> Privacy Info
    Login =API.Login(Id,Pw)
    """ Enter Database """
    EnterCall = QuoteCall.Enter()
    API.pumpwait2(8)
    EnterBack = QuoteBack.ConnResult
    """ Leave Database """
    #LeaveCall = QuoteCall.Leave()
    #LeaveBack = QuoteBack.ConnResult
    """ Call Stock History Data """
    QuoteBack.StockData = []
    StockTSEACall = QuoteCall.StockData("TSEA",4,1)
    StockTSEABack = QuoteBack.StockData
    QuoteBack.StockData = []
    StockMTXCall = QuoteCall.StockData("MTX00",4,1)
    StockMTXBack = QuoteBack.StockData
    """ Call Stock Current Price Data """
    StockTSEATodayCall = QuoteCall.StockPrice("TSEA")
    StockTSEATodayBack = QuoteBack.StockPrice
    time.sleep(1)
    LeaveCall = QuoteCall.Leave()
    LeaveBack = QuoteBack.ConnResult
    time.sleep(1)
    EnterCall = QuoteCall.Enter()
    API.pumpwait2(8)
    EnterBack = QuoteBack.ConnResult
    """ Call Stock List """
    StockListCall = QuoteCall.StockList(2)
    StockListBack = QuoteBack.StockList
    """ Initialize """
    Ini = OrderCall.Initialize()
    """ Read Cert """
    ReadCert = OrderCall.ReadCert(Id)
    """ Get Order Account """
    AccountCall = OrderCall.AccountGet()
    #API.pumpwait2(8)
    #AccountBack = OrderBack.AccountGet
    """ Futurn Open Interest Report """
    FutureOpenCall = OrderCall.FutureOpen(Id,"*********") # --> Privacy Info
    #API.pumpwait2(8)
    #FutureOpenBack = OrderBack.FutureOpen
    """ Futurn Right Report """
    FutureRightCall = OrderCall.FutureRight(Id,"*********") # --> Privacy Info
    #API.pumpwait2(8)
    #FutureRightBack = OrderBack.FutureRight
    """ Back return by one time """
    API.pumpwait2(8)
    AccountBack = OrderBack.AccountGet
    FutureOpenBack = OrderBack.FutureOpen
    FutureRightBack = OrderBack.FutureRight
    """ Futurn Order """ # Msg return by SKCenterLib_GetLastLogInfo
    #EX. FutureOrderForm = OrderForm.Future("account","stockNo","B/S","price",qty,"ROD/IOC/FOK",盤中 = 0、T盤預約 = 1)
    """
    FutureOrderForm = OrderForm.Future('*********',"name","Sell","M",1,'IOC',1) # --> Privacy Info
    FutureOrderCall = OrderCall.FutureOrder(Id,False,FutureOrderForm)
    API.pumpwait2(8)
    FutureOrderBack = OrderBack.OrderFeedback
    """
    """ Future Multiple Order """ # Msg return by OnAsyncOrder
    """
    FutureOrderCall = OrderForm.Future('*********',"name","Sell","M",1,'IOC',1) # --> Privacy Info
    FutureOrderCall = OrderCall.FutureOrder(Id,True,FutureOrderCall)
    FutureOrderForm = OrderForm.Future('*********',"name","Sell","M",2,'IOC',1) # --> Privacy Info
    FutureOrderCall = OrderCall.FutureOrder(Id,True,FutureOrderForm)
    FutureOrderForm = OrderForm.Future('*********',"neme","Sell","M",3,'IOC',1) # --> Privacy Info
    FutureOrderCall = OrderCall.FutureOrder(Id,True,FutureOrderForm)
    API.pumpwait2(8)
    FutureOrderBack = OrderBack.OrderFeedback
    """
    """ Stock Profit Report """
    StockProfitCall = OrderCall.StockProfit(Id,"*********") # --> Privacy Info
    API.pumpwait2(8)
    StockProfitBack = OrderBack.StockProfit
    """ Stock Balence Report """
    StockBalanceCall = OrderCall.StockBalance(Id,"*********") # --> Privacy Info
    API.pumpwait2(8)
    StockBalanceBack = OrderBack.StockBalance
    """ Stock Order """ # Msg return by SKCenterLib_GetLastLogInfo
    # EX. StockOrderForm = OrderForm.Stock("account","stockNo","B/S","price",qty)
    """
    StockOrderForm = OrderForm.Stock('*********',"2330","Buy","M",1) # --> Privacy Info
    StockeOrderCall = OrderCall.StockOrder(Id,False,StockOrderForm)
    API.pumpwait2(8)
    StockOrderBack = OrderBack.OrderFeedback
    """
    """ Stock Multiple Order """ # Msg return by OnAsyncOrder
    """
    StockOrderForm = OrderForm.Stock('*********',"2330","Buy","M",1) # --> Privacy Info
    StockeOrderCall = OrderCall.StockOrder(Id,True,StockOrderForm)
    API.pumpwait2(8)
    StockOrderBack = OrderBack.OrderFeedback
    """
else:
    pass