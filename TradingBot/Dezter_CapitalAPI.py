# -*- coding: utf-8 -*-
"""
Created on Sat Jun 29 16:59:20 2019

@author: Dezter
"""
#%% import Lib %%#
import os
import comtypes.client
import pythoncom
import time
import math
import comtypes.gen.SKCOMLib as sk
#%% Call Liberary %%#
#comtypes.client.GetModule(os.path.split(os.path.realpath(__file__))[0] + r'\SKCOM.dll') #加此行需將API放與py同目錄
skC = comtypes.client.CreateObject(sk.SKCenterLib,interface=sk.ISKCenterLib)
skO = comtypes.client.CreateObject(sk.SKOrderLib,interface=sk.ISKOrderLib)
skQ = comtypes.client.CreateObject(sk.SKQuoteLib,interface=sk.ISKQuoteLib)
skR = comtypes.client.CreateObject(sk.SKReplyLib,interface=sk.ISKReplyLib)
#%% For Call Database's Delay Time %%#
def pumpwait2(t=1):# For Call Database's Delay Time #
    Tend = time.time()+t
    while time.time() < Tend:
        pythoncom.PumpWaitingMessages()        
#%% Captial C++ API Setting %%#
""" Login """
def Login(ID,PW):
    try:
        skC.SKCenterLib_SetLogPath(os.path.split(os.path.realpath("__file__"))[0] + "\\CapitalLog_Quote")
        m_nCode = skC.SKCenterLib_Login(ID,PW)
        if(m_nCode==0 or m_nCode==2003):
            result = "Login Success"
        else:
            result = "Login Fail, " + str(m_nCode)
    except Exception as e:
        result = "Login Fun Error, "+ str(e)
    return result  
""" Quote Call """
class QuoteCall:
    def __init__(self):
        self.Page = 0
        self.QuoteResult = "NA"
    def Enter(self):
        try:
            m_nCode = skQ.SKQuoteLib_EnterMonitor()
            if (m_nCode==0):                
                self.QuoteResult = "Quote Connect Success, "+ str(m_nCode)
            else:
                self.QuoteResult = "Quote Connect Failed, "+ str(m_nCode)
        except Exception as e:
            self.QuoteResult = "Enter Fun Error, "+ str(e)
        return self.QuoteResult
    def Leave(self):
        try:
            m_nCode = skQ.SKQuoteLib_LeaveMonitor()
            if (m_nCode == 0):
                self.QuoteResult = "Quote disConnect Success, "+ str(m_nCode)
            else:
                self.QuoteResult = "Quote disConnect Failed, "+ str(m_nCode)
        except Exception as e:
            self.QuoteResult = "Leave Fun Error, "+ str(e)
        return self.QuoteResult
    def StockPrice(self,stockNo):
        try:
            m_nPage, m_nCode = skQ.SKQuoteLib_RequestStocks(self.Page,stockNo)
            if (m_nCode == 0):
                self.QuoteResult = "StockPrice Success, "+ str(m_nCode)#+ str(m_nPage)
            else:
                self.QuoteResult = "StockPrice Failed, "+ str(m_nCode)#+ str(m_nPage)
        except Exception as e:
            self.QuoteResult = "StockPrice Fun Error, "+ str(e)
        return self.QuoteResult
    def StockData(self, stockNo, scale, formet):
        #Scale: 0 = 1 min trading value, 4 =  day trading value, 5 = week trading value, 6 = month trading value
        #Formet: 0 = old formet, 1 = new formet
        try:
            m_nCode = skQ.SKQuoteLib_RequestKLine(stockNo, scale, formet)
            if(m_nCode==0):
                self.QuoteResult = "StockData Success, "+ str(m_nCode)
            else:
                self.QuoteResult = "StockData Failed, "+ str(m_nCode)
        except Exception as e:
            self.QuoteResult = "StockData Fun Error, "+ str(e)
        return self.QuoteResult
    def StockList(self, sort):
        #Sort: 0 = 上市, 1 = 上櫃, 2 = 期貨, 3 = 選擇權, 4 = 興櫃
        try:
            m_nCode = skQ.SKQuoteLib_RequestStockList(sort)
            if(m_nCode==0):
                self.QuoteResult = "StockList Success, "+ str(m_nCode)
            else:
                self.QuoteResult = "StockList Failed, "+ str(m_nCode)
        except Exception as e:
            self.QuoteResult = "StockList Fun Error, "+ str(e)
        return self.QuoteResult
""" Quote CallBack """
class QuoteBack:
    def __init__(self):
        self.ConnResult=[]
        self.StockData=[]
        self.StockPrice=[]
        self.StockList=[]
    def init(self):
        self.StockData=[]
    def OnConnection(self, nKind, nCode):
        if (nKind == 3001):
            self.ConnResult.append("Connecting!")
        elif (nKind == 3002):
            self.ConnResult.append("DisConnected!")
        elif (nKind == 3003):
            self.ConnResult.append("Stocks ready!")
        elif (nKind == 3021):
            self.ConnResult.append("Connect Error!")
    def OnNotifyQuote(self, sMarketNo, sStockidx):
        pStock = sk.SKSTOCK()
        skQ.SKQuoteLib_GetStockByIndex(sMarketNo, sStockidx, pStock)
        strMsg = pStock.bstrStockNo,pStock.bstrStockName,pStock.nOpen/math.pow(10,pStock.sDecimal),pStock.nHigh/math.pow(10,pStock.sDecimal),pStock.nLow/math.pow(10,pStock.sDecimal),pStock.nClose/math.pow(10,pStock.sDecimal),pStock.nTQty
        self.StockPrice.append(strMsg)
    def OnNotifyKLineData(self,bstrStockNo,bstrData):
        cutData = bstrData.split(',')
        self.StockNo = bstrStockNo
        self.StockData.append(cutData)
    def OnNotifyStockList(self,sMarketNo,bstrStockData):
        self.MarketNo = sMarketNo
        ProductData = bstrStockData.split(';')
        self.StockList.append(ProductData)
""" Order Call """
class OrderCall:
    def __init__(self):
        self.OrderResult = "NA"
        self.OrderNumber = "NA"
    def Initialize(self): # 1.下單物件初始
        try:
            m_nCode = skO.SKOrderLib_Initialize()
            if m_nCode == 0:
                self.OrderResult = "Order Initialize Success, "+ str(m_nCode)
            else:
                self.OrderResult = "Order Initialize Failed, "+ str(m_nCode)
        except Exception as e:
            self.OrderResult = "Initializ Fun Error, "+ str(e)
        return self.OrderResult
    def ReadCert(self,ID): # 2.讀取憑證
        try:
            m_nCode = skO.ReadCertByID(ID)
            if m_nCode == 0 or m_nCode == 2005:
                self.OrderResult = "Order ReadCert Success, "+ str(m_nCode)
            else:
                self.OrderResult = "Order ReadCert Failed, "+ str(m_nCode)
        except Exception as e:
            self.OrderResult = "ReadCert Fun Error, "+ str(e)
        return self.OrderResult
    def AccountGet(self): # 3.取得下單帳號
        try:
            m_nCode = skO.GetUserAccount()
            if m_nCode == 0:
                self.OrderResult = "Get Account Success, "+ str(m_nCode)
            else:
                self.OrderResult = "Get Account Failed, "+ str(m_nCode)
        except Exception as e:
            self.OrderResult = "GetAccount Fun Error, "+ str(e)
        return self.OrderResult
    
    def FutureOpen(self,ID,Account): # 查詢期貨未平倉
        try:
            m_nCode = skO.GetOpenInterestWithFormat(ID,Account,3)
            if m_nCode == 0:
                self.OrderResult = "Open Interest Success, "+ str(m_nCode)
            else:
                self.OrderResult = "Open Interest Failed, "+ str(m_nCode)
        except Exception as e:
            self.OrderResult = "FutureOpen Fun Error, "+ str(e)
        return self.OrderResult   
    def FutureRight(self,ID,Account): # 查詢國內權益數
        try:
            m_nCode = skO.GetFutureRights(ID,Account,1)
            if m_nCode == 0:
                self.OrderResult = "Future Right Success, "+ str(m_nCode)
            else:
                self.OrderResult = "Future Right Failed, "+ str(m_nCode)
        except Exception as e:
            self.OrderResult = "FutureRight Fun Error, "+ str(e)
        return self.OrderResult
    def FutureOrder(self,ID,bAsyncOrder,OrderForm): # 期貨下單 for the same product, maxium order 10 lot for each time
        try:
            self.OrderNumber ,m_nCode = skO.SendFutureOrderCLR(ID,bAsyncOrder,OrderForm)
            if m_nCode == 5:
                self.OrderResult = "Future Single Order Success, "+ str(m_nCode)+", Number:"+str(self.OrderNumber)
            elif m_nCode == 0:
                self.OrderResult = "Future Multiple Orders Success, "+ str(m_nCode)+", Number:"+str(self.OrderNumber)
            else:
                self.OrderResult = "Future Order Failed, "+ str(m_nCode)+", Number:"+str(self.OrderNumber)
        except Exception as e:
            self.OrderResult = "FutureOrder Fun Error, "+ str(e)
        return self.OrderResult
    
    def StockProfit(self,ID,Account): # 證券損益試算
        try:
            m_nCode = skO.GetRequestProfitReport(ID,Account)
            if m_nCode == 0:
                self.OrderResult = "Stock Profit Success, "+ str(m_nCode)
            else:
                self.OrderResult = "Stock Profit Failed, "+ str(m_nCode)
        except Exception as e:
            self.OrderResult = "StockProfit Fun Error, "+ str(e)
        return self.OrderResult    
    def StockBalance(self,ID,Account): # 查詢證券庫存
        try:
            m_nCode = skO.GetRealBalanceReport(ID,Account)
            if m_nCode == 0:
                self.OrderResult = "Stock Balance Success, "+ str(m_nCode)
            else:
                self.OrderResult = "Stock Balance Failed, "+ str(m_nCode)
        except Exception as e:
            self.OrderResult = "StockBalance Fun Error, "+ str(e)
        return self.OrderResult
    def StockOrder(self,ID,bAsyncOrder,OrderForm): # 證券下單
        try:
            OrderNumber ,m_nCode = skO.SendStockOrder(ID,bAsyncOrder,OrderForm)
            if m_nCode == 5:
                self.OrderResult = "Stock Single Order Success, "+ str(m_nCode)+", Number:"+str(OrderNumber)
            elif m_nCode == 0:
                self.OrderResult = "Stock Multiple Orders Success, "+ str(m_nCode)+", Number:"+str(OrderNumber)
            else:
                self.OrderResult = "Stock Order Failed, "+ str(m_nCode)
        except Exception as e:
            self.OrderResult = "StockOrder Fun Error, "+ str(e)
        return self.OrderResult
""" Order CallBack """
class OrderBack:
    def __init__(self):
        self.AccountGet = []
        self.FutureOpen= []
        self.FutureRight= []
        self.StockBalance= []
        self.StockProfit= []
        self.OrderFeedback = []
    def OnAccount(self, bstrLogInID, bstrAccountData):
        AccountData= bstrAccountData.split(',')
        self.AccountGet.append(AccountData)
    
    # Reture期貨未平倉
    def OnOpenInterest(self, OpInData):
        OpenInterestData = OpInData.split(',')
        self.FutureOpen.append(OpenInterestData)
        #市場別, 帳號, 商品, 買賣別, 未平倉部位, 當沖未平倉部位, 成本(三位小數), 一點價值, 單口手續費, 交易稅(萬分之X)

    # Reture國內權益數
    def OnFutureRights(self, FuRiData):
        FutureRightData = FuRiData.split(',')
        self.FutureRight.append(FutureRightData)
        #0 帳戶餘額,1 浮動損益,2 已實現費用,3 交易稅,4 預扣權利金,5 權利金收付,6 權益數v,7 超額保證金,8 存提款,9 買方市值x
        #10 賣方市值x,11 期貨平倉損益,12 盤中未實現,13 原始保證金,14 維持保證金,15 部位原始保證金,16 部位維持保證金,17 委託保證金
        #18 超額最佳保證金,19 權利總值,20 預扣費用,21 原始保證金,22 昨日餘額,23 選擇權組合單加不加收保證金,24 維持率
        #25 幣別,26 足額原始保證金,27 足額維持保證金,28 足額可用,29 抵繳金額,30 有價可用,31 超額保證金,32 足額現金可用
        #33 有價價值,34 風險指標,35 選擇權到期差異x,36 選擇權到期差損x,37 期貨到期損益,38 加收保證金

    # Reture證券庫存
    def OnRealBalanceReport(self, BeReData):
        BalenceReportData = BeReData.split(',')
        self.StockBalance.append(BalenceReportData)
        
    # Reture證券損益試算
    def OnRequestProfitReport(self, PrReData):
        ProfitReportData = PrReData.split(',')
        self.StockProfit.append(ProfitReportData)
        
    #Multiple Order Return
    def OnAsyncOrder(self, ThreadID, Code, Message):
        self._Thread = ThreadID
        self._nCode = Code
        self._msg = Message
        self.OrderFeedback.append([self._Thread,self._nCode,self._msg])

class OrderForm:
    def Stock(self,account,stockNo,sBuySell,sPrice,sQty):
        # 建立下單用的參數(STOCKORDER)物件(下單時要填股票代號,買賣別,委託價,數量等等的一個物件)
        oOrder = sk.STOCKORDER()
        # 填入完整帳號
        oOrder.bstrFullAccount = account
        # 填入股票代號
        oOrder.bstrStockNo = stockNo
        # 上市櫃 = 0、興櫃 = 1
        oOrder.sPrime = 0 #sPrime
        # 盤中 = 0、盤後 = 1、零股 = 2
        oOrder.sPeriod = 0 #sPeriod
        # 現股 = 0、融資 = 1、融券 = 2
        oOrder.sFlag = 0 #sFlag
        # 買進 = 0、賣出 = 1        
        if sBuySell == "Buy":
            oOrder.sBuySell = 0
        elif sBuySell == "Sell":
            oOrder.sBuySell = 1
        # 委託價、參考昨日收盤價 = M
        oOrder.bstrPrice = sPrice
        # 委託數量
        oOrder.nQty = int(sQty)
        return oOrder
    def Future(self,account,stockNo,sBuySell,sPrice,sQty,tradetype,reservetype):
        # 建立下單用的參數(FUTUREORDER)物件(下單時要填商品代號,買賣別,委託價,數量等等的一個物件)
        oOrder = sk.FUTUREORDER()
        # 填入完整帳號
        oOrder.bstrFullAccount =  account
        # 填入期權代號
        oOrder.bstrStockNo = stockNo
        # 買進 = 0、賣出 = 1
        if sBuySell == "Buy":
            oOrder.sBuySell = 0
        elif sBuySell == "Sell":
            oOrder.sBuySell = 1
        # ROD = 0、IOC = 1、FOK = 2
        if tradetype == 'ROD':
            oOrder.sTradeType = 0
        elif tradetype == 'IOC':
            oOrder.sTradeType = 1
        elif tradetype == 'FOK':
            oOrder.sTradeType = 2
        # 非當沖 = 0、當沖 = 1
        oOrder.sDayTrade = 0
        # 委託價、市價 = M
        oOrder.bstrPrice = sPrice
        # 委託數量
        oOrder.nQty = int(sQty)
        # 新倉 = 0、平倉 = 1、自動 = 2
        oOrder.sNewClose = 2
        # 盤中 = 0、T盤預約 = 1
        oOrder.sReserved = reservetype
        return oOrder

#%% Capital Python Lib Setting %%#
"""
AQuoteCall = QuoteCall()
AQuoteBack = QuoteBack()
AOrderCall = OrderCall()
AOrderBack = OrderBack()
"""
#QuoteBackActive = comtypes.client.GetEvents(skQ, AQuoteBack)
#OrderBackActive = comtypes.client.GetEvents(skO, AOrderBack)
#QuoteCall = QuoteCallback()
#CallResult = QuoteCall.Callresult
#StockDataCall = QuoteCall.StockDataCall
"""
ID = "*********" # --> Privacy Info
PW = "*********" # --> Privacy Info
ALogin =Login(ID,PW)
Enter = AQuoteCall.Enter()
pumpwait2(8)
EnterResult = AQuoteBack.ConnResult
#Leave = AQuote.Leave()
"""