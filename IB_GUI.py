from Tkinter import *
import datetime
from ib.ext.Contract import Contract
from ib.opt import ibConnection, message
import tkMessageBox
from ib.ext.Order import Order
from time import sleep
import pandas as pd
import time


def newOrder_GUI(action, orderID, quantity,startTime):
    global trade_params
    
    newStkOrder = Order()
    print orderID, action, trade_params['OrderTIF'],trade_params['OrderType'],quantity,trade_params['Account'],startTime
    newStkOrder.m_orderId = orderID
    
    newStkOrder.m_action = action
    newStkOrder.m_tif = 'DAY'
    newStkOrder.m_transmit = True
    newStkOrder.m_orderType = 'MKT'
    newStkOrder.m_totalQuantity = quantity
    newStkOrder.m_account = 'DU164541'
    newStkOrder.m_goodAfterTime = startTime
    #newStkOrder.m_goodAfterTime=endTime
    return newStkOrder
        
def read_config_GUI():
    global config_dir, trade_params
    #try:
    # read config.txt file in config folder
    with open(config_dir + '/config.txt','rb') as f:
        config = f.read()
    # read each line
    for line in config.replace('\r','').split('\n'):
        try:
            if len(line)>0:
                # split key and values by "="
                split = line.split('=')
                # type conversion
                if split[0] in ['TradeDivisor','CapitalCap']:
                    trade_params[split[0]] = float(split[1])
                else:
                    trade_params[split[0]] = split[1]
        except:
            # exception handling
            print '>> File "config/config.txt" is in wrong format. Program exits.'
            exit(1)

#Create new Contract
def newContract_GUI(contractTuple):
    newContract = Contract()
    symbol=contractTuple[0]
    newContract.m_symbol = symbol
    newContract.m_secType = contractTuple[1]
    newContract.m_exchange='SMART'
    #download exchange information from yahoo finance
    url='http://download.finance.yahoo.com/d/quotes.csv?s='+symbol+'&f=x'
    primexchange=pd.read_csv(url,header=None)
    primexchange=primexchange.iloc[0,0]
    if primexchange=='NasdaqNM' or primexchange=='NGM':
        primexchange='NASDAQ'
    #print primexchange
    newContract.m_primaryExch=primexchange

    return newContract


def Exit():
    return None

class Setup_Current_Order_Frame(LabelFrame):

    def __init__(self,parent):
        #Create list box with scrollbar on the right side.
        self.scr=Scrollbar(parent)
        
        self.scr.pack(side=RIGHT,fill=Y)
        self.Alist=Listbox(parent,width=150,height=5,yscrollcommand=self.scr.set)
        #insert title of list box
        self.Alist.insert(1,'Seq ID'+'      '+'Symbol'+'        '+'Effective'+'     '+'Start Time'+'        '+'End Time'+'      '+'Trims'+
                          '     '+'%'+'     '+'IB B/S'+'        '+'IB AMT'+'        '+'IB ID'+'     '+'Trims #'+'       '+'E?'+'        '+'is Cancelled?'+'     '+'Created On')
        self.Alist.pack(side=LEFT,fill=BOTH)
        self.scr.config(command=self.Alist.yview)
    #insert new observation to list box
    def add_entry(self,index,order_info):
        
        order_info=map(str,order_info)
        self.Alist.insert(index,'           '.join(order_info))
    def remove(self,index):
        self.Alist.delete(index)

class Setup_Add_Order_Frame(LabelFrame):
    def __init__(self,parent,con):
        def _updatePortfolio(msg):
            #remove symbols in list of symbols that we typed in until there is no symbols in that list.
            
            if len(self.Symbol_List)!=0:
                    
                if msg.contract.m_symbol in self.Symbol_List:
                        
                    self.Symbol_List.remove(msg.contract.m_symbol)
                    self.Symbols.append([msg.contract.m_symbol,msg.contract.m_secType,msg.position])
            else:
                print 'ALL_Received'
                self.All_Received=True
        def _nextValidId(msg):
            self.trade_count = msg.orderId

        def _orderStatus(msg):
            print msg.orderId,msg.status
            if msg.status=='Cancelled':
                Cancelled='Y'
                Status='N'
            elif msg.status=='Filled':
                Status='Y'
                Cancelled='N'
            else:
                Status='N'
                Cancelled='N'
            if len(self.send_order[msg.orderId])==14:
                self.send_order[msg.orderId][12]=Status
                self.send_order[msg.orderId][13]=Cancelled
            else:
                self.send_order[msg.orderId].extend([Status,Cancelled])

            if msg.orderId in self.send_order.keys():
                print 'add',self.send_order[msg.orderId][0]-1,'remove',self.send_order[msg.orderId][0]+1
                self.datagridview.add_entry(self.send_order[msg.orderId][0],self.send_order[msg.orderId])
                self.datagridview.remove(self.send_order[msg.orderId][0]+1)
            


        


        self.con=con
        self.con.register(_updatePortfolio, 'UpdatePortfolio')
        self.con.register(_nextValidId,message.nextValidId)
        self.con.register(_orderStatus,message.orderStatus)
        sleep(2)
        
        LabelFrame.__init__(self,parent,text='New Order')
        self.parent=parent

        self.DoWeNeedToRegister=True
        
        #Lay out entries and labels in frame.
        self.Order_Id=Label(self,text='Order Id')
        self.Order_Id.place(anchor='w',relx=0,rely=0.05)
        
        self.Symbol=Label(self,text='Symbol(s)')
        self.Symbol.place(anchor='w',relx=0,rely=0.15)

        self.Effective_Date=Label(self,text='Effective Date (format: YYYYMMDD)')
        
        self.Effective_Date.place(anchor='w',relx=0,rely=0.25)

        self.First_trim_time=Label(self,text='First trim time (format: HH:mm:ss)')
        self.First_trim_time.place(anchor='w',relx=0,rely=0.35)

        self.Last_trim_time=Label(self,text='Last trim time (format: HH:mm:ss)')
        self.Last_trim_time.place(anchor='w',relx=0,rely=0.45)

        self.Trims=Label(self,text='Trims (1-10)')
        self.Trims.place(anchor='w',relx=0,rely=0.55)

        self.percent=Label(self,text='percent (optional, 1-99)')
        self.percent.place(anchor='w',relx=0,rely=0.65)

        self.Symbol_Entry=Entry(self)
        self.Symbol_Entry.place(anchor='e',relx=0.9,rely=0.15)

        self.EffDate_Entry=Entry(self)
        self.EffDate_Entry.insert(0,datetime.datetime.strftime(datetime.datetime.now(),'%Y%m%d'))
        self.EffDate_Entry.place(anchor='e',relx=0.9,rely=0.25)

        self.FirstTrimTime_Entry=Entry(self)
        self.FirstTrimTime_Entry.place(anchor='e',relx=0.9,rely=0.35)

        self.Last_Trim_Time_Entry=Entry(self)
        self.Last_Trim_Time_Entry.place(anchor='e',relx=0.9,rely=0.45)

        self.Trims_Entry=Entry(self)
        self.Trims_Entry.place(anchor='e',relx=0.9,rely=0.55)

        self.Percent_Entry=Entry(self)
        self.Percent_Entry.place(anchor='e',relx=0.9,rely=0.65)
        self.datagridview=Setup_Current_Order_Frame(self.parent)

        print 'initial'
        Button(self, text='Add Order',command=self.Add_Order).place(relx=0.8,rely=0.8)
        Button(self, text='reqAccountUpdate',command=self.requestUpdate).place(relx=0.2,rely=0.8)

        #initial some variables.
        #this is IB id.
        self.trade_count=0

        

        #store order that we placed to IB.
        self.send_order={}

        #to identify which row in curr_order_frame.
        self.row=0


        
    def requestUpdate(self):
        
        #read content in Symbol entry and split it.
        string=self.Symbol_Entry.get()
        
        if string=='':
            tkMessageBox.showerror('Warning','You did not enter any symbols!')
            return
        
        #read start time of order.
        EffDate=self.EffDate_Entry.get()
        FirstTrimTime=self.FirstTrimTime_Entry.get()
        if FirstTrimTime=='':
            tkMessageBox.showerror('Warning','You did not enter start time!')
            return
        
        
        Last_Trim_Time=self.Last_Trim_Time_Entry.get()

        
        if Last_Trim_Time=='':
            Last_Trim_Time='15:59:59'

        #convert string into time format.
        try:
            start_time=datetime.datetime.strptime(FirstTrimTime,'%H:%M:%S')
            end_time=datetime.datetime.strptime(Last_Trim_Time,'%H:%M:%S')
        except:
            tkMessageBox.showerror('Warning','You did not enter valid start and end time!')
            return
        if start_time>end_time:
            tkMessageBox.showerror('Warning','End Time should not be Earlier than Start Time!')
            return

        
        #Calculate time difference between start and end time in seconds.
        timediff=end_time-start_time
        timediff=timediff.total_seconds()

        
        Trims=self.Trims_Entry.get()
        Percent=self.Percent_Entry.get()

        #Handle the case that either Percent or Trims entry does not have content.
        if Percent=='' and Trims!='':
            Percent=100.0/int(Trims)
        if Percent=='' and Trims=='':
            Percent=100
            Trims=1
        if Percent!='' and Trims=='':
            Percent=float(Percent)
            Trims=1
        if Percent!='' and Trims!='':
            Percent=float(Percent)
            Trims=int(Trims)
            
            
        Trims=int(Trims)

        #If want to close position larger than what we have in portofolio, report an warning.
        if Trims*float(Percent)>100:
            tkMessageBox.showerror('Warning','Your Are Trying To Close Positions Which Are More Than That What We Have in Portfolio!')
            return

        #Set time interval between trims.
        if Trims==1:
            time_interval=timediff
        else:
            time_interval=timediff/(Trims-1)
            
        #this list is to store different trim times.
        timer=[]

        #set trim times.
        for i in range(Trims):
            exe_time=datetime.datetime.strftime(start_time+datetime.timedelta(seconds=(i*time_interval)),'%H:%M:%S')
            
            timer.append(exe_time)

        #list of symbols we typed in
        self.Symbol_List=string.split(' ')

        #Store symbols we get from current portfolio.
        self.Symbols=[]
        self.All_Received=False
        #call _updateportfolio function.
        self.con.reqAccountUpdates(1,'DU164541')
        t0=time.time()
        
        #Until we get all information of symbols we typed in from portofolio or lag time more than 30 seconds, we do not proceed to next step
        #####################################################################################################################################
        while not self.All_Received and time.time()-t0<10:
            print self.Symbol_List
            sleep(1)
        
        self.All_Received=False
        if len(self.Symbols)==0:
            tkMessageBox.showerror('Warning','You did not enter any symbols which are in current positions!')
            return

        self.con.reqAccountUpdates(0,'DU164541')
        #store orders
        self.OrderQueue=[]

        #request orderId from IB
        self.con.reqIds(1)
        #sleep for 1 senconds to allow enough time to get nextValidId from IB.
        sleep(1)
        #this loop is to form orders and store them into self.orderQueue
        for j in range(1,int(Trims)+1):
            for each_order in self.Symbols:
                #point out which row it should be in list box.
                self.row+=1
                
                position=int(each_order[2])
                if position<0:
                    action='BUY'
                elif position>=0:
                    action='SELL'
                    
                quantity=int((float(Percent)/100)*abs(position))
                t=EffDate+' '+timer[j-1]
                stkOrder=newOrder_GUI(action,self.trade_count,quantity,t)
                stkContract=newContract_GUI(each_order)
                
                self.OrderQueue.append((self.trade_count,stkContract,stkOrder))
                #store orders into send_order dict.
                self.send_order[self.trade_count]=[self.row,each_order[0],EffDate,FirstTrimTime,Last_Trim_Time,
                                                   Trims,Percent,action,quantity,self.trade_count,j,
                                                   datetime.datetime.strftime(datetime.datetime.now(),'%Y%m%d%H%M%S')]

                #increase trade_count(order_id) by 1.
                self.trade_count+=1
        
        
    
    def Add_Order(self):
        for i in range(len(self.OrderQueue)):
            starttime=self.OrderQueue[i][2].m_goodAfterTime.replace(' ','')
            if datetime.datetime.strptime(starttime,'%Y%m%d%H:%M:%S')<datetime.datetime.now():
                self.OrderQueue[i][2].m_goodAfterTime=self.EffDate_Entry.get()+' '+datetime.datetime.strftime(datetime.datetime.now()+datetime.timedelta(seconds=30),'%H:%M:%S')

            self.con.placeOrder(self.OrderQueue[i][0],self.OrderQueue[i][1],self.OrderQueue[i][2])

        self.con.reqOpenOrders()

class Setup_Cancel_Order_Frame(LabelFrame):
    def __init__(self,parent,con):
        LabelFrame.__init__(self,parent,text='Cancel Order')

        self.con=con
        
        self.Order_Id_Label=Label(self,text='Order Id')

        self.Order_Id_Entries=Entry(self)

        self.Order_Id_Entries.pack()



        Button(self, text='Cancel Order',command=self.Cancel_Order).pack()

    def Cancel_Order(self):
            cancels=self.Order_Id_Entries.get()

            cancels=cancels.split(' ')

            cancels=map(int,cancels)
            for i in cancels:
                self.con.cancelOrder(i)

def initial_GUI(con):
    global log_dir,config_dir,trade_params
    log_dir = './log'
    config_dir = './config'

    # counters
    trade_count = 0
    round_digit= 0
    buyingPower = 0
    leverage = 0 


    # trade parameters
    trade_params = {}

    # read configuration
    print '[Read configuration]'
    read_config_GUI()

    trade_count=0
    Symbol_List=[]
    root=Tk()
    Rtitle=root.title('Close Positions')

    #root.resizable(width=False, height=False)

    root.geometry("600x500+350+150")
    #setup frames

    
    New_Order_Fm=Setup_Add_Order_Frame(root,con)
    New_Order_Fm.place(anchor='w',relx=0.7,rely=0.8,width=380,height=220)

    Cancel_Order_Fm=Setup_Cancel_Order_Frame(root,con)
    Cancel_Order_Fm.place(anchor='e',relx=1,rely=0.52,width=200)

    root.mainloop()

