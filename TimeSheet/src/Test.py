'''
Created on Feb 21, 2012

@author: Joselle Abagat
'''
from clsBWNamerun import *
from clsCAM import *
from clsContract import *
from clsEmployee import *
from clsMonthlyCalendar import *
from clsNetwork import *
from clsWeeklyCalendar import *
import re, os, sys, easygui
from clsETC import *
from clsReloadResources import *
import win32api
import win32com.client as win32
import wx
import datetime
import tempfile
import webbrowser
import stat, shutil
import pythoncom

print datetime.datetime.now()
print datetime.date.now()
dateConvert = datetime.datetime.strptime("2012-09-18", "%Y-%m-%d")
print dateConvert
print dateConvert.date()
dateConvert = datetime.datetime.strptime("2012-09-18 00:00:00", "%Y-%m-%d")
dateConvert = dateConvert.strftime("%m/%d/%Y")
dateConvert = datetime.datetime.strptime(dateConvert, "%m/%d/%Y")
print dateConvert.date()
ymd = dateConvert.date()
print ymd.strftime("%Y%m")
print dateConvert.date() < datetime.datetime.now().date()
print dateConvert.strftime("%Y%m%d")
print datetime.datetime.now().strftime("%Y%m")
print "20" + datetime.datetime.now().strftime("%y") + datetime.datetime.now().strftime("%m")

#print pythoncom._GetInterfaceCount()
#print win32api.FormatMessage(-2147418111)
#xl = win32.DispatchEx('Excel.Application')
#print pythoncom._GetInterfaceCount()
#xl2 = win32.DispatchEx('Excel.Application')
#print pythoncom._GetInterfaceCount()
#xl3 = win32.DispatchEx('Excel.Application')
#print pythoncom._GetInterfaceCount()
#
#xl.Visible = True
#xl2.Visible = True
#xl3.Visible = True
#
#xl.Quit()
#xl2.Quit()
#xl3.Quit()

#cur = os.getcwd()

#winmpm = os.path.abspath("C:\Program Files\WINMPM")
#print winmpm
#
#direct = 'resources'
#timesheet = os.path.join(os.getenv("USERPROFILE"),"My Documents\Timesheet")
#timestamp = os.path.getmtime(timesheet)
#print timestamp
#print datetime.datetime.fromtimestamp(timestamp)
#print datetime.date.fromtimestamp(timestamp)
#print os.access(timesheet, os.W_OK)
#print os.path.abspath(timesheet)
#print os.stat(timesheet)
#os.chmod(timesheet, stat.S_IWRITE)
#os.chmod(timesheet, stat.S_IWGRP)
#print os.stat(timesheet)

#shutil.rmtree(timesheet, True)
#print os.path.abspath(timesheet)

#os.unl
#webadd = 'http://wiki.northgrum.com/wiki/HALE_TimeSheet'
#webbrowser.open(webadd, 0, True)
                   
#win32.Dispatch('Excel.Application')
#win32c = win32.constants
#print win32c.__dicts__
#
#print win32com.__gen_path__
#print win32.__path__
#
#tempgencache = os.path.join(os.path.dirname(tempfile.TemporaryFile().name), 'gencache')
#print os.path.abspath(tempgencache)
#
#win32c.__init__()
#
#mso10lib = win32.gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 2)
#print mso10lib

#print datetime.date.today()

#print win32api.FormatMessage(-2147352565)
#print win32api.FormatMessage(-2147352561)
#print win32api.FormatMessage(-2147352567)
#print win32api.FormatMessage(-2146827284)
#print win32api.FormatMessage(-2147352571)

#class App(wx.App):
#
#    def OnInit(self):
#        frame = MainFrame()
#        frame.Show()
#        self.SetTopWindow(frame)
#        return True
#
#class MainFrame(wx.Frame):
#
#    title = "Main Frame"
#
#    def __init__(self):
#        wx.Frame.__init__(self, None, 1, self.title) #id = 5
#
#        menuFile = wx.Menu()
#
#        menuAbout = wx.Menu()
#        menuAbout.Append(2, "&About...", "About this program")
#
#        menuBar = wx.MenuBar()
#        menuBar.Append(menuAbout, "&Help")
#        self.SetMenuBar(menuBar)
#
#        self.CreateStatusBar()
#
#        self.Bind(wx.EVT_MENU, self.OnAbout, id=2)
#
#    def OnQuit(self, event):
#        self.Close()
#
#    def OnAbout(self, event):
#        AboutFrame().Show()
#
#class AboutFrame(wx.Frame):
#
#    title = "About this program"
#
#    def __init__(self):
#        wx.Frame.__init__(self, 1, -1, self.title) #trying to set parent=1 (id of MainFrame())
#
#
#if __name__ == '__main__':
#    app = App(False)
#    app.MainLoop()



#
#import  wx
#
#class MyFrame(wx.Frame):
#    def __init__(self, parent, title):
#        wx.Frame.__init__(self, parent, title = title, size=(200,100))
#        self.control = wx.TextCtrl(self, style = wx.TE_MULTILINE)
#        self.show(True)
#
#app = wx.App(False)
#frame = MyFrame(None,'')
#
#frame.Show(True)
#app.MainLoop()
#
##compile = src.compileAllMPMFiles()
#
##file = "..\\resources\\Namerun.csv"
##sap = Nameruns(file)
##for each in sap.namerun_list:
##    print(each)
#
#def SearchTitle(lMatch, Title):
#    '''
#    returns the position of the title in the list of the CSV file
#    '''
#    
#    ListTitle = []
#    
#    for item in lMatch:
#        ListTitle.append(item.strip())
#    
#    index = 0
#    for item in ListTitle:
#        regex = re.compile(Title, re.I)
#        findre = regex.search(item)
#        
#        if not findre == None:
#            return index
#        index = index + 1
#    
#    return None


#file = "..\\resources\\EMDLBRF.csv"
#readcsv = csv.reader(open(file, 'r'), dialect = "excel", quotechar = '"', delimiter = ",")
#headerfile = readcsv.next()
#found = SearchTitle(headerfile, "RESP")
#print(found)

#print(headerfile)
#for each in readcsv:
#    print(each[2].strip())

#etcs = ETC(file)
#count = 0
#for each in etcs.etc_list:
#    print(each)
#    count += 1
#    if count == 5:
#        pass

#file2 = "..\\resources\\Namerun3.csv"
#sap = Nameruns(file2)
#count = 0
#for each in sap.namerun_list:
#    print(each)
#    count += 1
#    if count == 5:
#        break

#etccams = etcs.networkToCAM_dict
#for each in etccams:
#    print(each)
#    print(etccams[each])
#print(etcs.getContractFromNetwork("LX806HBD1"))
#print(etcs.getPOPfromNetwork("LX806HBD1"))
#print(etcs.getETCFromNetworkCAMAndPeriod("LX806HBD1", "HBX", 201207))
#print(etcs.getDescriptionFromNetwork("LX806HBD1"))



#(self, WorkCC, HomeCC, SCE, NetworkNo, Description, Activity, Supp1, Supp2, Supp3, Ship, EmpID, 
# EmpName, PayrollCode, Shift, Source, HrsCode, WorkDate, WeekEndDate, HoursValid, HoursPaid):
#===============================================================================
# def loadCAM(filepath):
#    "FUNCTION TO LOAD CAMS; need to be in CSV format"
#    
#    myCAMs = CAM
#    
#    filename = open(filepath,'r')
#    CAMlist = []
#    
#    for CAMLine in filename:
#        if not CAMLine.startswith('//'):
#            CAMItem = CAMLine.split(',')
#            CAMlist.append(myCAMs(CAMItem[0].strip(),CAMItem[1].strip(),CAMItem[2].strip(),CAMItem[3].strip()))
#    
#    return CAMlist
#===============================================================================

#===========================================================================
    # def loadContract(self, filepath):
    # '''
    # FUNCTION TO LOAD CONTRACTS; need to be in CSV format"
    # '''
    #    
    # myContracts = Contract
    # 
    # filename = open(filepath,'r')
    # Contractlist = []
    # 
    # for ContractLine in filename:
    #    if not ContractLine.startswith('//'):
    #        ContractItem = ContractLine.split(',')
    #        Contractlist.append(myContracts(ContractItem[0].strip(), ContractItem[1].strip()))
    # 
    # return Contractlist
    # 
    # def CallContractsDict(self, filepath):
    #    '''
    #    ' CALL CONTRACT DICTIONARY
    #    '''
    #    ContractClass = Contract
    #    
    #    MyContracts = loadContract(filepath)
    #    
    #    lContract = []
    #    dictContract = {}
    #    
    #    for item in MyContracts:
    #        
    #        lContract = ContractClass.ContractItem(item)
    #        #print(lContract)
    #        dictContract[int(lContract[0])] = lContract[1]
    #        
    #    return dictContract
    #===========================================================================

#===============================================================================
# #FUNCTION TO LOAD CLASS: NETWORK
# def loadNetwork(filepath):
#    "need to be in CSV format"
#    
#    myNetworks = Network
#    
#    filename = open(filepath,'r')
#    Networklist = []
#    
#    for NetworkLine in filename:
#        if not NetworkLine.startswith('//'):
#            NetworkItem = NetworkLine.split(',')
#            Networklist.append(myNetworks(NetworkItem[0], NetworkItem[1], NetworkItem[2], NetworkItem[3], NetworkItem[4], NetworkItem[5], NetworkItem[6], NetworkItem[7], NetworkItem[8]))
#    
#    return Networklist
# 
# def CallMapNetworkContract(filepath):
#    
#    MyNetworks = loadNetwork(filepath)
#    lNetworkContract = []
#    network = Network
#    
#    for each_item in MyNetworks:
#        if not network.MapNetworkToContract(each_item) in lNetworkContract:
#            lNetworkContract.append(network.MapNetworkToContract(each_item))
#    
#    return lNetworkContract
# 
# def CallMapNetworkCAM(filepath):
#    
#    network = Network
#    MyNetworks = loadNetwork(filepath)
#    lNetworkContract = []
#    
#    for each_item in MyNetworks:
#        if not network.MapNetworkToCAM(each_item) in lNetworkContract:
#            lNetworkContract.append(network.MapNetworkToCAM(each_item))
#    
#    return lNetworkContract
# 
# def CallNetworkActivity(filepath):
#    
#    MyNetworks = loadNetwork(filepath)
#    lNetworkContract = []
#    network = Network
#    
#    for each_item in MyNetworks:
#        if not network.NetworkActivityPair(each_item) in lNetworkContract:
#            lNetworkContract.append(network.NetworkActivityPair(each_item))
#    
#    return lNetworkContract
#===============================================================================
    
#===============================================================================
# #FUNCTION TO LOAD CLASS: BWNAMERUN
# def loadBWNamerun(filepath):
#    "need to be in CSV format"
#    
#    mySAPNamerun = Namerun
#    
#    filename = open(filepath,'r')
#    BWNamerunlist = []
#    
#    for BWNamerunLine in filename:
#        if not BWNamerunLine.startswith('//'):
#            BWNamerunItem = BWNamerunLine.split(',')
#            BWNamerunlist.append(mySAPNamerun(BWNamerunItem[0], BWNamerunItem[1], BWNamerunItem[2], BWNamerunItem[3], BWNamerunItem[4], BWNamerunItem[5], BWNamerunItem[6], BWNamerunItem[7], BWNamerunItem[8], BWNamerunItem[9], BWNamerunItem[10], BWNamerunItem[11], BWNamerunItem[12], BWNamerunItem[13], BWNamerunItem[14], BWNamerunItem[15], BWNamerunItem[16], BWNamerunItem[17]))
#    
#    return BWNamerunlist
#===============================================================================
    
#===============================================================================
# #FUNCTION TO LOAD CLASS: EMPLOYEE
# def loadEmployee(filepath):
#    "need to be in CSV format"
# 
#    myEmployees = Employee
#    
#    filename = open(filepath,'r')
#    Employeelist = []
#    
#    for EmployeeLine in filename:
#        if not EmployeeLine.startswith('//'):
#            EmployeeItem = EmployeeLine.split(',')
#            Employeelist.append(myEmployees(EmployeeItem[0], EmployeeItem[1], EmployeeItem[2], EmployeeItem[3], EmployeeItem[4], EmployeeItem[5]))
#    
#    return Employeelist
#===============================================================================

#===============================================================================
# filename = "CAMS.csv"
# test = loadCAM(filename)
# for item in test:
#    print(item)
#===============================================================================

#===============================================================================
# filename = "Contracts.csv"
# test = loadContract(filename)
# print(test)
# 
# counter = 0
# for item in test:
#    counter = counter + 1
# print(counter)
#===============================================================================

#===============================================================================
# 
# test2 = CallContractsDict(filename)
# print(test2)
#===============================================================================

#===============================================================================
# filename = "Networks.csv"
# test3 = CallMapNetworkContract(filename)
# for each_item in test3:
#    print(each_item)
# test4 = CallMapNetworkCAM(filename)
# for item in test4:
#    print(item)
# test5 = CallNetworkActivity(filename)
# for item in test5:
#    print(item)
#===============================================================================

#filename = "WeeklyAcctCal.csv"
#test6 = WeeklyCalendar
#printtest = test6.CallWeeklyCalendar(test6,filename)
#for item in printtest:
#    print(item)
#print(printtest)
#sample = printtest
#print(len(sample))
#print(sample[420][1])

#===============================================================================
# f = open("CAMS.csv", 'r')
# 
# def trial(f,strng):
#    data = f.readlines()
#    print(data[0])
#    splitdata = data[0].split(',')
#    newlist = []
#    for item in splitdata:
#        newlist.append(item.strip())
#    print(newlist)
#    counter = 0
#    for item in newlist:
#        regex = re.compile(strng,re.I)
#        findre = regex.match(item)
#    #    print(findre)
#        if not findre == None:
#            return counter
#        counter = counter + 1
#    #print(newlist[None])
#    return 0
# found = trial(f,"last name")
# print(found)
#===============================================================================
#===============================================================================
# 
# csvfile = "ETCs.csv"
# #readfile = csvfile.readlines()
# #firstline = readfile[0].split(',')
# #print(firstline)
# etctest = ETC.loadETC(ETC, csvfile)
# print(etctest[0])
# print(etctest[1])
# #for item in etctest:
# #    print(item)
#===============================================================================

#emp = "134641"
#try:
#    empid = len(emp)
#    print(empid)
#except ValueError:
#    empname = emp
#    print(empname)

#network = "Sustaining Engineering and Technical Ser"
#length = len(network)
#if length == 9:
#    print("ok")
#else:
#    print(None)

