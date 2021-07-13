'''
Created on Feb 29, 2012

@author: Joselle Abagat
'''

from clsBWNamerun import *
from clsCAM import *
from clsContract import *
from clsEmployee import *
from clsMonthlyCalendar import *
from clsNetwork import *
from clsWeeklyCalendar import *
import re
from clsETC import *
import os

def loadNamerun(filepath):
    '''
    'FUNCTION TO LOAD SAP Namerun; need to be in CSV format"
    '''
    
    mySAPNamerun = NamerunEntry
    
    filename = open(filepath,'r')
    BWNamerunlist = []
    readfile = filename.readlines()
    
    if readfile[0] in readfile:
        item = readfile[0].split(',')
        WorkCC_index = SearchTitle(item, "Work")
        HomeCC_index = SearchTitle(item, "Home")
        SCE_index = SearchTitle(item, "SCE")
        Activity_index = SearchTitle(item, "Activity")
        Supp1_index = SearchTitle(item, "Supplemental 1")
        Supp2_index = SearchTitle(item, "Supplemental 2")
        Supp3_index = SearchTitle(item, "Supplemental 3")
        Ship_index = SearchTitle(item, "Ship")
        PayrollCode_index = SearchTitle(item, "Payroll Code")
        Shift_index = SearchTitle(item, "Shift")
        Source_index = SearchTitle(item, "Source")
        HrsCode_index = SearchTitle(item, "Hours Code")
        WorkDate_index = SearchTitle(item, "Work Date")
        WeekEndDate_index = SearchTitle(item, "Weekending Date")
        HoursValid_index = SearchTitle(item, "Hours Valid")
        HoursPaid_index = SearchTitle(item, "Hours Paid")
        
        ###############
        # NOT REAL INDEX VALUES
        ###############
        emp_index = SearchTitle(item, "Employee")
        network_index = SearchTitle(item, "Network")
        
    if readfile[1] in readfile:
        item = readfile[1].split(',')
        try:
            emp_info = int(item[emp_index].strip())
            EmpID_index = emp_index
            EmpName_index = emp_index + 1
        except ValueError:
            EmpName_index = emp_index
            EmpID_index = emp_index + 1     
    
    if NetworkInNamerun(network_index, readfile) == True:
        NetworkNo_index = network_index
        Description_index = network_index + 1
    else:
        NetworkNo_index = network_index + 1
        Description_index = network_index           

    for BWNamerunLine in readfile:
        if not BWNamerunLine == readfile[0]:
            # need a way to not split descriptions that contain commas
            item = BWNamerunLine.split(',')
            
            WorkCC = item[WorkCC_index].strip()
            HomeCC = item[HomeCC_index].strip()
            SCE = item[SCE_index].strip()
            NetworkNo = item[NetworkNo_index].strip()
            Description = item[Description_index].strip()
            Activity = item[Activity_index].strip()
            Supp1 = item[Supp1_index].strip()
            Supp2 = item[Supp2_index].strip()
            Supp3 = item[Supp3_index].strip()
            Ship = item[Ship_index].strip()
            EmpID = item[EmpID_index].strip()
            EmpName = item[EmpName_index].strip()
            PayrollCode  = item[PayrollCode_index].strip()
            Shift = item[Shift_index].strip()
            Source = item[Source_index].strip()
            HrsCode = item[HrsCode_index].strip()
            WorkDate = item[WorkDate_index].strip()
            WeekEndDate = item[WeekEndDate_index].strip()
            HoursValid = item[HoursValid_index].strip()
            HoursPaid = item[HoursPaid_index].strip()
            BWNamerunlist.append(mySAPNamerun(WorkCC, HomeCC, SCE, NetworkNo, Description, Activity, Supp1, Supp2, Supp3, Ship, EmpID, EmpName, PayrollCode, Shift, Source, HrsCode, WorkDate, WeekEndDate, HoursValid, HoursPaid))
    
    return BWNamerunlist

def NetworkInNamerun(index, list_args):
    """
    checks for the location of the description and network number in namerun
    """
    
    counter = 0
    for single_list in list_args:
        #if (not list_args[0] in list_args) and (counter < 200):
        if counter < 200:
            item = single_list.split(',')
            network_info = item[index].strip()
            
            if len(network_info) == 9:
                bool = True
            else:
                bool = False
            
            counter = counter + 1
            
    return bool

def loadETC(filepath):
    '''
    Loads file into instantation function
    '''
    
    filename = open(filepath, 'r')
    ETClist = []
    readfile = filename.readlines()
    
    if readfile[0] in readfile:
        item = readfile[0].split(',')
        contract_index = SearchTitle(item, "PROJECT")
        resp_index = SearchTitle(item, "RESP") 
        wbs_index = SearchTitle(item, "WBS ID")
        desc_index = SearchTitle(item, "DESCRIPTION")
        cec_index = SearchTitle(item, "CEC")
        perf_index = SearchTitle(item, "PERF")
        clin_index = SearchTitle(item, "CLIN")
        network_index = SearchTitle(item, "CHARGE")
        period_index = SearchTitle(item, "YYYYMM")
        bcws_index = SearchTitle(item, "BCWS HRS/UTS")
        bcwp_index = SearchTitle(item, "BCWP HRS/UTS")
        act_index = SearchTitle(item, "ACT HRS/UTS")
        etc_index = SearchTitle(item, "ETC HRS/UTS")
        elem_index = MatchTitle(item, "E")

    for line in readfile:
        if not line == readfile[0]:
            item = line.split(',')
            contract = item[contract_index].strip()
            resp = item[resp_index].strip()
            wbs = item[wbs_index].strip()
            desc = item[desc_index].strip()
            cec = item[cec_index].strip()
            perf = item[perf_index].strip()
            clin = item[clin_index].strip()
            network = item[network_index].strip()
            period = item[period_index].strip()
            bcws = item[bcws_index].strip()
            bcwp = item[bcwp_index].strip()
            act = item[act_index].strip()
            etc = item[etc_index].strip()
            elem = item[elem_index].strip()
            ETClist.append(ETCEntry(contract, resp, wbs, desc, cec, perf, clin, network, period, bcws, bcwp, act, etc, elem))
            
    return ETClist

def loadCAM(filepath):
    "FUNCTION TO LOAD CAMS; need to be in CSV format"
    
    myCAMs = CAMEntry
    
    filename = open(filepath,'r')
    CAMlist = []
    
    for CAMLine in filename:
        if not CAMLine.startswith('//'):
            CAMItem = CAMLine.split(',')
            CAMlist.append(myCAMs(CAMItem[0].strip(), CAMItem[1].strip(), CAMItem[2].strip(), CAMItem[3].strip()))
    
    return CAMlist

def loadEmployee(filepath):
    '''
    Function to load Employee CSV File
    '''

    myEmployees = EmployeeEntry
    
    filename = open(filepath,'r')
    Employeelist = []
    
    for EmployeeLine in filename:
        if not EmployeeLine.startswith('//'):
            EmployeeItem = EmployeeLine.split(',')
            Employeelist.append(myEmployees(EmployeeItem[0].strip(), EmployeeItem[1].strip(), EmployeeItem[2].strip(), EmployeeItem[3].strip(), EmployeeItem[4].strip(), EmployeeItem[5].strip()))
    
    return Employeelist

def loadNetworks(filepath):
    '''
    ' FUNCTION TO LOAD list of NETWORKS in csv format
    '''
    
    myNetworks = NetworkEntry
    
    filename = open(filepath,'r')
    Networklist = []
    
    for NetworkLine in filename:
        if not NetworkLine.startswith('//'):
            NetworkItem = NetworkLine.split(',')
            Networklist.append(myNetworks(NetworkItem[0].strip(), NetworkItem[1].strip(), NetworkItem[2].strip(), NetworkItem[3].strip(), NetworkItem[4].strip(), NetworkItem[5].strip(), NetworkItem[6].strip(), NetworkItem[7].strip(), NetworkItem[8].strip()))
    
    return Networklist

def loadWeeklyCalendar(filepath):
    '''
    function to load the monthly accounting calendar; need csv
    '''
    
    cal = WeeklyCalendarEntry
    filename = open(filepath,'r')
    listCal = []
    
    for line in filename:
        if not line.startswith('//'):
            linevalues = line.split(',')
            listCal.append(cal(linevalues[0].strip(),linevalues[1].strip(),linevalues[2].strip(),linevalues[3].strip()))
        
    return listCal

def loadMonthlyCalendar(filepath):
    '''
    function to load the monthly accounting calendar; need csv
    '''
    cal = MonthlyCalendarEntry
    filename = open(filepath,'r')
    listCal = []
    
    for line in filename:
        if not line.startswith('//'):
            linevalues = line.split(',')
            listCal.append(cal(linevalues[0].strip(),linevalues[1].strip(),linevalues[2].strip()))
        
    return listCal

def loadContract(filepath):
    '''
    ' FUNCTION TO LOAD CONTRACTS; need to be in CSV format"
    '''
         
    myContracts = ContractEntry
    
    filename = open(filepath,'r')
    Contractlist = []
    
    for ContractLine in filename:
        if not ContractLine.startswith('//'):
            ContractItem = ContractLine.split(',')
            Contractlist.append(myContracts(ContractItem[0].strip(), ContractItem[1].strip()))
    
    return Contractlist

def MatchTitle(lMatch,Title):
    '''
    returns the position of the title in the list of the CSV file
    '''
    
    ListTitle = []
    
    for item in lMatch:
        ListTitle.append(item.strip())
    
    index = 0
    for item in ListTitle:
        
        if item == Title:
            return index
        index = index + 1
    
    return 0
        
def SearchTitle(lMatch,Title):
    '''
    returns the position of the title in the list of the CSV file
    '''
    
    ListTitle = []
    
    for item in lMatch:
        ListTitle.append(item.strip())
    
    index = 0
    for item in ListTitle:
        regex = re.compile(Title, re.I)
        findre = regex.search(item)
        
        if not findre == None:
            return index
        index = index + 1
    
    return 0

