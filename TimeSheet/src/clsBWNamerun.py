'''
Created on Feb 21, 2012

@author: Joselle Abagat
'''

import csv, re, os, shutil, collections
import win32com.client as win32
import xlwt
from clsCAM import *
from clsContract import *
from clsEmployee import *
from clsETC import *
from clsMonthlyCalendar import *
from clsNetwork import *
from clsWeeklyCalendar import *

class Nameruns(object):
    """ This class is dependent on the instances of the following resource classes: clsETC ==> ETC(),
                                                                                    clsMonthlyCalendar ==> MonthlyCalendar()
                                                                                    clsWeeklyCalendar ==> WeeklyCalendar()
                                                                                    clsEmployee ==> Employees()
    """
    
    def __init__(self, MonthlyCalendar, WeeklyCalendar, Employees, ETC):

        #LOAD RESOURCES
        self.mcal = MonthlyCalendar
        self.wcal = WeeklyCalendar
        self.emp = Employees
        self.mpm = ETC
        
        self.namerun_list = [] #loaded by: self.parseFile
        self.namerun_listItems = [] #loaded by: self.convertToListItems()
        self.set_NetworkActivityEmployeeHrsCode = [] #loaded by: self.parseNANHW()
        
        self.dict_NetworkToDescription = {} 
        self.dict_NetworkToActivities = {} #loaded by: self.mapNetworkToActivity()
        self.dict_NetworkToPeriod = {} #loaded by: self.mapNetworkToPeriod()
        self.dict_NetworkTotalHoursByPeriod = {} #loaded by: self.mapNetworkTotalHoursByPeriod()
        self.dict_NetworkToWeekEndDate = {} #loaded by: self.mapNetworkToWeekEndDate()
        self.dict_NetworkTotalHoursByWeekEndDate = {} #loaded by: self.mapNetworkTotalHoursByWeekEndDate()
        self.dict_NetworkActivityEmployeeHrsCodeToHours = {} #loaded by: self.mapNetworkActivityEmployeeHrsCodeWeekEndingToHours()
        self.dict_EmployeeNetworkActivityHrsCodeToHours = {} #loaded by: self.mapEmployeeNetworkActivityHrsCodeWeekEndingToHours()
        
        self.list_NetworkActivityEmployeeHrsCode = [] #loaded by: self.mapNetworkActivityEmployeeHrsCodeWeekEndingToHours()
        self.list_Period = [] #loaded by: self.listAccountingMonths()
        self.list_Activity = [] #loaded by: self.listActivityCodes()
        self.list_ChargeComments = [] #loaded by: self.listChargeComments()

        
    def parseFile(self, file_name):
        '''
        ' FUNCTION TO LOAD SAP Namerun; need to be in CSV format
        '''
        
        filename = open(file_name,'r')
        readfile = filename.readlines()
        
        readfile = self.replaceItem(readfile, ',', '')
#        readfile = self.replaceItem(readfile, '', ' ')
        readfile = self.replaceItem(readfile, '"', '')
        readfile =  self.replaceItem(readfile, '\n', '')

                
        file_header = readfile[0]
        file_header = file_header.split("|")

        WorkCC_index = self.SearchTitle(file_header, "Work Cost Center")
        HomeCC_index = self.SearchTitle(file_header, "Home Cost Center")
        SCE_index = self.SearchTitle(file_header, "SCE")
        Activity_index = self.SearchTitle(file_header, "Activity")
        Supp1_index = self.SearchTitle(file_header, "Supplemental 1")
        Supp2_index = self.SearchTitle(file_header, "Supplemental 2")
        Supp3_index = self.SearchTitle(file_header, "Supplemental 3")
        Ship_index = self.SearchTitle(file_header, "Ship")
        PayrollCode_index = self.SearchTitle(file_header, "Payroll Code")
        Shift_index = self.SearchTitle(file_header, "Shift")
        Source_index = self.SearchTitle(file_header, "Source")
        HrsCode_index = self.SearchTitle(file_header, "Hours Code")
        WorkDate_index = self.SearchTitle(file_header, "Work Date")
        WeekEndDate_index = self.SearchTitle(file_header, "Weekending Date")
        HoursValid_index = self.SearchTitle(file_header, "Hours Valid")
        HoursPaid_index = self.SearchTitle(file_header, "Hours Paid")
#        print(str(WorkCC_index) + '\n' + str(HomeCC_index) + '\n' + str(SCE_index) + '\n' + str(Activity_index) + '\n' + 
#              str(Supp1_index) + '\n' + str(Supp2_index) + '\n' + str(Supp3_index) + '\n' + str(Ship_index) + '\n' + 
#              str(PayrollCode_index) + '\n' + str(Shift_index) + '\n' + str(Source_index))


#            item = readfile[0].split(',')
#            WorkCC_index = self.SearchTitle(item, "Work")
#            HomeCC_index = self.SearchTitle(item, "Home")
#            SCE_index = self.SearchTitle(item, "SCE")
#            Activity_index = self.SearchTitle(item, "Activity")
#            Supp1_index = self.SearchTitle(item, "Supplemental 1")
#            Supp2_index = self.SearchTitle(item, "Supplemental 2")
#            Supp3_index = self.SearchTitle(item, "Supplemental 3")
#            Ship_index = self.SearchTitle(item, "Ship")
#            PayrollCode_index = self.SearchTitle(item, "Payroll Code")
#            Shift_index = self.SearchTitle(item, "Shift")
#            Source_index = self.SearchTitle(item, "Source")
#            HrsCode_index = self.SearchTitle(item, "Hours Code")
#            WorkDate_index = self.SearchTitle(item, "Work Date")
#            WeekEndDate_index = self.SearchTitle(item, "Weekending Date")
#            HoursValid_index = self.SearchTitle(item, "Hours Valid")
#            HoursPaid_index = self.SearchTitle(item, "Hours Paid")
            
        ###############
        # NOT REAL INDEX VALUES
        ###############
        emp_index = self.SearchTitle(file_header, "Employee")
        network_index = self.SearchTitle(file_header, "Network")

        headersList = [WorkCC_index, HomeCC_index, SCE_index, Activity_index, Supp1_index, Supp2_index, Supp3_index, Ship_index, PayrollCode_index, Shift_index, Source_index, emp_index, network_index]
        if self.checkEmptyHeaders(headersList) == True:
            return False
#        print(str(WorkCC_index) + '\ ' + str(HomeCC_index) + '\ ' + str(SCE_index) + '\ ' + str(Activity_index) + '\ ' + 
#              str(Supp1_index) + '\ ' + str(Supp2_index) + '\ ' + str(Supp3_index) + '\ ' + str(Ship_index) + '\ ' + 
#              str(PayrollCode_index) + '\ ' + str(Shift_index) + '\ ' + str(Source_index) + '\ ' + str(emp_index) + '\ ' + str(network_index))
        
        #CHECK FIRST IF FILEHEADER CONTAINS 2 'EMPLOYEE 'and 2 'NETWORK'
        itemCount = collections.Counter(file_header)
        
        if itemCount["Employee"] == 2: #this should always be true
            try:
                emp_info = int(file_header[emp_index].strip())
                EmpID_index = emp_index + 1
                EmpName_index = emp_index
            except ValueError:
                EmpName_index = emp_index +1
                EmpID_index = emp_index
        else:
                EmpID_index = emp_index
                EmpName_index = None
        
        if itemCount["Network"] == 2:
            if self.checkNetworkInNamerun(network_index, readfile) == True:
                NetworkNo_index = network_index
                Description_index = network_index + 1
            else:
                NetworkNo_index = network_index + 1
                Description_index = network_index           
        else:
            NetworkNo_index = network_index
            Description_index = None           

        for item in readfile[1:]:
            item = item.split("|")
            
            if len(item) <> 1:
                WorkCC = self.FindInList(item, WorkCC_index)
                HomeCC = self.FindInList(item, HomeCC_index)
                SCE = self.FindInList(item, SCE_index)
                NetworkNo = self.FindInList(item, NetworkNo_index)
                #print NetworkNo
                Description = self.FindInList(item, Description_index)
                Activity = self.FindInList(item, Activity_index)
                Supp1 = self.FindInList(item, Supp1_index)
                Supp2 = self.FindInList(item, Supp2_index)
                Supp3 = self.FindInList(item, Supp3_index)
                Ship = self.FindInList(item, Ship_index)
                EmpID = self.FindInList(item, EmpID_index)
                EmpName = self.FindInList(item, EmpName_index)
                PayrollCode = self.FindInList(item, PayrollCode_index)
                Shift = self.FindInList(item, Shift_index)
                Source = self.FindInList(item, Source_index)
                HrsCode = self.FindInList(item, HrsCode_index)
                WorkDate = self.FindInList(item, WorkDate_index)
                WeekEndDate = self.FindInList(item, WeekEndDate_index)
                HoursValid = self.FindInList(item, HoursValid_index)
                HoursPaid = self.FindInList(item, HoursPaid_index)
                
                #APPEND REMAINING ITEMS HERE: NETWORK PROGRAM, NETWORK CAM, EMPLOYEE CAM, SUSPICIOUS CHARGE, ACCOUNTING MONTH
                AccountingMonth = self.wcal.getPeriodFromWeekEndDate(WeekEndDate)
                NetworkProgram = self.mpm.getContractFromNetwork(NetworkNo)
                NetworkWBS = self.mpm.getWBSFromNetwork(NetworkNo)
                NetworkCAM = self.mpm.getCAMsFromNetwork(NetworkNo)
                POP = self.mpm.getPOPfromNetwork(NetworkNo)
                EmployeeCAM = self.emp.getCAMfromEmployeeNumber(EmpID, EmpName)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           
                EmployeeFHR = self.emp.getFHRfromEmployeeNumber(EmpID, HomeCC) 
                Status = self.getStatus(AccountingMonth, POP)
                ChargeComment = ""
                PossibleActivities = ""
                ETC = 0
                PercentSpent = 0
                TotalHours = 0
                
                if EmpID != "":                 
                    self.namerun_list.append(Namerun(WorkCC, HomeCC, SCE, NetworkNo, Description, Activity, Supp1, Supp2, Supp3, Ship, EmpID, EmpName, PayrollCode, Shift, Source, HrsCode, WorkDate, WeekEndDate, HoursValid, HoursPaid, AccountingMonth, NetworkProgram, NetworkWBS, NetworkCAM, EmployeeCAM, POP, Status, ChargeComment, PossibleActivities, ETC, PercentSpent, TotalHours))
#                print(WorkCC + '\n' + HomeCC + '\n' + SCE + '\n' + NetworkNo + '\n' + Description + '\n' + 
#                      Activity + '\n' + Supp1 + '\n' + Supp2 + '\n' + Supp3 + '\n' + Ship + '\n' + EmpID + 
#                      '\n' + EmpName + '\n' + PayrollCode + '\n' + Shift + '\n' + Source + '\n' + HrsCode + 
#                      '\n' + WorkDate + '\n' + WeekEndDate + '\n' + HoursValid + '\n' + HoursPaid)    
            
        filename.close()
    
    def checkEmptyHeaders(self, listarg):
        '''checks if any of the headers are found'''
        if all(each is None for each in listarg):
            return True
    
    def parseNANHW(self):
        '''creates a unique set of lists with class attributes
            listarg order: [network, activity, name, ID, hrs code, weekending date]
            NANHW order: [network, activity, name, ID, hrs code, weekending date]
            append order [0,1,2,3,4,5]
        '''
        for eachlist in self.uniqueSet(self.list_NetworkActivityEmployeeHrsCode):
            self.set_NetworkActivityEmployeeHrsCode.append(NANHW(eachlist[0],eachlist[1],eachlist[2],eachlist[3],eachlist[4],eachlist[5]))

    def convertToListItems(self):
        for each in self.namerun_list:
            cList = each.listConvert()
            if cList == False:
                print "Item skipped.  Cannot append to Namerun List" + str(each) + "\n" 
                #return False
            elif cList != False:
                self.namerun_listItems.append(cList)
    
    def FindInList(self, listArgument, index):
        """
            method to strip a list and assign an element of the list to a variable; 
            if it's a type error, it will return ""
        """
        
        try:
            return listArgument[index].strip()
        except TypeError:
            return ""
        except IndexError:
            return ""
    
    def replaceItem(self, data, replaceThis, replaceWith):
        """
        uses map to edit each line in the data
        """
        
        return map(lambda line: line.replace(replaceThis, replaceWith), data)
    
    def checkEmployeeInNamerun(self, index, list_args):
        """
        checks for the location of the employee name and employee id in namerun
        """
        pass
    
    def checkNetworkInNamerun(self, index, list_args):
        """
        checks for the location of the description and network number in namerun
        1. loop through the first 200 arguments
        2. create a boolean counter.  return which one is the greater number
        """
        
        trueCount = 0
        falseCount = 0
        
        for single_list in list_args[1:200]:
            if single_list != "\n":
                network_info = single_list.split("|")
                
                try:
                    if len(network_info[index]) == 9:
                        trueCount = trueCount + 1
                    else:
                        falseCount = falseCount + 1
                except IndexError:
                    print "list: " + network_info
                    print "index: " + index
                
        if trueCount > falseCount:
            return True
        else:
            return False
    
    def MatchTitle(self, lMatch, Title):
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
        
        return None
            
    def SearchTitle(self, lMatch, Title):
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
        
        return None
    
    def listAccountingMonths(self):
        """lists accounting months / periods present in the BW namerun"""
        for each_namerun in self.namerun_list:
            if str(each_namerun.AccountingMonth) not in self.list_Period:
                self.list_Period.append(str(each_namerun.AccountingMonth))
        self.list_Period.sort()
    
    def listActivityCodes(self):
        """lists activity codes present in the BW Namerun"""
        for each_namerun in self.namerun_list:
            if each_namerun.Activity not in self.list_Activity:
                self.list_Activity.append(each_namerun.Activity)
        self.list_Activity.sort()
    
    def listChargeComments(self):
        """lists charge comments present in the BW namerun"""
        for each_namerun in self.namerun_list:
            if each_namerun.ChargeComment not in self.list_ChargeComments:
                self.list_ChargeComments.append(each_namerun.ChargeComment)
                
    def mapNetworkToDescription(self):
        """ creates relationship between network and description """
        for each_namerun in self.namerun_list:
            if each_namerun.NetworkNo not in self.dict_NetworkToDescription:
                self.dict_NetworkToDescription[each_namerun.NetworkNo] = each_namerun.Description
    
    # we want to use this to update the network list file.. 
    # but how to determine which activity goes to what CAM?
    def mapNetworkToActivity(self):
        """
        returns a dictionary of the form {network#: [activity1, activity2, ...], ..., network#n: [activity1, activity2, ...]}
        """
        
        #input keys with an empty list in the dictionary        
        for each_namerun in self.namerun_list:
            if not each_namerun.NetworkNo in self.dict_NetworkToActivities:
                self.dict_NetworkToActivities[each_namerun.NetworkNo] = []
        
        #append all activity codes to their respective keys
        for each_namerun in self.namerun_list:
            self.dict_NetworkToActivities[each_namerun.NetworkNo].append(each_namerun.Activity)
        
        #remove duplicates within the list of each key
        for each_key in self.dict_NetworkToActivities:
            self.dict_NetworkToActivities[each_key] = self.uniqueList(self.dict_NetworkToActivities[each_key])
        
    def mapNetworkToPeriod(self):
        """
        returns a dictionary of the form {network#1: [period1,
                                                      ..., 
                                                      periodn],
                                          network#n: [period1,
                                                      ..., 
                                                      periodn]}
        """
        
        for each_entry in self.namerun_list:

            if not each_entry.NetworkNo in self.dict_NetworkToPeriod:
                self.dict_NetworkToPeriod[each_entry.NetworkNo] = [each_entry.AccountingMonth]
            else:
                self.dict_NetworkToPeriod[each_entry.NetworkNo].append(each_entry.AccountingMonth)
        
        #make periods in each network list unique
    
        
    def mapNetworkTotalHoursByPeriod(self):
        """
        Network Total Hours = fn(network #, period)
        get period from WeeklyCalendar Class
        dictionary will come in the form: {network#1: {period1: [hrs1, hrs2],
                                                      period2: [hrs1, hrs2]},
                                           network#2: {period1: [hrs1, hrs2],
                                                      period2: [hrs1, hrs2]},}
        """
        
        for each_entry in self.namerun_list:
            if not each_entry.NetworkNo in self.dict_NetworkTotalHoursByPeriod:
                self.dict_NetworkTotalHoursByPeriod[each_entry.NetworkNo] = {each_entry.AccountingMonth: []}
            else:
                self.dict_NetworkTotalHoursByPeriod[each_entry.NetworkNo][each_entry.AccountingMonth] = []        
      
        for each_entry in self.namerun_list:            
            self.dict_NetworkTotalHoursByPeriod[each_entry.NetworkNo][each_entry.AccountingMonth].append(float(each_entry.HoursValid))
    
    def mapNetworkToWeekEndDate(self):
        """
        returns a dictionary of the form {network#1: {WeekEndDate1: [],
                                                      ..., 
                                                      WeekEndDaten: []},
                                                      ..., 
                                          network#n: {WeekEndDate1: [],
                                                      ..., 
                                                      WeekEndDaten: []}}
        """
        
        for each_entry in self.namerun_list:

            if not each_entry.NetworkNo in self.dict_NetworkToWeekEndDate:
                self.dict_NetworkToWeekEndDate[each_entry.NetworkNo] = {each_entry.WeekEndDate: []}
            else:
                self.dict_NetworkToWeekEndDate[each_entry.NetworkNo][each_entry.WeekEndDate] = []
    
    def mapNetworkTotalHoursByWeekEndDate(self):
        """
        Network Total Hours = fn(network #, Weekend Date)
        get period from WeeklyCalendar Class
        dictionary will come in the form: {network#1: {WeekEndDate1: [hrs1, hrs2],
                                                      WeekEndDate: [hrs1, hrs2]},
                                           network#2: {WeekEndDate: [hrs1, hrs2],
                                                      WeekEndDate: [hrs1, hrs2]},}
        """
        
        if len(self.dict_NetworkTotalHoursByWeekEndDate) == 0:
            self.dict_NetworkTotalHoursByWeekEndDate = self.dict_NetworkToWeekEndDate
        
        for each_entry in self.namerun_list:
            
            if self.dict_NetworkTotalHoursByWeekEndDate[each_entry.NetworkNo][each_entry.WeekEndDate] == []:
                self.dict_NetworkTotalHoursByWeekEndDate[each_entry.NetworkNo][each_entry.WeekEndDate] = [each_entry.HoursValid]
            else:
                self.dict_NetworkTotalHoursByWeekEndDate[each_entry.NetworkNo][each_entry.WeekEndDate].append(each_entry.HoursValid)
    
    def mapNetworkActivityEmployeeHrsCodeWeekEndingToHours(self):
        """ builds a dictionary of the form:
            {Network1: {Activity1: {Employee1:    {Hrs_Code1:    {week1: [hrs1, hrs2],
                                                                  week2: [hrs1, hrs2],
                                                                  week3: [hrs1, hrs2],
                                                                  week4: [hrs1, hrs2]
                                                                  },
                                                   Hrs_Code2:    {week1: [hrs1, hrs2],
                                                                  week2: [hrs1, hrs2],
                                                                  week3: [hrs1, hrs2],
                                                                  week4: [hrs1, hrs2]
                                                                  }
        """
        
        for each in self.namerun_list:
            #fill first hierarchy of dictionary with network numbers
            if each.NetworkNo not in self.dict_NetworkActivityEmployeeHrsCodeToHours:
                self.dict_NetworkActivityEmployeeHrsCodeToHours[each.NetworkNo] = {}
            #get values of network, activity code, name, hours code, and weekend date and append to list
            self.list_NetworkActivityEmployeeHrsCode.append(Namerun.nanhw(each))
                        
        #second loop to fill 2nd hierarchy of dictionary with activity codes
        for each in self.namerun_list:        
            if each.Activity not in self.dict_NetworkActivityEmployeeHrsCodeToHours[each.NetworkNo]:
                self.dict_NetworkActivityEmployeeHrsCodeToHours[each.NetworkNo][each.Activity] = {}
        
        #third loop to fill 3rd hierarchy of dictionary with employee names
        for each in self.namerun_list:    
            if each.EmpName not in self.dict_NetworkActivityEmployeeHrsCodeToHours[each.NetworkNo][each.Activity]:
                self.dict_NetworkActivityEmployeeHrsCodeToHours[each.NetworkNo][each.Activity][each.EmpName] = {}
        
        #fourth loop to fill 4th hierarchy with hours code
        for each in self.namerun_list:                
            if each.HrsCode not in self.dict_NetworkActivityEmployeeHrsCodeToHours[each.NetworkNo][each.Activity][each.EmpName]:
                self.dict_NetworkActivityEmployeeHrsCodeToHours[each.NetworkNo][each.Activity][each.EmpName][each.HrsCode] = {}
                
        #fifth loop to fill 5th level with week-ending dates
        for each in self.namerun_list:                
            if each.WeekEndDate not in self.dict_NetworkActivityEmployeeHrsCodeToHours[each.NetworkNo][each.Activity][each.EmpName][each.HrsCode]:
                self.dict_NetworkActivityEmployeeHrsCodeToHours[each.NetworkNo][each.Activity][each.EmpName][each.HrsCode][each.WeekEndDate] = []
        
        #sixth loop for list of valid hours; convert to floating point for data manipulation
        for each in self.namerun_list:                                    
            if self.dict_NetworkActivityEmployeeHrsCodeToHours[each.NetworkNo][each.Activity][each.EmpName][each.HrsCode][each.WeekEndDate] == []:
                self.dict_NetworkActivityEmployeeHrsCodeToHours[each.NetworkNo][each.Activity][each.EmpName][each.HrsCode][each.WeekEndDate] = [float(each.HoursValid)]                                    
            else:
                self.dict_NetworkActivityEmployeeHrsCodeToHours[each.NetworkNo][each.Activity][each.EmpName][each.HrsCode][each.WeekEndDate].append(float(each.HoursValid))
        
    def mapEmployeeNetworkActivityHrsCodeWeekEndingToHours(self):
        """ builds a dictionary of the form:
            {Employee1: {Network1: {Activity1:    {Hrs_Code1:    {week1: [hrs1, hrs2],
                                                                  week2: [hrs1, hrs2],
                                                                  week3: [hrs1, hrs2],
                                                                  week4: [hrs1, hrs2]
                                                                  },
                                                   Hrs_Code2:    {week1: [hrs1, hrs2],
                                                                  week2: [hrs1, hrs2],
                                                                  week3: [hrs1, hrs2],
                                                                  week4: [hrs1, hrs2]
                                                                  }
        """        
        
        for each in self.namerun_list:
            if each.EmpName not in self.dict_EmployeeNetworkActivityHrsCodeToHours:
                self.dict_EmployeeNetworkActivityHrsCodeToHours[each.EmpName] = {}
        
        for each in self.namerun_list:
            if each.NetworkNo not in self.dict_EmployeeNetworkActivityHrsCodeToHours[each.EmpName]:
                self.dict_EmployeeNetworkActivityHrsCodeToHours[each.EmpName][each.NetworkNo] = {}
                
        for each in self.namerun_list:
            if each.Activity not in self.dict_EmployeeNetworkActivityHrsCodeToHours[each.EmpName][each.NetworkNo]:
                self.dict_EmployeeNetworkActivityHrsCodeToHours[each.EmpName][each.NetworkNo][each.Activity] = {}
        
        for each in self.namerun_list:
            if each.HrsCode not in self.dict_EmployeeNetworkActivityHrsCodeToHours[each.EmpName][each.NetworkNo][each.Activity]:
                self.dict_EmployeeNetworkActivityHrsCodeToHours[each.EmpName][each.NetworkNo][each.Activity][each.HrsCode] = {}
        
        for each in self.namerun_list:
            if each.WeekEndDate not in self.dict_EmployeeNetworkActivityHrsCodeToHours[each.EmpName][each.NetworkNo][each.Activity][each.HrsCode]:
                self.dict_EmployeeNetworkActivityHrsCodeToHours[each.EmpName][each.NetworkNo][each.Activity][each.HrsCode][each.WeekEndDate] = []
        
        for each in self.namerun_list:
            self.dict_EmployeeNetworkActivityHrsCodeToHours[each.EmpName][each.NetworkNo][each.Activity][each.HrsCode][each.WeekEndDate].append(float(each.HoursValid))
    
    def getDescriptionFromNetwork(self, Network):
        return self.dict_NetworkToDescription[Network]
    
    def getNetworkTotalHoursByPeriod(self, network_number, period):
        """
        Return the total hours for a network number for the given accounting calendar month
        """
        try:
            return sum(self.dict_NetworkTotalHoursByPeriod[network_number][period])
        except KeyError:
            print("Cannot calculate total hours for " + network_number + " " + period + ". Check sap data")
        except TypeError:
            return 0
    
    def getNetworkTotalHoursByWeekEndDate(self, network_number, Weekending_Date):
        """
        Return the total hours for a network number for the given accounting calendar month
        """
        try:
            return sum(self.dict_NetworkTotalHoursByPeriod[network_number][Weekending_Date])
        except KeyError:
            print("Cannot calculate total hours for " + network_number + " " + Weekending_Date + ". Check sap data")
        except TypeError:
            return 0
    
    def getNetworkActivityEmployeeHrsCodeWeekEndingToHours(self, Network, Activity, EmpName, HrsCode, WeekEnd):
        return sum(self.dict_NetworkActivityEmployeeHrsCodeToHours[Network][Activity][EmpName][HrsCode][WeekEnd])
    
    def getEmployeeNetworkActivityHrsCodeWeekEndingToHours(self, EmpName, Network, Activity, HrsCode, WeekEnd):
        return sum(self.dict_EmployeeNetworkActivityHrsCodeToHours[EmpName][Network][Activity][HrsCode][WeekEnd])
    
    def getStatus(self, CurPeriod, EndPeriod):
        """ determine's a network's status"""
        
        try:
            if int(CurPeriod) > int(EndPeriod):
                return "Closed"
            else:
                return "Open"
        except TypeError:
            return "Unknown"
        except ValueError:
            return "Unknown"
    
    def htmlHeaderAllUp(self):
        h = "<tr>" + \
            "<th>Program</th>" + \
            "<th>Network</th>" + \
            "<th>Description</th>" + \
            "<th>Activity</th>" + \
            "<th>Employee's CAM</th>" + \
            "<th>Employee</th>" + \
            "<th>Hours Code</th>" + \
            "<th>Week 1</th>" + \
            "<th>Week 2</th>" + \
            "<th>Week 3</th>" + \
            "<th>Week 4</th>" + \
            "<th>Week 5</th>" + \
            "<th>Total</th>" + \
            "<th>Network ETC</th>" + \
            "<th>% Spent</th>" + \
            "<th>Network Status</th>" + \
            "<th>Charge Comment</th>" + \
            "</tr>"
        return h
    
#    def htmlReportAllUp(self):
#        """ all up report summary"""
#        
#        for each in self.list_NetworkActivityEmployeeHrsCode:
#            Program = self.mpm.getContractFromNetwork(each.NetworkNo)
#            Network = each.NetworkNo
#            Description = self.getDescriptionFromNetwork(each.NetworkNo)
#            Activity = each.Activity
#            EmpCAM = self.emp.getCAMfromEmployeeNumber(employee_number, employee_name)
#            Name = each.EmpName
    
    def toExcelAllUp(self, fullpath):
        '''writes each Namerun in xlsx format; do not use since it's too long'''
        
        excel = win32.Dispatch('Excel.Application')
        excel.DisplayAlerts = False        
        
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets.Add()
        ws.Name = "Namerun"

        
        header = self.namerun_list[0].header()
        headerList = header.split(",")
        
        #write headers
        for column in range(len(headerList)):
            ws.Cells(1,column+1).Value = headerList[column]
        
        
        #write data
        for row in range(len(self.namerun_listItems)):
            
            for column in range(len(headerList)):
                ws.Cells(row+2,column+1).Value = self.namerun_listItems[row][column]

        wb.SaveAs(fullpath)
        excel.Workbooks.Open(fullpath)
        excel.Visible = True
        excel.DisplayAlerts = True
        
    def toCSV(self):
        ''' writes each Namerun in csv format '''

                 
        s = ""
        
        for each in self.namerun_list:
            s += each.toCSV()
        
        h = self.namerun_list[0].header()
        
        return  h + s
    
    def toCSVFile(self, filename):
        f = open(filename, 'w')
        f.write(self.toCSV())         

    def updateChargeComments(self):
        """ auto comment checks:
                - closed network = mischarge
                - more than 0030 activities = needs a check of activity codes
                - if not within list of cams, check where employee should belong to
        """

        for each in self.namerun_list:
            if each.Status != 'Closed':
                if str(each.Activity) == '0030' or str(each.Activity) == '30':
                    #technically, there shouldn't be a key error since the dictionary was created from the namerun list, but just in case use an exception
                    try:
                        if len(self.dict_NetworkToActivities[each.NetworkNo]) > 1:
#                            print str(self.dict_NetworkToActivities[each.NetworkNo]) + " = " + str(len(self.dict_NetworkToActivities[each.NetworkNo]))

                            each.ChargeComment = "Check Activity Code"
                            
                            strAct = ""
                            for act in self.dict_NetworkToActivities[each.NetworkNo]:
                                strAct += str(act) + "; "
                            each.PossibleActivities = strAct
                        
                        else: 
                            if len(each.EmployeeCAM) > 3:
                                each.ChargeComment = "No CAM Code"
                            else:                               
                                if each.EmployeeCAM[:2] != each.NetworkCAM[:2]:
                                    networkCAMs = each.NetworkCAM.split(";")
                                    for cam in networkCAMs:
                                        cam = cam.replace(" ","")
                                        if each.EmployeeCAM[:2] != cam[:2]:
                                            each.ChargeComment = "Employee belongs to " + each.EmployeeCAM
                                        else:
                                            each.ChargeComment = "Ok"
                                else:
                                    each.ChargeComment = "Ok"
                    except KeyError:
                        each.ChargeComment = each.NetworkNo
                else:
                    if len(each.EmployeeCAM) > 3:
                        each.ChargeComment = "No CAM Code"
                    else:
                        if each.EmployeeCAM[:2] != each.NetworkCAM[:2]:
                            networkCAMs = each.NetworkCAM.split(";")
                            for cam in networkCAMs:
                                cam = cam.replace(" ","")
                                if each.EmployeeCAM[:2] != cam[:2]:
                                    each.ChargeComment = "Employee belongs to " + each.EmployeeCAM
                                else:
                                    each.ChargeComment = "Ok"
                        else:
                            each.ChargeComment = "Ok"
            else:
                each.ChargeComment = "No ETC"
    
    def updateETC(self):
        """ changes the value of ETC in each namerun value """
        for each in self.namerun_list:            
            getSumETC = self.mpm.getETCFromNetworkCAMAndPeriod(each.NetworkNo, each.NetworkCAM, each.AccountingMonth)
            each.ETC = getSumETC
    
    def updateTotalHours(self):
        """ updates the values of the Total hours in each namerun value"""
        for each in self.namerun_list:
            each.TotalHours = self.getNetworkTotalHoursByPeriod(each.NetworkNo, each.AccountingMonth)
        
    def updatePercentSpent(self):    
        """ calculates percent spent in each of the total value """
        for each in self.namerun_list:
            try:
                each.PercentSpent = each.TotalHours/each.ETC
            except ZeroDivisionError:
                each.PercentSpent = 0
            except TypeError:
                each.PercentSpent = 0
    
    def uniqueSet(self, listarg):
        '''removes duplicate lists from a list of lists'''
        return [list(sublist) for sublist in set(tuple(sublist) for sublist in listarg)]
    
    def uniqueList(self, listarg):
        '''removes duplicates from a list'''
        return list(set(listarg))
    
    def SumNamerun(self, sap_report):
        """
        Method to sum up the hours if the following are the same: Charge#, Act, Desc, EmpID, Payroll Code, WeekEnd Date
        SAP Report looks like:
                       [n][0]    [n][1]  [n][2]    ...                               [n][7]
                [
             0       [Charge#_1, Act_1, Desc_1, EmpID_1, Name_1, Payroll_1, Date_1, Hours_1]
             1       [Charge#_2, Act_2, Desc_2, EmpID_2, Name_2, Payroll_2, Date_2, Hours_2]
                    .
                    .
                    .
             n       [Charge#_n, Act_n, Desc_n, EmpID_n, Name_n, Payroll_n, Date_n, Hours_n]
                ]
        """
            
        sum_list = []
        
        # loop through the report in order to look at each row of data
        for each_list in sap_report:
            #convert all hours into floats
            each_list[-1] = float(each_list[-1])
            # take out all the elements of that row except for the last one since those are the ones we want to compare
            same_elems = each_list[0:(len(each_list) - 1)]
            # find those list elements in sum_list; if it is found in_sum_list != []
            in_sum_list = [findlist for findlist in sum_list if findlist[0:len(same_elems)] == same_elems]
            
            # if those items are already in sum_list, then we just want to add the last to the last value
            # if these items are not in sum_list, then add the whole row into the list
            try:
                find_index = sum_list.index(in_sum_list[0])
                sum_list[find_index][-1] = sum_list[find_index][-1] + each_list[-1]
            except IndexError:
                sum_list.append(each_list)
                
        return sum_list
    
    def get_total_hours(self):
        total = 0
        for nameRunEntry in self.name_run_list:
            total += nameRunEntry.HoursValid
        return total
    
    def reportSuspiciousCharging(self):
        
        suspiciousCharges = []
        nameRunEntry = Namerun()
        for nameRunEntry in self.name_run_list:
            if nameRunEntry.HoursCode == "EWW":
                suspiciousCharges.append( (nameRunEntry, "This person is on EWW") )
            
            if nameRunEntry.getNetworkCAM() != nameRunEntry.getOrganizationCAM():
                suspiciousCharges.append( (nameRunEntry, "This person belongs to " + nameRunEntry.getNetworkCAM() + " but charged to a network " + nameRunEntry.NetworkNo + " belonging to " + nameRunEntry.getNetworkCAM()))
        
        return suspiciousCharges
        

#
#lrip = NameRun("lrip.txt")
#emd = NameRun("emd.txt")
#
#lrip.get_total_hours()
#emd.get_total_hours()
#
#lrip.reportSuspiciousCharging()

class Namerun(object):
    '''
    Class to for SAP Nameruns
    '''
    
    def __init__(self, WorkCC, HomeCC, SCE, NetworkNo, Description, Activity, Supp1, Supp2, Supp3, Ship, EmpID, EmpName, PayrollCode, Shift, Source, HrsCode, WorkDate, WeekEndDate, HoursValid, HoursPaid, AccountingMonth, NetworkProgram, NetworkWBS, NetworkCAM, EmployeeCAM, POP, NetworkStatus, ChargeComment, PossibleActivities, ETC, PercentSpent, TotalHours):
        '''
        Constructor: needs the following parameters:
            Work Cost Center, Home Cost Center, Summary Cost Element, Network Number, Description, \
            Activity Code, Supp1, Supp2, Supp3, Ship, Employee ID, Employee Name, Payroll Code, Shift, \
            Soruce, Hours Code, Work Date, Week Ending Date, Hours Valid, Hours Paid
        '''
        
        self.WorkCC = WorkCC
        self.HomeCC = HomeCC
        self.SCE = SCE
        self.NetworkNo = NetworkNo
        self.Description = Description
        self.Activity = Activity
        self.Supp1 = Supp1
        self.Supp2 = Supp2
        self.Supp3 = Supp3
        self.Ship = Ship
        self.EmpID = EmpID
        self.EmpName = EmpName
        self.PayrollCode = PayrollCode
        self.Shift = Shift
        self.Source = Source
        self.HrsCode = HrsCode
        self.WorkDate = WorkDate
        self.WeekEndDate = WeekEndDate
        self.HoursValid = HoursValid
        self.HoursPaid = HoursPaid
        
        #ADD PROGRAM, CAM-NETWORK, CAM-EMPLOYEE, Suspicious Charge (default is no)
        self.NetworkProgram = NetworkProgram
        self.NetworkWBS = NetworkWBS
        self.NetworkCAM = NetworkCAM
        self.EmployeeCAM = EmployeeCAM
        self.POP = POP
        self.AccountingMonth = AccountingMonth
        self.Status = NetworkStatus
        self.ChargeComment = ChargeComment
        self.PossibleActivities = PossibleActivities
        self.ETC = ETC
        self.PercentSpent = PercentSpent
        self.TotalHours = TotalHours
        
    def __str__(self):
        
        s = "\nNamerun:" + \
            "\n\tWork Cost Center: " + str(self.WorkCC) + \
            "\n\tHome Cost Center: " + str(self.HomeCC) + \
            "\n\tSummary Cost Element: " + str(self.SCE) + \
            "\n\tNetwork Program: " + str(self.NetworkProgram) + \
            "\n\tNetwork WBS: " + str(self.NetworkWBS) + \
            "\n\tNetwork CAM Code: " + str(self.NetworkCAM) + \
            "\n\tNetwork Status: " + str(self.Status) + \
            "\n\tCharge Number: " + str(self.NetworkNo) + \
            "\n\tDescription: " + str(self.Description) + \
            "\n\tActivity Code: " + str(self.Activity) + \
            "\n\tSupp1: " + str(self.Supp1) + \
            "\n\tSupp2: " + str(self.Supp2) + \
            "\n\tSupp3: " + str(self.Supp3) + \
            "\n\tShip: " + str(self.Ship) + \
            "\n\tEmployee ID: " + str(self.EmpID) + \
            "\n\tEmployee Name: " + str(self.EmpName) + \
            "\n\tEmployee CAM Code: " + str(self.EmployeeCAM) + \
            "\n\tPayroll Code: " + str(self.PayrollCode) + \
            "\n\tShift: " + str(self.Shift) + \
            "\n\tSource: " + str(self.Source) + \
            "\n\tHours Code: " + str(self.HrsCode) + \
            "\n\tWorkDate: " + str(self.WorkDate) + \
            "\n\tWeekending Date: " + str(self.WeekEndDate) + \
            "\n\tAccounting Period: " + str(self.AccountingMonth) + \
            "\n\tPOP Ends: " + str(self.POP) + \
            "\n\tHours Valid: " + str(self.HoursValid) + \
            "\n\tHours Paid: " + str(self.HoursPaid) + \
            "\n\tETC: " + str(self.ETC) + \
            "\n\tTotal Hours: " + str(self.TotalHours) + \
            "\n\tPercent Spent: " + str(self.PercentSpent) + \
            "\n\tComment: " + str(self.ChargeComment) + \
            "\n\tPossible Activities: " + str(self.PossibleActivities)
 
        return s
    
    def __eq__(self, test_case):
        
        if self.Activity != test_case.Activity:
            return False
        if self.Description != test_case.Description:
            return False
        if self.EmpID != test_case.EmpID:
            return False
        if self.EmpName != test_case.EmpName:
            return False
        if self.HomeCC != test_case.HomeCC:
            return False
        if self.HoursPaid != test_case.HoursPaid:
            return False
        if self.HoursValid != test_case.HoursValid:
            return False
        if self.HrsCode != test_case.HrsCode:
            return False
        if self.NetworkNo != test_case.NetworkNo:
            return False
        if self.PayrollCode != test_case.PayrollCode:
            return False
        if self.SCE != test_case.SCE:
            return False
        if self.Shift != test_case.Shift:
            return False
        if self.Ship != test_case.Ship:
            return False
        if self.Source != test_case.Source:
            return False
        if self.Supp1 != test_case.Supp1:
            return False
        if self.Supp2 != test_case.Supp2:
            return False
        if self.Supp3 != test_case.Supp3:
            return False
        if self.WeekEndDate != test_case.WeekEndDate:
            return False
        if self.WorkCC != test_case.WorkCC:
            return False
        if self.WorkDate != test_case.WorkDate:
            return False
        
        return True
    
    def listConvert(self):
        "returns a list version of the class"
        try:
            float(self.HoursPaid)
        except ValueError:
            self.HoursPaid = 0
        
        try:
            float(self.HoursValid)
        except ValueError:
            self.HoursValid = 0
        
        try: 
            int(self.EmpID)
        except ValueError:
            print "Name: " + self.EmpName
            print "ID: " + self.EmpID
            return False
        
        if int(self.AccountingMonth) == 0:
            print "No Period"    
            return False        
         
        if str(self.WorkDate) is "" or str(self.WeekEndDate) is "":
            print "No Date worked"
            return False
               
        if self.EmpID is "":
            print "EMPTY ID"
            return False

        return [str(self.WorkCC), str(self.HomeCC), str(self.SCE), str(self.NetworkProgram), str(self.NetworkWBS), str(self.NetworkCAM), str(self.POP),
                str(self.Status), str(self.NetworkNo), str(self.Description), str(self.Activity), str(self.Supp1), str(self.Supp2),
                str(self.Supp3), str(self.Ship), str(self.EmployeeCAM), int(self.EmpID), str(self.EmpName), str(self.PayrollCode),
                str(self.Shift), str(self.Source), str(self.HrsCode), str(self.WorkDate), str(self.WeekEndDate), int(self.AccountingMonth),
                float(self.HoursValid), float(self.HoursPaid), float(self.ETC), float(self.TotalHours), float(self.PercentSpent), str(self.ChargeComment), str(self.PossibleActivities)]
            
    def nanhw(self):
        ''' N = Network Number
            A = Activity Code
            N = Employee Name
            E = Employee ID
            H = Hours Code
            W = Week-Ending Date
            '''
        return [self.NetworkNo, self.Activity, self.EmpName, self.EmpID, self.HrsCode, self.WeekEndDate]
    
#    def nnahw(self):
#        ''' N = Network Number
#            N = Employee Name
#            A = Activity Code
#            H = Hours Code
#            W = Week-Ending Date
#            '''
#        return [self.EmpName, self.NetworkNo, self.Activity, self.HrsCode, self.WeekEndDate]
    
    def header(self):
        """creates the header"""
        
        s =  "Work Cost Center," + \
             "Home Cost Center," + \
             "Summary Cost Element," + \
             "Program," + \
             "WBS," + \
             "Network CAM Code (RESP)," + \
             "POP End," + \
             "Network Status," + \
             "Network," + \
             "Description," + \
             "Activity," + \
             "Supplemental 1," + \
             "Supplemental 2," + \
             "Supplemental 3," + \
             "Ship," + \
             "Employee CAM Code," + \
             "Employee ID," + \
             "Employee Name," + \
             "Payroll Code," + \
             "Shift," + \
             "Source," + \
             "Hours Code," + \
             "Work Date," + \
             "Weekend Date," + \
             "Accounting Month," + \
             "Hours Valid," + \
             "Hours Paid," + \
             "ETC," + \
             "TotalHours," + \
             "Percent Spent," + \
             "Charge Comment," + \
             "Possible Activities\n"
        return s
    
    def toCSV(self):
        '''creates csv format'''
        
        s = str(self.WorkCC) + ',' + \
            str(self.HomeCC) + ',' + \
            str(self.SCE) + ',' + \
            str(self.NetworkProgram) + ',' + \
            str(self.NetworkWBS) + ',' + \
            str(self.NetworkCAM) + ',' + \
            str(self.POP) + ',' + \
            str(self.Status) + ',' + \
            str(self.NetworkNo) + ',' + \
            str(self.Description) + ',' + \
            str(self.Activity) + ',' + \
            str(self.Supp1) + ',' + \
            str(self.Supp2) + ',' + \
            str(self.Supp3) + ',' + \
            str(self.Ship) + ',' + \
            str(self.EmployeeCAM) + ',' + \
            str(self.EmpID) + ',' + \
            str(self.EmpName) + ',' + \
            str(self.PayrollCode) + ',' + \
            str(self.Shift) + ',' + \
            str(self.Source) + ',' + \
            str(self.HrsCode) + ',' + \
            str(self.WorkDate) + ',' + \
            str(self.WeekEndDate) + ',' + \
            str(self.AccountingMonth) + ',' + \
            str(self.HoursValid) + ',' + \
            str(self.HoursPaid) + ',' + \
            str(self.ETC) + ',' + \
            str(self.TotalHours) + ',' + \
            str(self.PercentSpent) + ',' + \
            str(self.ChargeComment) + ',' + \
            str(self.PossibleActivities) + "\n"
        
        return s
    
class NANHW(object):
    ''' N = Network Number
        A = Activity Code
        N = Employee Name
        E = Employee ID
        H = Hours Code
        W = Week-Ending Date
    '''
    
    def __init__(self, NetworkNo, Activity, EmpName, ID, HrsCode, WeekEndDate):
        '''construct class'''
        self.NetworkNo = NetworkNo
        self.Activity = Activity
        self.EmpName = EmpName
        self.EmpID = ID
        self.HrsCode = HrsCode
        self.WeekEndDate = WeekEndDate
    
    def __str__(self):
        s = "\nCharge Number: " + str(self.NetworkNo) + \
            "\nActivity Code: " + str(self.Activity) + \
            "\nEmployee Name: " + str(self.EmpName) + \
            "\nEmployee ID: " + str(self.EmpID) + \
            "\nHours Code: " + str(self.HrsCode) + \
            "\nWeekending Date: " + str(self.WeekEndDate)
 
        return s


#    def NetworkActivityPair(self):
#        '''
#        Returns a List of the form [Network Number, Activity Code]
#        '''
#        
#        lNetworkActivity = [self.NetworkNo, self.Activity, self.NetworkNo + self.Activity]
#        
#        return lNetworkActivity
#    
#    def NetworkDescriptionPair(self):
#        '''
#        Returns a list with a Description of each network and their respective Activity Code of the form [Network Number, Activity Code, NetworkActivity, Description]
#        '''
#        
#        lNetworkDescription = [self.NetworkNo, self.Activity, self.Description]
#        
#        return lNetworkDescription
#    
#    def EmployeeInfo(self):
#        '''
#        Returns a list of the Employees and their IDs
#        '''
#        
#        lEmployee = [self.EmpID, self.EmpName]
#        
#        return lEmployee
#    
#    def ReportData(self):
#        '''
#        Returns a list of the form [Network #, Activity Code, Description, Employee ID, Employee Name, Payroll Code, WeekEnd Date, Hours Valid]
#        '''
#        
#        lReportData = [self.NetworkNo, self.Activity, self.Description, self.EmpID, self.EmpName, self.PayrollCode, self.WeekEndDate, self.HoursValid]
#        
#        return lReportData
#     
#    def CallNetworkActivity(self, loaded_file):
#        """
#        Function that returns a the mapping between network and activity code
#        """
#        lmapping = []
#        
#        for item in loaded_file:
#            if not self.NetworkActivityPair(item) in lmapping:
#                lmapping.append(self.NetworkActivityPair(item))
#        
#        return lmapping
#
#    def CallNetworkDescription(self, loaded_file):
#        """
#        Function that returns the mapping between a network and its description
#        """
#        lmapping = []
#        
#        for item in loaded_file:
#            if not self.NetworkDescriptionPair(item) in lmapping:
#                lmapping.append(self.NetworkDescriptionPair(item))
#        
#        return lmapping
#
#    def CallEmployeeInfo(self, loaded_file):
#        """
#        Function that returns the mapping between a network and its description
#        """
#        lmapping = []
#        
#        for item in loaded_file:
#            if not self.EmployeeInfo(item) in lmapping:
#                lmapping.append(self.EmployeeInfo(item))
#        
#        return lmapping
#
#    def CallReportData(self, loaded_file):
#        """
#        Function that returns the necessary info needed from the SAP Namerun
#        """
#        lmapping = []
#        
#        for item in loaded_file:
#            if not self.ReportData(item) in lmapping:
#                lmapping.append(self.ReportData(item))
#        
#        return lmapping