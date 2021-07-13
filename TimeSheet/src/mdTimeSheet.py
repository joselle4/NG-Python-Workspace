'''
Created on Mar 8, 2012

@author: Joselle Abagat
'''

from clsBWNamerun import *
from clsCAM import *
from clsContract import *
from clsEmployee import *
from clsETC import *
from clsMonthlyCalendar import *
from clsNetwork import *
from clsWeeklyCalendar import *
from clsReloadResources import *
from mdLoadFiles import *
import os
import re

##################
# LOAD ALL DATA INTO CLASSES
##################

raw_cam = "..\\resources\\CAMS.csv"
raw_contracts = "..\\resources\\Contracts.csv"
raw_employees = "..\\resources\\Employees.csv"
raw_etc = "..\\resources\\ETCs.csv"
raw_FHR = "..\\resources\\FHR.csv"
raw_monthlycal = "..\\resources\\MonthlyAcctCal.csv"
raw_namerun = "..\\resources\\NamerunCopy.csv"
raw_networks = "..\\resources\\Networks.csv"
raw_weeklycal = "..\\resources\\WeeklyAcctCal.csv"

cam = CAMs(raw_cam)
contract = Contracts(raw_contracts)
emp = Employees(raw_employees)
etc = ETC(raw_etc)
monthlycal = MonthlyCalendar(raw_monthlycal)
namerun = Nameruns(raw_namerun)
network = Networks(raw_networks)
weeklycal = WeeklyCalendar(raw_weeklycal)

#raw_weeklycal = "..\\resources\\WeeklyAcctCal.csv"
#wc = WeeklyCalendar(raw_weeklycal)
#print(wc.getHours('1/8/2010'))
#raw_employees = "..\\resources\\Employees.csv"
#
#employees = Employees(raw_employees)
##print(employees.employee_list[3].toHTML())
##print(employees.employee_list[4].toHTML())
#print(employees.toHTMLFile("employees.html"))
##print(employees.getCAMfromEmployeeNumber(127221))
#
#
#raw_cam = "..\\resources\\CAMS.csv"
#cams = CAMs(raw_cam)
#print(cams.getCAMOwner('HFX'))
#
#
##emp_cam_map = dict()
#
#def Timesheet():
#    
#    
#    report = []    
#    
#    ##################
#    # LOAD ALL DATA
#    ##################
#    
#    raw_cam = "..\\resources\\CAMS.csv"
#    raw_contracts = "..\\resources\\Contracts.csv"
#    raw_employees = "..\\resources\\Employees.csv"
#    raw_etc = "..\\resources\\ETCs.csv"
#    raw_FHR = "..\\resources\\FHR.csv"
#    raw_monthlycal = "..\\resources\\MonthlyAcctCal.csv"
#    raw_namerun = "..\\resources\\NamerunCopy.csv"
#    raw_networks = "..\\resources\\Networks.csv"
#    raw_weeklycal = "..\\resources\\WeeklyAcctCal.csv"
#    
#    cam_load = loadCAM(raw_cam)
#    contracts_load = loadContract(raw_contracts)
#    employees_load = loadEmployee(raw_employees)
#    etc_load = loadETC(raw_etc)
#    #fhr_load = loadFHR(raw_FHR)
#    monthlycal_load = loadMonthlyCalendar(raw_monthlycal)
#    namerun_load = loadNamerun(raw_namerun)
#    networks_load = loadNetworks(raw_networks)
#    weeklycal_load = loadWeeklyCalendar(raw_weeklycal)
#    
#    #####################
#    # CALL LIST MAPPINGS
#    #####################
#    
##    contract_list = Contract.CallContracts(Contract, contracts_load)
##    cam_list = CAM.CallCAMOwner(CAM, cam_load)
#    employee = Employee()
#    
#    emp_cam_list = employee.CallEmployeeCAM(employees_load)
##    emp_fhr_list =  Employee.CallEmployeeFHR(Employee, employees_load)
##    monthly_cal_list = MonthlyCalendar.CallMonthlyCalendar(MonthlyCalendar, monthlycal_load)
##    network_contract_list = Network.CallNetworkContract(Network, networks_load)
##    network_cam_list = Network.CallNetworkCAM(Network, networks_load)
##    network_activity_list = Network.CallNetworkActivity(Network, networks_load)
##    network_status_list = Network.CallNetworkStatus(Network, networks_load)
##    weekly_cal_list = WeeklyCalendar.CallWeeklyCalendar(WeeklyCalendar, weeklycal_load)
##    
##    sap_network_activity_list = NamerunEntry.CallNetworkActivity(NamerunEntry, namerun_load)
##    sap_network_description_list = NamerunEntry.CallNetworkDescription(NamerunEntry, namerun_load)
##    sap_emp_info_list = NamerunEntry.CallEmployeeInfo(NamerunEntry, namerun_load)
##    sap_reportdata_list = NamerunEntry.CallReportData(NamerunEntry, namerun_load)
##    
##    etc_network_contract_list = ETC.CallNetworkContract(ETC, etc_load)
##    etc_network_cam_list = ETC.CallNetworkCAM(ETC, etc_load)
##    etc_network_etc_list = ETC.CallNetworkToETC(ETC, etc_load)
##    etc_reportdata_list = ETC.CallReportData(ETC, etc_load)
##    
#    wc = WeeklyCalendar(raw_weeklycal)
#    dict = wc.WeeklyHoursDict()
#    print(dict['1/1/2010'])
#
#    ######################
#    # 
#    ######################
#    #report = SumNamerun(sap_reportdata_list)
#    
#    return report
#
#def SumNamerun(sap_report):
#    """
#    Method to sum up the hours if the following are the same: Charge#, Act, Desc, EmpID, Payroll Code, WeekEnd Date
#    SAP Report looks like:
#                   [0][0]    [0][1]  [0][2]    3      4        5        6        7
#            [
#         0       [Charge#_1, Act_1, Desc_1, EmpID_1, Name_1, Payroll_1, Date_1, Hours_1]
#         1       [Charge#_2, Act_2, Desc_2, EmpID_2, Name_2, Payroll_2, Date_2, Hours_2]
#                .
#                .
#                .
#         n       [Charge#_n, Act_n, Desc_n, EmpID_n, Name_n, Payroll_n, Date_n, Hours_n]
#            ]
#    """
#        
#    sum_list = []
#    
#    # loop through the report in order to look at each row of data
#    for each_list in sap_report:
#        #convert all hours into floats
#        each_list[-1] = float(each_list[-1])
#        # take out all the elements of that row except for the last one since those are the ones we want to compare
#        same_elems = each_list[0:(len(each_list) - 1)]
#        # find those list elements in sum_list; if it is found in_sum_list != []
#        in_sum_list = [findlist for findlist in sum_list if findlist[0:len(same_elems)] == same_elems]
#        
#        # if those items are already in sum_list, then we just want to add the last to the last value
#        # if these items are not in sum_list, then add the whole row into the list
#        try:
#            find_index = sum_list.index(in_sum_list[0])
#            sum_list[find_index][-1] = sum_list[find_index][-1] + each_list[-1]
#        except IndexError:
#            sum_list.append(each_list)
#            
#    return sum_list
#
#
#def query(ID=None, EmployeeFirstName=None):
#    
#    if not ID == None:
#        pass
#    if not EmployeeFirstName == None:
#        pass
#
#
#query(ID=5)
#query(EmployeeFirstName="Jesus")
#
##output = Timesheet()
##print(output[0])
