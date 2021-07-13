'''
Created on Mar 9, 2012

@author: Joselle Abagat
'''
import unittest
from clsNetwork import *
from mdLoadFiles import *
import os
from clsMonthlyCalendar import *
from clsWeeklyCalendar import *
from clsCAM import *
from clsEmployee import *
from clsETC import *
from clsBWNamerun import *
from mdLoadFiles import *

global weeklyentry
global weeklycal
global weekly_cal_list
global monthlycal


class TestMappings(unittest.TestCase):

    def setUp(self):
        self.contract = 'BOA 16'    
        self.cam = ['HXX', 'Royal Gardner']
        self.empcam = ['HJX', '11917', 'Brad Bergman']
        self.empfhr = ['L065LF', '11917', 'Brad Bergman']
        self.monthcal = ['200101', '192']
        self.networkcontract = ['EHSS', 'LEH2120B2']
        self.networkcam = ['H0X', 'LEH2120B2']
        self.networkactivity = ['LEH2120B2', '0030', 'LEH2120B20030']
        self.networkstatus = ['LEH2120B2', 'Open']
        self.weeklycal = ['201001', '1/1/2010', '0']
        self.sapnetworkactivity = ['LT44561NZ', 'Q39G', 'LT44561NZQ39G']
        self.sapnetworkdescription = ['LT44561NZ', 'Q39G', 'Sustaining Engineering and Technical Ser']
        self.sapemp = ['224715', 'Danny Raglin']
        self.sapreport = ['LT44561NZ', 'Q39G', 'Sustaining Engineering and Technical Ser', '224715', 'Danny Raglin', 'P-GLS4', '03/02/2012', '1.800']
        self.etcnetworkcontract = ['EMDLBRF', 'LX806HBD1']
        self.etcnetworkcam = ['HBX', 'LX806HBD1']
        self.etcnetworketc = ['LX806HBD1', '201101', '10']
        self.etcreport = ['EMDLBRF', 'HBX', 'LX806HBD1', '201101', '10']
        self.weeklyhoursdict = ['1/1/2010', 0]
        
        ############################
        # Location of test file
        ############################
        
        self.network_input_file = "..\\..\\..\\..\\resources\\Networks.csv"
        #print(os.path.abspath(self.network_input_file))
        self.weeklycal_input_file = "..\\..\\..\\..\\resources\\WeeklyAcctCal.csv"
        self.monthlycal_input_file = "..\\..\\..\\..\\resources\\MonthlyAcctCal.csv"
        self.cam_input_file = "..\\..\\..\\..\\resources\\CAMs.csv"
        self.contract_input_file = "..\\..\\..\\..\\resources\\Contracts.csv"
        self.employee_input_file = "..\\..\\..\\..\\resources\\Employees.csv"
        #self.namerun_input_file = "..\\..\\..\\..\\resources\\Namerun.csv"
        self.namerun_input_file = "..\\..\\..\\..\\resources\\Namerun2.csv"
        #self.namerun_input_file = "..\\..\\..\\..\\resources\\Namerun3.csv"
        self.etc_input_file = "..\\..\\..\\..\\resources\\ETCs.csv"

    def tearDown(self):
        self.contract = None
        self.cam = None
        self.empcam = None
        self.empfhr = None
        self.monthcal = None
        self.networkactivity = None
        self.networkcam = None
        self.networkcontract = None
        self.networkstatus = None
        self.weeklycal = None
        self.sapemp = None
        self.sapnetworkactivity = None
        self.sapnetworkdescription = None
        self.sapreport = None
        self.etcnetworkcontract = None
        self.etcnetworkcam = None
        self.etcnetworketc = None
        self.etcreport = None

    def testMapContract(self):
        contract_list = ContractEntry.CallContracts(ContractEntry, loadContract(self.contract_input_file))
        self.assertEquals(self.contract, contract_list[0])
    
    def testMapCAM(self):
        cam_list = CAMEntry.CallCAMOwner(CAMEntry, loadCAM(self.cam_input_file))
        self.assertEquals(self.cam, cam_list[0])
    
    def testMapEmpCAM(self):
        emp_cam_list = EmployeeEntry.CallEmployeeCAM(EmployeeEntry, loadEmployee(self.employee_input_file))
        self.assertEquals(self.empcam, emp_cam_list[0])
    
    def testMapEmpFHR(self):
        emp_fhr_list =  EmployeeEntry.CallEmployeeFHR(EmployeeEntry, loadEmployee(self.employee_input_file))
        self.assertEquals(self.empfhr, emp_fhr_list[0])
    
    def testMapMonthCal(self):
        monthly_cal_list = MonthlyCalendarEntry.CallMonthlyCalendar(MonthlyCalendarEntry, loadMonthlyCalendar(self.monthlycal_input_file))
        self.assertEquals(self.monthcal, monthly_cal_list[0])
    
    def testMapNetworkContract(self):
        network_contract_list = NetworkEntry.CallNetworkContract(NetworkEntry, loadNetworks(self.network_input_file))
        self.assertEquals(self.networkcontract, network_contract_list[0])
        
    def testMapNetworkCAM(self):
        network_cam_list = NetworkEntry.CallNetworkCAM(NetworkEntry, loadNetworks(self.network_input_file))
        self.assertEquals(self.networkcam, network_cam_list[0])
    
    def testMapNetworkActivity(self):
        network_activity_list = NetworkEntry.CallNetworkActivity(NetworkEntry, loadNetworks(self.network_input_file))
        self.assertEquals(self.networkactivity, network_activity_list[0])
        
    def testMapNetworkStatus(self):
        network_status_list = NetworkEntry.CallNetworkStatus(NetworkEntry, loadNetworks(self.network_input_file))
        self.assertEquals(self.networkstatus, network_status_list[0])
    
    def testMapWeeklyCal(self):
        weekly_cal_list = WeeklyCalendarEntry.CallWeeklyCalendar(WeeklyCalendarEntry, loadWeeklyCalendar(self.weeklycal_input_file))
        self.assertEquals(self.weeklycal, weekly_cal_list[0])
    
    def testSAPNetworkActivity(self):
        sap_network_activity_list = NamerunEntry.CallNetworkActivity(NamerunEntry, loadNamerun(self.namerun_input_file))
        self.assertEquals(self.sapnetworkactivity, sap_network_activity_list[0])
    
    def testSAPNetworkDescription(self):
        sap_network_description_list = NamerunEntry.CallNetworkDescription(NamerunEntry, loadNamerun(self.namerun_input_file))
        self.assertEquals(self.sapnetworkdescription, sap_network_description_list[0])
    
    def testSAPEmp(self):
        sap_emp_info_list = NamerunEntry.CallEmployeeInfo(NamerunEntry, loadNamerun(self.namerun_input_file))
        self.assertEquals(self.sapemp, sap_emp_info_list[0])

    def testSAPreportdata(self):
        sap_reportdata_list = NamerunEntry.CallReportData(NamerunEntry, loadNamerun(self.namerun_input_file))
        self.assertEquals(self.sapreport, sap_reportdata_list[0])
        
    def testETCNetworkContract(self):
        etc_network_contract_list = ETCEntry.CallNetworkContract(ETCEntry, loadETC(self.etc_input_file))
        self.assertEquals(self.etcnetworkcontract, etc_network_contract_list[0])
    
    def testETCNetworkCAM(self):
        etc_network_cam_list = ETCEntry.CallNetworkCAM(ETCEntry, loadETC(self.etc_input_file))
        self.assertEquals(self.etcnetworkcam, etc_network_cam_list[0])
    
    def testETCNetworkETC(self):
        etc_network_etc_list = ETCEntry.CallNetworkToETC(ETCEntry, loadETC(self.etc_input_file))
        self.assertEquals(self.etcnetworketc, etc_network_etc_list[0])
    
    def testETCreportdata(self):
        etc_reportdata_list = ETCEntry.CallReportData(ETCEntry, loadETC(self.etc_input_file))
        self.assertEquals(self.etcreport, etc_reportdata_list[0])
        
    def testWeeklyHrsDict(self):
        weeklyentry = WeeklyCalendarEntry
        weeklycal = WeeklyCalendar
        
        dict = weeklycal.WeeklyHoursDict(weeklycal, weeklycal.parseFile(weeklycal, self.weeklycal_input_file))
        self.assertEquals(self.weeklyhoursdict, dict[0])
        
if __name__ == "__main__":
    #import sys;sys.argv = ['', 'Test.testName']
    unittest.main()
