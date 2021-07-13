'''
Created on Mar 13, 2012

@author: Jesus Medrano, Joselle Abagat
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
from clsContract import *
#from mdTimeSheet import *

class TestLoad(unittest.TestCase):


    def setUp(self):
        self.network = NetworkEntry("1", "EHSS", "H0X", "LEH2120B2", "0030", "AV Program MGmt", "L0810H", "", "Open")
        self.weeklycal = WeeklyCalendarEntry("1", "201001", "1/1/2010", "0")
        self.monthlycal = MonthlyCalendarEntry("1", "200101", "192")
        self.cams = CAMEntry("1", "HXX", "Royal", "Gardner")
        self.contract = ContractEntry("1", "BOA 16")
        self.employee = EmployeeEntry("1", "Brad", "Bergman", "11917", "HJX", "L065LF")
        self.etc = ETCEntry("EMDLBRF", "HBX", "121-06HBD1", "ASIP DESIGN ANALY & DEV TESTS-LBR", "SUST", "", "", "LX806HBD1", "201101", "", "", "", "10", "L")
        #self.namerun = NamerunEntry("K08669", "K08669", "","EURO HAWK - BAF EMI Extension (CR 20197)", "LEH2120B2", "MAMU", "EH20083", "5190", "#", "#", "89078", "Elaine Bailey", "P-GEHC", "1", "E2", "REG", "","40963", "3", "0")
        #self.namerun = NamerunEntry("D08127", "D08127", "752030", "LT44561NZ", "Sustaining Engineering and Technical Ser", "Q39G", "#", "#", "#", "#", "224715", "Danny Raglin", "P-GLS4", "1", "E", "REG", "02/27/2012", "03/02/2012", "1.800", "1.800")
        self.namerun = NamerunEntry("D08127", "D08127", "752030", "LT44561NZ", "Sustaining Engineering and Technical Ser", "Q39G","#", "#", "#", "#", "224715", "Danny Raglin", "P-GLS4", "1", "E", "REG", "02/27/2012", "03/02/2012", "1.800", "1.800")
        
        
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
        self.network = None
        self.weeklycal = None
        self.monthlycal = None
        self.cams = None
        self.contract = None
        self.employee = None
        self.etc = None
        self.namerun = None
        
    def testLoadingNetworks(self):
        networks = loadNetworks(self.network_input_file)      
        self.assertEquals(self.network, networks[0])
    
    def testLoadingWeeklyCal(self):
        weekly_cal = loadWeeklyCalendar(self.weeklycal_input_file)
        self.assertEquals(self.weeklycal, weekly_cal[0])
    
    def testLoadingMonthlyCal(self):
        monthly_cal = loadMonthlyCalendar(self.monthlycal_input_file)
        self.assertEquals(self.monthlycal, monthly_cal[0])
    
    def testLoadingCAMs(self):
        cams = loadCAM(self.cam_input_file)
        self.assertEqual(self.cams, cams[0])
    
    def testLoadingContracts(self):
        contracts = loadContract(self.contract_input_file)
        self.assertEquals(self.contract, contracts[0])
    
    def testLoadingEmployee(self):
        employees = loadEmployee(self.employee_input_file)
        self.assertEquals(self.employee, employees[0])
    
    def testLoadingETC(self):
        etcs = loadETC(self.etc_input_file)
        self.assertEquals(self.etc, etcs[0])
    
    def testLoadingNamerun(self):
        nameruns = loadNamerun(self.namerun_input_file)
        self.assertEquals(self.namerun, nameruns[0])

             
if __name__ == "__main__":
    #import sys;sys.argv = ['', 'Test.testName']
    unittest.main()