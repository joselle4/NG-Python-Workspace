'''
Created on Feb 21, 2012

@author: Joselle Abagat
'''

class Employees(object):
    '''
    classdocs
    '''
    
    def __init__(self):
        """
        """
        self.employee_list = []
        self.noEmployeeNumber_list = []
        self.noEmployeeFirstName_list = []
        self.noEmployeeLastName_list = []
        self.noEmployeeCamCode_list = []
        self.noEmployeeFHR_list = []        
        self.validateEmployeeEntries()
        self.employee_cam_dictionary = {} #self.employeeCamMapping()
        self.employee_fhr_dictionary = {} #self.employeeFhrMapping()
        self.employee_name_dictionary = {} #self.employeeNameMapping()
        self.employee_firstname_dictionary = {} #self.employeeFirstNameMapping()
        self.employee_lastname_dictionary = {} #self.employeeLastNameMapping()
        self.unknownEmployee_dictionary = {}
        self.unknownEmployeeFHR_dictionary = {}
        self.unknownEmployee_dictionaryInverted = {}
        self.unknownEmployee_list = []
        self.newEmployees_dictionary = {}
        self.editEmployee_list = []
        
#    def __init__(self, filepath):
#        """
#        """
#        self.employee_list = self.parseFile(filepath)
#        self.validateEmployeeEntries()
#        self.employee_cam_dictionary = self.employeeCamMapping()
#        self.employee_fhr_dictionary = self.employeeFhrMapping()
#        self.employee_name_dictionary = self.employeeNameMapping()
#        self.employee_firstname_dictionary = self.employeeFirstNameMapping()
#        self.employee_lastname_dictionary = self.employeeLastNameMapping()
        
    def parseFile(self, filepath):
        '''
        Function to load Employee CSV File
        '''
        
        filename = open(filepath,'r')

        for EmployeeLine in filename:
            if not EmployeeLine.startswith('//'):
                EmployeeItem = EmployeeLine.split(',')
                self.employee_list.append(Employee(EmployeeItem[0].strip(), EmployeeItem[1].strip(), EmployeeItem[2].strip(), EmployeeItem[3].strip(), EmployeeItem[4].strip()))
        
        filename.close()
        
    def validateEmployeeEntries(self):
        ''' Checks the raw data for missing or incomplete entries '''

        row = 0
        
        for employee in self.employee_list:
            
            if (employee.EmpFirstName == ""):
                if row not in self.noEmployeeFirstName_list:
                    self.noEmployeeFirstName_list.append(row)
            if (employee.EmpNo == ""):
                if employee.EmpFullName not in self.noEmployeeNumber_list:
                    self.noEmployeeNumber_list.append(employee.EmpFullName)
            if (employee.EmpLastName == ""):
                if row not in self.noEmployeeLastName_list:
                    self.noEmployeeLastName_list.append(row)
            if (employee.CAMCode == ""):
                if employee.EmpFullName not in self.noEmployeeCamCode_list:
                    self.noEmployeeCamCode_list.append(employee.EmpFullName)
            if (employee.FuncHomeRoom == ""):
                if employee.EmpFullName not in self.noEmployeeFHR_list:
                    self.noEmployeeFHR_list.append(employee.EmpFullName)
                
            row += 1
        
        if len(self.noEmployeeFirstName_list) > 0:
            print "Employees with no first name:"
            print self.noEmployeeFirstName_list
        if len(self.noEmployeeNumber_list) > 0:
            print "Employees with no Employee number:"
            print self.noEmployeeNumber_list
        if len(self.noEmployeeLastName_list) > 0:
            print "Employees with no last name:"
            print self.noEmployeeLastName_list
        if len(self.noEmployeeCamCode_list) > 0:
            print "Employees with no CAM Code:"
            print self.noEmployeeCamCode_list
        if len(self.noEmployeeFHR_list) > 0:
            #NEED CODE TO UPDATE FHR LIST WITH SAP DATA
            print "Employees with no FHR:"
            print self.noEmployeeFHR_list
    
    def checkListItem(self, item, listarg):
        """checks if an item is in a list"""
        return bool([lambda item: item in listarg])
                    
    def toHTML(self):
        """ returns a string formatted in html """
       
        s = "<table>\n<tr><th>Employee Name</th><th>Employee Number</th><th>CAM Code</th><th>FHR</th></tr>"
        
        for employee in self.employee_list:
            s += "\n\t" + employee.toHTML()
    
        return s + "\n</table>"
    
    def toHTMLFile(self, file_name):
        """ creates a file with html data """
        
        f = open(file_name, "w")
        
        f.write(self.toHTML())
    
    def employeeCamMapping(self):
        """ returns a dictionary of the form: { EmployeeNumber1 : CAMCode, ... , EmployeeNumbern : CAMCode} """
        
        for employee in self.employee_list:
            if employee.EmpNo not in self.employee_cam_dictionary: 
                try:
                    self.employee_cam_dictionary[employee.EmpNo] = employee.CAMCode
                except ValueError:
                    if employee.EmpFullName not in self.noEmployeeNumber_list:
                        self.noEmployeeNumber_list.append(employee.EmpFullName)

    # this will be built from the BW namerun
    def employeeFhrMapping(self):
        """ returns a dictionary of the form: { EmployeeNumber1 : FHR, ... , EmployeeNumbern : FHR} """

        for employee in self.employee_list:
            if employee.EmpNo not in self.employee_fhr_dictionary:
                try:
                    self.employee_fhr_dictionary[employee.EmpNo] = employee.FuncHomeRoom
                except ValueError:
                    if employee.EmpNo not in self.noEmployeeFHR_list:
                        self.noEmployeeFHR_list.append(employee.EmpNo)
        
    def employeeNameMapping(self):
        """ returns a dictionary of the form: { EmployeeNumber1 : FirstName LastName, ... , EmployeeNumbern : FirstName LastName} """

        for employee in self.employee_list:
            if employee.EmpNo not in self.employee_name_dictionary:
                try:
                    self.employee_name_dictionary[employee.EmpNo] = employee.EmpFullName
                except ValueError:
                    pass
                    #print("Employee " + employee.EmpFullName + " does not have an Employee Number.")
                    #captured in validateEmployeeEntries()
    
    def employeeFirstNameMapping(self):
        """ returns a dictionary of the form: { EmployeeNumber1 : FHR, ... , EmployeeNumbern : FHR} """

        for employee in self.employee_list:
            if employee.EmpNo not in self.employee_firstname_dictionary:
                try:
                    self.employee_firstname_dictionary[employee.EmpNo] = employee.EmpFirstName
                except ValueError:
                    if employee.EmpNo not in self.noEmployeeFirstName_list:
                        self.noEmployeeFirstName_list.append(employee.EmpNo)
        
    def employeeLastNameMapping(self):
        """ returns a dictionary of the form: { EmployeeNumber1 : FHR, ... , EmployeeNumbern : FHR} """
        
        for employee in self.employee_list:
            if employee.EmpNo not in self.employee_lastname_dictionary:
                try:
                    self.employee_lastname_dictionary[employee.EmpNo] = employee.EmpLastName
                except ValueError:
                    if employee.EmpNo not in self.noEmployeeLastName_list:
                        self.noEmployeeLastName_list.append(employee.EmpNo)
                    
    def getCAMfromEmployeeNumber(self, employee_number, employee_name):
        """ returns the employee's cam_code assignation and updates unknown employee dictionary {empNo: Name}"""
        try:
            return self.employee_cam_dictionary[employee_number]
        except KeyError:
            #print "ID: " + employee_number + " Not Found in CAMCode Lookup"
            if employee_number not in self.unknownEmployee_dictionary:
                self.unknownEmployee_dictionary[employee_number] = employee_name
            return "Not in CAMcode File"

    def getFHRfromEmployeeNumber(self, employee_number, employee_FHR):
        """ returns the employee's cam_code assignation and updates the unknown employee dictionary {empNo: FHR}"""        
        try:
            return self.employee_fhr_dictionary[employee_number]
        except KeyError:
            #print "ID: " + employee_number + " Not Found in CAMCode Lookup"
            if employee_number not in self.unknownEmployeeFHR_dictionary:
                self.unknownEmployeeFHR_dictionary[employee_number] = employee_FHR
            return "Not in FHR File"
        
    def getNamefromEmployeeNumber(self, employee_number):
        """ returns the employee's name based on his/her employee number """
        try:
            return self.employee_name_dictionary[employee_number]
        except KeyError:
            #print "ID: " + employee_number + " Not Found in CAMCode Lookup"
            return "Not in CAMcode File"
                
    def getFirstNamefromEmployeeNumber(self, employee_number):
        """ returns the employee's name based on his/her employee number """
        try:
            return self.employee_firstname_dictionary[employee_number]
        except KeyError:
            #print "ID: " + employee_number + " Not Found in CAMCode Lookup"
            return "Not in CAMcode File"
        
    def getLastNamefromEmployeeNumber(self, employee_number):
        """ returns the employee's name based on his/her employee number """
        try:
            return self.employee_lastname_dictionary[employee_number]
        except KeyError:
            #print "ID: " + employee_number + " Not Found"
            return "Not in CAMcode File"
    
    def getUnknownEmployeeID(self, employeeName):
        """ returns the employee's key based on the employee name."""
        try:
            return self.unknownEmployee_dictionaryInverted[employeeName]
        except KeyError:
            return '000000'
    
    def getUnknownEmployeeFHR(self, empID):
        """ returns the employee's FHR based on the employee's ID"""
        try:
            return self.unknownEmployeeFHR_dictionary[empID]
        except KeyError:
            return ""
    
    def UnknownEmployeeList(self):
        """ creates a list of the full names of the unknown employees"""
        for eachKey in self.unknownEmployee_dictionary.keys():
            if self.unknownEmployee_dictionary[eachKey] not in self.unknownEmployee_list:
                self.unknownEmployee_list.append(self.unknownEmployee_dictionary[eachKey])
        self.unknownEmployee_list.sort()
    
    def invertUnknownEmployeeDictionary(self):
        """inverts key to value relationship of the unknownEmployeeDictionary"""
        for eachKey in self.unknownEmployee_dictionary:
            if self.unknownEmployee_dictionary[eachKey] not in self.unknownEmployee_dictionaryInverted:
                self.unknownEmployee_dictionaryInverted[self.unknownEmployee_dictionary[eachKey]] = eachKey
    
    def storeNewEmployee(self, Name, CAM):
        ''' this method will add the new employee to self.employee_list and will update the initially empty self.newEmployees_dictionary
        
        will be a dictionary of the form:
        {empID: [firstname, lastname, CAM, FHR]}'''
        empID = self.getUnknownEmployeeID(Name)
        firstName = Name.split()[0]
        #print Name.split()
        #print Name.split()[1:]
        lastName = ' '.join(Name.split()[1:]) #to account for lastName with Jr, multiple words in a last name, etc
        #print lastName
        FHR = self.getUnknownEmployeeFHR(empID)
        
        #CONSTRUCT EMPLOYEE AND APPEND TO LIST
        # we don't have to check it is in self.employee_list since it got generated because it's not in the list in the first place
        newEmp = Employee(firstName, lastName, empID, CAM, FHR)
        self.employee_list.append(newEmp)
        
        #ADD TO DICTIONARY
        if empID not in self.newEmployees_dictionary:
            self.newEmployees_dictionary[newEmp.EmpNo] = [newEmp.EmpFirstName, newEmp.EmpLastName, newEmp.CAMCode, newEmp.FuncHomeRoom]
        
#        print empID + " " + firstName + " " + lastName + " " + CAM + " " + FHR + "." 
    
    def editCurrentEmployee(self, empName, camCode):
        ''' edits the camcode of the employee that's currently on the list'''
        
        for each in self.employee_list:
            if each.EmpFullName == empName:
                each.CAMCode = camCode
                self.editEmployee_list.append(empName)
                return

    def deleteCurrentEmployee(self, argList):
        ''' edits the camcode of the employee that's currently on the list'''
        
        for name in argList:
            for each in self.employee_list:
                if each.EmpFullName == name:
                    self.employee_list.remove(each)
                    break

        
class Employee(object):
    '''
    classdocs
    '''

    def __init__(self, EmpFirstName, EmpLastName, EmpNo, CAMCode, FuncHomeRoom):
        '''
        Constructor
        '''
        
        self.EmpFirstName = EmpFirstName
        self.EmpLastName = EmpLastName
        self.EmpNo = EmpNo
        self.CAMCode = CAMCode
        self.EmpFullName = EmpFirstName + " " + EmpLastName
        self.FuncHomeRoom = FuncHomeRoom
        
    def __str__(self):
        
        s = "\nEmployee:" + \
            "\n\tFirst Name: " + str(self.EmpFirstName) + \
            "\n\tLast Name: " + str(self.EmpLastName) + \
            "\n\tFull Name: " + str(self.EmpFullName) + \
            "\n\tEmployee Number: " + str(self.EmpNo) + \
            "\n\tCAM Code: " + str(self.CAMCode) + \
            "\n\tFunctional Home Room: " + str(self.FuncHomeRoom)
        
        return s
    
    def __eq__(self, test_case):
        """
        Overloads the equals operators.  Ensures that all instance variables
        are equal.
        """
        
        if self.CAMCode != test_case.CAMCode:
            return False
        if self.EmpFirstName != test_case.EmpFirstName:
            return False
        if self.EmpFullName != test_case.EmpFullName:
            return False
        if self.EmpLastName != test_case.EmpLastName:
            return False
        if self.EmpNo != test_case.EmpNo:
            return False
        if self.FuncHomeRoom != test_case.FuncHomeRoom:
            return False
        
        return True

    def toHTML(self):
        ''' list of employee in that CAM'''
        
        
        html = "<tr><td>" + self.EmpFullName + " </td>" + \
                "<td>" + self.EmpNo + "</td>" + \
                "<td>" + self.CAMCode + "</td>" + \
                "<td>" + self.FuncHomeRoom + "</td>" + \
                "</tr>"
                
        return html