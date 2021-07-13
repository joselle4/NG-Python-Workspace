'''
Created on Feb 23, 2012

@author: Joselle Abagat
'''

class MonthlyCalendar(object):
    '''
    classdocs
    '''
    
    def __init__(self):
        self.monthlycal_list = [] #self.parseFile(filepath)
        self.periodHours_dictionary = {} #self.periodHoursMapping()
        self.noPeriodHours_list = []
        
    def parseFile(self, filepath):
        """
        """
        
        filename = open(filepath,'r')
        
        for line in filename:
            if not line.startswith('//'):
                linevalues = line.split(',')
                self.monthlycal_list.append(MonthlyCalendarEntry(linevalues[0].strip(),linevalues[1].strip(),linevalues[2].strip()))
    
    def periodHoursMapping(self):
        """
        returns a dictionary of the form { YYYYMM_1: hours}
        """
        
        for cal in self.monthlycal_list:
            try:
                if not int(cal.Period) in self.periodHours_dictionary:
                    try:
                        self.periodHours_dictionary[int(cal.Period)] = float(cal.Hours)
                    except ValueError:
                        print cal.Hours
            except ValueError:
                print cal.Period
    
    def getPeriodHours(self, period):
        """ returns the amount of hours for a given period """
        try:
            return self.periodHours_dictionary[period]
        except KeyError:
            if period not in self.noPeriodHours_list:
                self.noPeriodHours_list.append(period)
            return "period not in calendar"

class MonthlyCalendarEntry(object):
    '''
    classdocs
    '''

    def __init__(self, ID, Period, Hours):
        '''
        Constructor
        '''
        self.ID = ID
        self.Period = Period
        self.Hours = Hours
    
    def __str__(self):
        
        s = "\nMonthly Accounting Calendar:" + \
            "\n\tID: " + str(self.ID) + \
            "\n\tPeriod: " + str(self.Period) + \
            "\n\tHours: " + str(self.Hours)
        
        return s

    def __eq__(self,test_case):
        """
        Overloads the equals operators.  Ensures that all instance variables
        are equal.
        """
        if self.Hours != test_case.Hours:
            return False
        if self.ID != test_case.ID:
            return False
        if self.Period != test_case.Period:
            return False
        
        return True

#    def ListMonthlyCalendar(self):
#        '''
#        returns a list of the Monthly Accounting Calendar in the form [Period, Hours]
#        '''
#        
#        lMonthlyCal = [self.Period, self.Hours]
#        
#        return lMonthlyCal
#    
#    def CallMonthlyCalendar(self, loaded_file):
#        '''
#        Returns a list of the Monthly Accounting Calendar in the form of [YYYYMM,Hours]
#        '''
#        
#        listCal = []
#        
#        for each_item in loaded_file:
#            if not self.ListMonthlyCalendar(each_item) in listCal:
#                listCal.append(self.ListMonthlyCalendar(each_item))
#        
#        return listCal
