'''
Created on Feb 23, 2012

@author: Joselle Abagat
'''

import datetime

class WeeklyCalendar(object):
    '''
    classdocs
    '''
        
    def __init__(self):
        self.weeklycal_list = [] #self.parseFile(filepath)
        self.hours_dict = {} #self.mapWeeklyHours()
        self.weekenddates_dict = {} #self.mapPeriodToWeekEndDates()
        self.weeksInPeriod_dict = {} #self.mapNumberOfWeeksInPeriod()()
        self.PeriodToWeek_dict = {} #self.mapWeekEndDateToPeriod()
        self.noPeriod_list = []
        self.noDate_list = []
        self.noWeekEndDate_list = []
    
    def parseFile(self, filepath):
        '''
        function to load the monthly accounting calendar; need csv
        '''
    
        filename = open(filepath,'r')
        
        for line in filename:
            if not line.startswith('//'):
                linevalues = line.split(',')
                self.weeklycal_list.append(WeeklyCalendarEntry(linevalues[0].strip(),linevalues[1].strip(),linevalues[2].strip(),linevalues[3].strip()))

                # need to account for date combinations: M/D/YYYY, MM/DD/YYYY, etc
                # loop through each of the values to search for the weekend dates.  weekend dates when split will be a list [MM, DD, YYYY] 
                for item in linevalues:
                    datesplit = item.split('/')
                    if len(datesplit) == 3:
                        
                        month = datesplit[0]
                        if len(month) == 1:
                            month = "0" + str(month)

                        day = datesplit[1]
                        if len(day) == 1:
                            day = "0" + str(day)
                            
                        year = datesplit[2]                        
                        if len(datesplit[0]) == 1 or len(datesplit[1]) == 1:
                            period = "/".join([month, day, year])
                            self.weeklycal_list.append(WeeklyCalendarEntry(linevalues[0].strip(),linevalues[1].strip(),period.strip(),linevalues[3].strip()))
        
        filename.close()
        
    def mapWeeklyHours(self):
        """
        returns a dictionary of the float(hours) using weekending date as the key {'date':hrs,}
        """
        
        for item in self.weeklycal_list:
            dict_key = str(item.WeekEndDate)
            dict_value = float(item.Hours)
            if not dict_key in self.hours_dict: #dict: #dict.__contains__(dict_key):
                self.hours_dict[dict_key] = dict_value
        
    def getHours(self, date):
        """
        return the hours associated to the weekending date
        """
        try:
            return self.hours_dict[date]
        except KeyError:
            if date not in self.noDate_list:
                self.noDate_list.append(date)
            print str(date) + " not found.  Pleaes update Accounting Calendar"
            return 0.0
    
    def mapPeriodToWeekEndDates(self):
        """
        returns a dictionary of the form { YYYYMM : [week1, week2,...], ..., YYYYMMn : [week1, week2,...]}
        """
        
        for entry in self.weeklycal_list:
            try:
                if int(entry.Period) not in self.weekenddates_dict:
                    self.weekenddates_dict[int(entry.Period)] = [entry.WeekEndDate]
                else:
                    self.weekenddates_dict[int(entry.Period)].append(entry.WeekEndDate)
            except ValueError:
                print entry.Period
                print entry.WeekEndDate
            except KeyError:
                print "Length of keys: " + len(self.weekenddates_dict.keys())
        
    def getWeekEndDatesFromPeriod(self, period): 
        #print self.weekenddates_dict
        return self.weekenddates_dict[int(period)]
        
    def mapNumberOfWeeksInPeriod(self):
        """
        returns a dictionary of the form { YYYYMM: #weeks, ... , YYYYMMn: #weeks }
        """
        
        #empty out the dictionary in case it is reloaded
        
        self.weeksInPeriod_dict = {}
        
        # since we are reusing the keys in the weekenddates_dictionary, we already know each is unique
        for each_key in self.weekenddates_dict.keys():
            self.weeksInPeriod_dict[each_key] = len(self.weekenddates_dict[each_key])
    
    def getWeeksInAPeriod(self, period):
        """
        returns the number of weeks within a period
        """
        
        try:
            return self.weeksInPeriod_dict[period]
        except KeyError:
            if period not in self.noPeriod_list:
                self.noPeriod_list.append(period)
            print (period + "Does not exist.  Update Weekly Accounting Calendar")
            return 0
            
    def mapWeekEndDateToPeriod(self):
        """
        returns dictionary of the form {week1:Period, week2:Period, ..., weekn:Period}
        """
        
        for each_entry in self.weeklycal_list:
            
            if not each_entry.WeekEndDate in self.PeriodToWeek_dict:
                self.PeriodToWeek_dict[each_entry.WeekEndDate] = int(each_entry.Period)
            else:
                if not self.PeriodToWeek_dict[each_entry.WeekEndDate] == int(each_entry.Period):
                    self.PeriodToWeek_dict[each_entry.WeekEndDate] = int(each_entry.Period)
        
#        print self.PeriodToWeek_dict
        
    def getPeriodFromWeekEndDate(self, WeekEndDate):
        """
        returns the Period given a weekend date
        """
        
        try:
            return self.PeriodToWeek_dict[WeekEndDate]
        except KeyError:
            if WeekEndDate not in self.noWeekEndDate_list:
                self.noWeekEndDate_list.append(WeekEndDate)
            #print(str(WeekEndDate) + " does not exist in the raw data.  Check format or Update the raw data")
            return "000000"
    
    def getHoursToDate(self):
        curWeekEndDate = self.convertDate(self.getCurrentWeekEndingDate())
        #print curWeekEndDate
        curPeriod = self.getPeriod(curWeekEndDate)
        #print curPeriod
        weekEndsList = self.getWeekEndDatesFromPeriod(curPeriod)

        newList = []
        for each in weekEndsList:
            convertedDate = self.convertDate(each)
            if convertedDate not in newList:
                newList.append(convertedDate)
            else:
                weekEndsList.remove(each) 
        
        sumHours = 0
        for each in weekEndsList:
            if self.convertDate(each) < curWeekEndDate:
                sumHours = sumHours + self.getHours(each)
        
        #print sumHours
        return float(sumHours)
    
    def removeDuplicates(self, listarg):
        return list(set(listarg))
    
    def getCurrentWeekEndingDate(self):
        ''' NEED A CHECK FOR NON-EXISTING ITEMS IN ACCOUNTING CAL 
            Obtains the correct weekend date for date now()'''
        
        '''obtain the current period'''
        curPeriod = self.getCurrentPeriod()
        
        '''to account for the accounting calendars, obtain the weekend dates for curPeriod +/- 1'''
        lastPeriod = self.getPeriod(self.dateNow() + datetime.timedelta(-30))#.strftime("%Y/%m/%d")
        nextPeriod = self.getPeriod(self.dateNow() + datetime.timedelta(+30))#.strftime("%Y/%m/%d")
        
        '''combine lists'''
        accountingWeeks = self.getWeekEndDatesFromPeriod(str(curPeriod)) + self.getWeekEndDatesFromPeriod(lastPeriod) + self.getWeekEndDatesFromPeriod(nextPeriod)
        
        newList = []
        for each in accountingWeeks:
            newList.append(self.convertDate(each))
        
        '''obtain the correct accounting weekending date dateNow belongs to'''
        newList.sort()
        for each in newList:
            if each >= self.dateNow():
                return each
    
    def convertDate(self, weekEndDate):
        ''' converts accounting calendar dates of format MM/D/YYYY or MM/DD/YYYY to python format YYYY-MM-DD'''
        
        #print "date: " + str(weekEndDate)
        try:
            dateConvert = datetime.datetime.strptime(str(weekEndDate), "%Y-%m-%d")
            dateConvert = dateConvert.strftime("%m/%d/%Y")
            dateConvert = datetime.datetime.strptime(dateConvert, "%m/%d/%Y").date()
        except ValueError:
            try:
                dateConvert = datetime.datetime.strptime(str(weekEndDate), "%m/%d/%Y").date()
            except ValueError:
                weekEndDate = weekEndDate.date()
                dateConvert = datetime.datetime.strptime(str(weekEndDate), "%Y-%m-%d")
                dateConvert = dateConvert.strftime("%m/%d/%Y")
                dateConvert = datetime.datetime.strptime(dateConvert, "%m/%d/%Y").date()

        return dateConvert
        
    def dateString(self, dateArg): return dateArg.strftime("%m/%d/%Y")
    
    def dateNow(self): return datetime.datetime.now().date()
    
    def getCurrentPeriod(self): return int(datetime.datetime.now().strftime("%Y%m"))
    
    def getPeriod(self, weekEndDate): return int(self.convertDate(weekEndDate).strftime("%Y%m"))
        
class WeeklyCalendarEntry(object):
    '''
    classdocs
    '''

    def __init__(self, ID, Period, WeekEndDate, Hours):
        '''
        Constructor
        '''
        
        self.ID = ID
        self.Period = Period
        self.WeekEndDate = WeekEndDate
        self.Hours = Hours
    
    def __str__(self):
        
        s = "\nWeekly Accounting Calendar" + \
            "\n\tID: " + str(self.ID) + \
            "\n\tPeriod: " + str(self.Period) + \
            "\n\tWeekEndDate: " + str(self.WeekEndDate) + \
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
        if self.WeekEndDate != test_case.WeekEndDate:
            return False
        
        return True
    
#    def ListWeeklyCalendar(self):
#        '''
#        returns a list of the Weekly Accounting Calendar in the form [Period, Week-Ending Date, Hours]
#        '''
#        lWeeklyCal = [self.Period, self.WeekEndDate, self.Hours]
#        
#        return lWeeklyCal
#    
#    def CallWeeklyCalendar(self, loaded_file):
#        '''
#        Returns a list of the Monthly Accounting Calendar in the form of [YYYYMM,Hours]
#        '''
#        
#        listCal = []
#        
#        for each_item in loaded_file:
#            if not self.ListWeeklyCalendar(each_item) in listCal:
#                listCal.append(self.ListWeeklyCalendar(each_item))
#        
#        return listCal