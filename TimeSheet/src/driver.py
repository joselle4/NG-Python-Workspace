'''
Created on May 14, 2012

@author: Joselle Abagat
'''

import easygui, wx, sys, os, re, win32com.client, csv, shutil, threading, time, datetime, webbrowser
from threading import Thread
#from wx import *
import wx.lib.scrolledpanel as scroll
from clsReloadResources import *
from clsBWNamerun import *
from clsCAM import *
from clsContract import *
from clsEmployee import *
from clsMonthlyCalendar import *
from clsNetwork import *
from clsWeeklyCalendar import *
from clsETC import *
from clsGUI import *
from clsExcel import *

'''
' DIRECTORY MAPPINGS GLOBAL VARIABLES
'''

tsaIcon = "clock.ico"
ghImage = "gh2.jpg"
gradient = "gradient.jpg"
mappingDirectory = "J:\Air Vehicle\IMF\Mappings"
mappingUserGuide = os.path.join(mappingDirectory, "User Guide")
mappingResources = os.path.join(mappingDirectory, "Resources")
resourcesDirectory = os.path.join(os.getenv("USERPROFILE"), "My Documents", "Timesheet")
mpmExports = os.path.join(resourcesDirectory, "MPMExports")
namerun = os.path.join(resourcesDirectory, "Namerun")
reports = os.path.join(resourcesDirectory, "Reports")
userguide = os.path.join(resourcesDirectory, "User Guide")
resources = os.path.join(resourcesDirectory, "Resources")
winmpm = os.path.abspath("C:\Program Files\WINMPM")
tsaWiki = "http://wiki.northgrum.com/wiki/HALE_TimeSheet"


mCalFilename = "MonthlyAccountingCalendar.csv"
wCalFilename = "WeeklyAccountingCalendar.csv"
empFilename = "Employees.csv"
mpmFormat = "TSA_EOC_EXPORT.fmt"

mainInstructions = os.path.join(userguide, "MainInstructions.txt") 
mpmRequirements = os.path.join(userguide, "MPMRequirements.txt")
bwRequirements = os.path.join(userguide, "BWRequirements.txt")

'''
' DECLARE ALL CLASSES AS GLOBAL VARIABLES
'''

src = ReloadResources(mappingDirectory, resourcesDirectory, mpmExports, namerun, reports, userguide, resources, mappingUserGuide, mappingResources)
mcal = MonthlyCalendar()
wcal = WeeklyCalendar()
emp = Employees()
mpm = ETC()
bw = Nameruns(mcal, wcal, emp, mpm)
xl = Excel()
#bw = Nameruns(None, None, None, None)

'''
' CREATE GUIs
'''

class DriverApp(wx.App):
    """ need wx.app in order to create frames and subframes"""
    
    def __init__(self, redirect = False, filename = None, useBestVisual = True, clearSigInt = True):
        """ overwrite wx.app instance """
        wx.App.__init__(self, redirect, filename, useBestVisual, clearSigInt)
        
    def OnInit(self):
        """ create on initialize method to run the data loading main frame when class instance is called"""
        frame = MainFrame()
        frame.dataLoadMainFrame()
        src.copyResources(mappingUserGuide, userguide)
        return True

class Threading(Thread):
    
    def __init__(self):
        '''initialize worker thread class'''
        Thread.__init__(self)
        self.start()    #starts thread
    
    def runSubframe(self, parent, *listarg):
        ''' this code will execute in the new thread '''
        newEmployeeWindow = SubFrame(parent = parent, title = "")
        newEmployeeWindow.newEmployeeMainFrame(listarg)
    
class MainFrame(wx.Frame):
    """ data loading main frame"""
    
    def __init__(self, parent = None, title = ""):
        """ overwrite wx.frame __init__ with parent id = 1"""
        wx.Frame.__init__(self, parent, 1, title) 
    
    def boldFont(self): return wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD)
    
    def resizeWH(self, button, width, height): return button.SetSizeWH(width, height)
    
    def textColor(self, button, color = "#000000"): return button.SetForegroundColour(color)

    def backgroundColor(self, button, color = "#000000"): return button.SetBackgroundColour(color)
        
    def dataLoadMainFrame(self):
        """
        creates buttons for the mainframe:
                Load Resources
                Load Load MPM/ETC Files
                Generate Networks
                Load BW/SAP Namerun
                     reset
                RUN        CLOSE
                
        generate networks and load BW/SAP will not turn on until ETC files have been loaded
        RUN will not turn on until all files have been loaded
        """

        self.SetSize((745,470))
        self.SetMinSize((745,470)) #LOCKS SIZE
        self.SetMaxSize((745,470)) #LOCKS SIZE 
        self.SetTitle("Timesheet Report")      
        self.SetIcon(wx.Icon(tsaIcon, wx.BITMAP_TYPE_ICO))
        
        id = wx.ID_ANY

        self.addMenu()
        
        '''set up panel and buttons for loading'''        
        self.commandPanel = wx.Panel(self, id)
        self.commandPanel.SetSize((270,350))
        #commandPanel.SetBackgroundColour("#008080")
        #self.commandPanel.SetBackgroundColour("#003366")
        self.commandPanel.SetBackgroundColour("#191970")
        self.commandButtons = wx.FlexGridSizer(rows = 4, vgap = 30)
        
        self.loadResourcesButton =  wx.Button(self.commandPanel, id, "1. Load Resources", (30, 40))
        self.Bind(wx.EVT_BUTTON, self.onLoadResources, self.loadResourcesButton, id)

        self.loadMPMFilesButton = wx.Button(self.commandPanel, id, "2. Load MPM/ETC Files", (30, 100))
        self.Bind(wx.EVT_BUTTON, self.onLoadMPM, self.loadMPMFilesButton, id)
        
        self.generateNetworksButton = wx.Button(self.commandPanel, id, "Generate Networks (optional)", (30, 160))
        ''''DISABLE GENERATE BUTTON'''
        self.generateNetworksButton.Enable(False)
        self.Bind(wx.EVT_BUTTON, self.onGenerate, self.generateNetworksButton, wx.ID_OK) #wx.ID_OK = optional 
        
        self.loadNamerunButton = wx.Button(self.commandPanel, id, "3. Load BW/SAP Namerun", (30, 220))
        self.Bind(wx.EVT_BUTTON, self.onLoadNamerun, self.loadNamerunButton, id)
        
        self.commandList = [self.loadResourcesButton, self.loadMPMFilesButton, self.generateNetworksButton, self.loadNamerunButton]
        for eachButton in self.commandList:
            self.resizeWH(eachButton, 200, 30)
            eachButton.SetBackgroundColour(wx.WHITE)
            eachButton.SetFont(self.boldFont())
            self.commandButtons.Add(eachButton)

        ''''add a reset button'''
        self.resetButton = wx.Button(self.commandPanel, id, "RESET", (100, 300))
        self.resetButton.SetFont(self.boldFont())
        self.resetButton.SetBackgroundColour(wx.WHITE)
        self.Bind(wx.EVT_BUTTON, self.onReset, self.resetButton, id)
        
        ''''set up panel and buttons for program run/exit'''
        self.systemPanel = wx.Panel(self, id, style = wx.SIMPLE_BORDER)
        self.systemPanel.SetSize((270,100))
        self.systemPanel.SetBackgroundColour(wx.WHITE)
        self.systemButtons = wx.FlexGridSizer()

        self.openButton = wx.Button(self.systemPanel, wx.ID_OPEN, "RUN", (40, 25)) #ID_OPEN = will open new form
        ''''DISABLE RUN BUTTON'''
        self.openButton.Enable(False) 
        self.backgroundColor(self.openButton, "#003300")
        self.Bind(wx.EVT_BUTTON, self.onOpen, self.openButton, wx.ID_OPEN)
        
        self.closeButton = wx.Button(self.systemPanel, wx.ID_EXIT, "EXIT", (150, 25)) #ID_EXIT = EXIT Program
        self.backgroundColor(self.closeButton, "#800000")
        self.Bind(wx.EVT_BUTTON, self.onExit, self.closeButton, wx.ID_EXIT)
        
        systemList = [self.openButton, self.closeButton]
        for eachButton in systemList:
            eachButton.SetFont(self.boldFont())
            self.textColor(eachButton, wx.WHITE)
            self.systemButtons.Add(eachButton)

        '''set up image'''
        img = wx.Bitmap(ghImage)
        self.imagePanel = wx.Panel(self, id)
        self.imagePanel.SetSize((470,470))
        self.imagePanel.Move((270,0))
        wx.StaticBitmap(self.imagePanel, -1, img)      
        
        '''add gradient'''
        self.commandPanel.Bind(wx.EVT_ERASE_BACKGROUND, self.OnEraseBackground)
                        
        ''''set up frame, panels'''
        self.CenterOnScreen()
        self.systemPanel.Move((0, 350))
               
        self.Update()
        self.Show()

    def OnEraseBackground(self, evt):
        """
        Add a picture to the background
        """
        # yanked from ColourDB.py
        dc = evt.GetDC()
     
        if not dc:
            dc = wx.ClientDC(self)
            rect = self.GetUpdateRegion().GetBox()
            dc.SetClippingRect(rect)
        dc.Clear()
        bmp = wx.Bitmap(gradient)
        dc.DrawBitmap(bmp, 0, 0)
    
    def addMenu(self):
        ''' SETS UP MENU '''
        
        '''SET UP MENUBAR'''
        self.menu = wx.MenuBar()
        
        ''''ADD MENU ITEMS HERE'''
        self.menuHelp = wx.Menu()
        self.menu.Append(self.menuHelp, '&Help...')
        self.menuEdit = wx.Menu()
        self.menu.Append(self.menuEdit, '&Edit...', )
        self.menuOpen = wx.Menu()
        self.menu.Append(self.menuOpen, '&Open...')
        self.menuAdd = wx.Menu()
        self.menu.Append(self.menuAdd, '&Add...')
        
        '''ADD SUB-MENU ITEMS HERE'''
        '''SUB MENU ITEM 1'''
        self.menuHelp.Append(101, '&Main Instructions')
        self.Bind(wx.EVT_MENU, self.onHelpMain, id=101)
        self.menuHelp.Append(102, '&MPM Requirements')
        self.Bind(wx.EVT_MENU, self.onHelpMPM, id=102)    
        self.menuHelp.Append(103, '&BW Requirements')
        self.Bind(wx.EVT_MENU, self.onHelpBW, id=103)
        self.menuHelp.Append(104, '&wiki')
        self.Bind(wx.EVT_MENU, self.onHelpWiki, id=104)
        '''SUB MENU ITEM 2'''
        self.menuEditEmployee = wx.Menu()
        self.menuEdit.AppendMenu(201, '&Employee(s)', self.menuEditEmployee)
        self.menuEditEmployee.Append(2011, '&Change CAM')
        self.Bind(wx.EVT_MENU, self.onEditEmployee, id=2011)
        self.menuEditEmployee.Append(2012, '&Delete')
        self.Bind(wx.EVT_MENU, self.onDeleteEmployee, id=2012)        
        self.menuEdit.Append(202, '&Accounting Calendar')
        self.Bind(wx.EVT_MENU, self.onEditAccountingCalendar, id=202)
        ''' EDIT SUB MENU '''
        
        '''SUB MENU ITEM 3'''
        self.menuOpen.Append(301, '&Timesheet')
        self.Bind(wx.EVT_MENU, self.onOpenTimesheet, id=301)
        self.menuOpen.Append(302, '&MPM Exports')
        self.Bind(wx.EVT_MENU, self.onOpenMPMExports, id=302)
        self.menuOpen.Append(303, '&Namerun')
        self.Bind(wx.EVT_MENU, self.onOpenNamerun, id=303)
        self.menuOpen.Append(304, '&Reports')
        self.Bind(wx.EVT_MENU, self.onOpenReports, id=304)
        '''SUB MENU ITEM 4'''
        self.menuAdd.Append(401, '&New Employee')
        self.Bind(wx.EVT_MENU, self.onAddEmployee, id=401)
        self.menuAdd.Append(402, '&New Calendar Items')
        self.Bind(wx.EVT_MENU, self.onAddAccountingCalendar, id=402)
        
        '''set up menu'''
        self.SetMenuBar(self.menu)
        '''DISABLE MENU ITEMS UNTIL DATA LOAD'''
        self.menuItemList = [201, 202, 401, 402]
        for each in self.menuItemList:
            self.menu.Enable(each, enable = False)
        
    def onLoadResources(self, event):
        """ 1. copies resources from the mapping directory to the resources directory
            2. loads monthly accounting calendar (max 2017)
            3. loads weekly calendar (max 2017)
            4. loads employee file """
        
        t0 = time.time()
        
        self.loadResourcesButton.Disable()
        
        '''COPY RESOURCES'''
        if src.copyResources(mappingResources, resources) != False:    
            
            self.loadMonthlyCalendar()
            self.loadWeeklyCalendar()
            self.loadEmployeeFile()
            
            if os.path.isdir(winmpm) == True:
                src.copyResource(os.path.join(resources, mpmFormat), winmpm)
            #else:
                #MsgDialog(None, msg = ("WINMPM not installed.  EOC format not loaded."), caption = "Message", style = wx.OK).msg()
            
            '''CHECK LOADS AND CONFIRM COMPLETE'''
            t1 = time.time()
            span = self.timeSpan(t0, t1)
            self.checkLoads()
            MsgDialog(None, msg = ("Load Complete\n" + span), caption = "Message", style = wx.OK).msg()
        else:
            self.loadResourcesButton.Enable(True)
    
    def loadMonthlyCalendar(self):
        """loads monthly calendar"""
        
        '''LOAD RESOURCES CLASSES: MONTHLY CALENDAR, WEEKLY CALENDAR, EMPLOYEE'''
        mCalFilepath = src.createFilePath(resources, mCalFilename)
        mcal.parseFile(mCalFilepath)
        
        '''LOAD MAPPINGS'''
        mcal.periodHoursMapping() #mcal.periodHours_dictionary
                
    def loadWeeklyCalendar(self):
        """loads weekly calendar"""
        
        '''LOAD RESOURCES CLASSES: MONTHLY CALENDAR, WEEKLY CALENDAR, EMPLOYEE'''
        wCalFilepath = src.createFilePath(resources, wCalFilename)            
        wcal.parseFile(wCalFilepath)
        
        wcal.mapWeeklyHours() #loads wcal.hours_dict
        wcal.mapPeriodToWeekEndDates() # loads wcal.weekenddates_dict
        wcal.mapNumberOfWeeksInPeriod() #loads wcal.weeksInPeriod_dict()
        wcal.mapWeekEndDateToPeriod() #loads wcal.PeriodToWeek_dict
                
    def loadEmployeeFile(self):
        """loads employee cam code file"""
        
        '''LOAD RESOURCES CLASSES: MONTHLY CALENDAR, WEEKLY CALENDAR, EMPLOYEE'''
        empFilepath = src.createFilePath(resources, empFilename)            
        emp.parseFile(empFilepath)
        
        emp.validateEmployeeEntries() #re-validate employee list
        emp.employeeCamMapping() #loads emp.employee_cam_dictionary
        emp.employeeNameMapping() #loads emp.employee_name_dictionary
        emp.employeeFirstNameMapping() #loads emp.employee_firstname_dictionary
        emp.employeeLastNameMapping() #loads emp.employee_lastname_dictionary
                    
    def onLoadMPM(self, event):
        """ converts all mpm files to csv and parses through each of the files and
            loads each file as part of the mpm class """
            
        t0 = time.time()
        
        self.loadMPMFilesButton.Disable()
        self.Hide()
        
        '''CONVERT and PARSE THROUGH ALL MPM FILES'''
        value = src.convertAllMPMFiles()
        if value == True:
            '''CHECK THAT ALL FILES WERE PROPERLY CONVERTED TO CSV'''
            if len(src.notConverted) == 0:
                '''SET MAX TO NUMBER OF STEPS'''
                max = 8
                count = 0
                
                dlg = wx.ProgressDialog("Progress", "Parsing MPM Files", maximum = max, parent = self, 
                                        style = 
                                        #wx.PD_CAN_ABORT 
                                         wx.PD_APP_MODAL
                                        | wx.PD_ELAPSED_TIME
                                        #| wx.PD_ESTIMATED_TIME
                                        | wx.PD_REMAINING_TIME
                                        | wx.PD_AUTO_HIDE)
                
                count += 1
                dlg.Update(count, 'Parsing MPM files')
                for each_file in src.getMPMfiles():
                    '''PARSE ONLY CONVERTED FILES'''                        
                    
                    if each_file in src.converted:
                        #print each_file
                        mpm.parseFile(each_file)

                '''LOAD ALL CLASS MAPPINGS/DICTIONARIES'''
                #print 'run1'
                count += 1
                dlg.Update(count, 'Mapping Descriptions')
                mpm.mapNetworkToDescription()
                
                #print 'run2'
                count += 1
                dlg.Update(count, 'Mapping Contracts')
                mpm.mapNetworkToContract() #loads mpm.dict_networkToContract
                
                #print 'run3'
                count += 1
                dlg.Update(count, 'Mapping PoPs')
                mpm.mapNetworkToPOP() #loads mpm.dict_networkToPOP
                
                #print 'run4'
                count += 1
                dlg.Update(count, 'Mapping Responsible')
                mpm.mapNetworkToCAM() #loads mpm.dict_networkToCAM

#                for key in mpm.dict_networkToCAM:
#                    print str(key) + " has " + str(len(mpm.dict_networkToCAM[key])) + " CAM(s)"
                
                #print 'run5'
                count += 1
                dlg.Update(count, 'Mapping Period')
                mpm.mapNetworkToCAMPeriod() #loads mpm.dict_networkCAMPeriod
                
                #print 'run6'
                count += 1
                dlg.Update(count, 'Mapping ETCs')
                mpm.mapNetworkToETC() #loads mpm.dict_networkCAMPeriodETC

                #print 'run7'
                count += 1
                dlg.Update(count, 'Mapping WBS')
                mpm.mapNetworkToWBS() #loads mpm.dict_networkWBS
                
                self.checkLoads()
                t1 = time.time()
                span = self.timeSpan(t0, t1)
                MsgDialog(None, msg = "Load Complete.\n" + span, caption = "Message", style = wx.OK).msg()
                dlg.Destroy()

            else:
                msg = "The following files was not converted to csv:\n\t" + \
                        str(src.notConverted) + \
                        "\n1. Check and ensure each file contains only one sheet." + \
                        "\n2. Remove non-MPM exported files from folder." + \
                        "\n3. Press load again."                        
                MsgDialog(self, msg, "Conversion Failed", wx.OK).msg()
                self.loadMPMFilesButton.Enable(True)
        else:
            MsgDialog(None, msg = "Please load again", caption = "Message", style = wx.OK).msg()
            self.loadMPMFilesButton.Enable(True)
        
        self.Show()
        
    def onGenerate(self, event):
        """GENERATE NETWORKS GIVEN RESP FIELD; PROVIDE SKIP OPTION IF THEY ALREADY HAVE IT"""
        
        title = "Generate Networks?"
        msg = "Do you want to generate a list of networks for a given CAM(s) or skip the process?\n\tYES = Generate Networks\n\tNO = Skip Step\n\tCancel = Exit"
        yesNoGenerate = MsgDialog(None, msg, title).yesNoSelect() 
        
        if  yesNoGenerate == 1:
            
            mpm.getCAMs()
            camList = mpm.camList
            
            cams = self.displayCAMChoice(self.insertSelectAll(camList))
            try:
                if "Select All" not in cams:    
                    mpm.getNetworks(cams)
                else:
                    mpm.getNetworks(mpm.camList)
    
                self.displayNetworks()
            except TypeError:
                pass
                #MsgDialog(self, "No CAM selected", "No Data", wx.OK).msg()

    def insertSelectAll(self, listarg):
        """inserts a Select All option to a list"""
        
        if "Select All" not in listarg:
            listarg.insert(0, "Select All")
        return listarg
        
    def onLoadNamerun(self, event):
        """loads BW namerun file(s) into namerun class"""
         
        t0 = time.time()
        
        self.loadNamerunButton.Enable(False)
        self.Hide()
         
        '''RE-INITIALIZE BW WITH FILLED CLASSES'''
        bw.__init__(mcal, wcal, emp, mpm)
         
        '''CHECK NAMERUN FOLDER AND OBTAIN FILEPATHS; IF EMPTY...'''
        if src.checkNamerunDirectory() == True:
            self.loadNamerunButton.Disable()
            '''GET THE LIST OF FILEPATHS AND LOAD ONTO NAMERUN CLASS'''
            namerunFilepaths = src.getNamerunFiles()
            try:
                for each_file in namerunFilepaths:
                    parser = bw.parseFile(each_file)
                
                if parser != False:
                    self.loadNamerunMappings()                    
                else:
                    MsgDialog(None, msg = "No data loaded.\nPlease check contents of loaded file and load again.", caption = "Message", style = wx.OK).msg()
                    self.loadNamerunButton.Enable(True)
            except TypeError:
                MsgDialog(None, msg = "No data loaded.\nPress Reset to load again and load again.", caption = "Message", style = wx.OK).msg()
                self.loadNamerunButton.Enable(True)
        else:
            MsgDialog(None, msg = "Please load again", caption = "Message", style = wx.OK).msg()
        
        self.Show()

    def loadNamerunMappings(self):
        """ method to load namerun mappings that will in turn load data into the directories """
        
        t1 = time.time()

        '''SET MAX TO NUMBER OF STEPS'''
        max = 17
        count = 0
        
        dlg = wx.ProgressDialog("Progress", "Parsing Namerun", maximum = max, parent = self, 
                                style = 
                                #wx.PD_CAN_ABORT 
                                 wx.PD_APP_MODAL
                                | wx.PD_ELAPSED_TIME
                                #| wx.PD_ESTIMATED_TIME
                                | wx.PD_REMAINING_TIME
                                | wx.PD_AUTO_HIDE)        
        
        
        '''LOAD ALL NAMERUN MAPPINGS AND DICTIONARIES'''
        
        '''CONVERT TO LIST'''
        count += 1
        dlg.Update(count, 'Converting Namerun')
        
        if bw.convertToListItems() == False:
            MsgDialog(self, msg = 'Error in converting BW. Please check CSV file and try again', caption = 'Conversion Error', style = wx.OK).msg()
            dlg.Destroy()
            self.loadNamerunButton.Enable(True)        
        else:
            '''MAPPINGS'''
            count += 1
            dlg.Update(count, 'Mapping Descriptions')
            bw.mapNetworkToDescription()
            
            count += 1
            dlg.Update(count, 'Mapping Activities')
            bw.mapNetworkToActivity()
            
            count += 1
            dlg.Update(count, 'Mapping Period')
            bw.mapNetworkToPeriod()
            
            count += 1
            dlg.Update(count, 'Mapping Hours by Period')
            bw.mapNetworkTotalHoursByPeriod()
            
            count += 1
            dlg.Update(count, 'Mapping Weekend Dates')
            bw.mapNetworkToWeekEndDate()        
            
            count += 1
            dlg.Update(count, 'Mapping Hours by Weekend Dates')
            bw.mapNetworkTotalHoursByWeekEndDate()
            
            count += 1
            dlg.Update(count, 'Mapping Hours Code')
            bw.mapNetworkActivityEmployeeHrsCodeWeekEndingToHours()
            
            count += 1
            dlg.Update(count, 'Mapping Activity')
            bw.mapEmployeeNetworkActivityHrsCodeWeekEndingToHours()
                    
            '''PARSE'''
            count += 1
            dlg.Update(count, 'Parsing Namerun')
            bw.parseNANHW()
            
            '''UPDATE'''
            count += 1
            dlg.Update(count, 'Inputting Charge Comments')
            bw.updateChargeComments()
            
            count += 1
            dlg.Update(count, 'Mapping to MPM Files')
            bw.updateETC()
            
            count += 1
            dlg.Update(count, 'Calculating Total Hours')
            bw.updateTotalHours()
            
            count += 1
            dlg.Update(count, 'Calculating % Spent')
            bw.updatePercentSpent()
            
            '''LISTS'''
            count += 1
            dlg.Update(count, 'Mapping Accounting Calendar')
            bw.listAccountingMonths()
            
            count += 1
            dlg.Update(count, 'Activity Code Check')
            bw.listActivityCodes()
            
            count += 1
            dlg.Update(count, 'Updating Charge Comments')
            bw.listChargeComments()
            
            #test by printing
    #        print bw.dict_NetworkToActivities
    #        print bw.dict_NetworkToPeriod
    #        print bw.dict_NetworkToWeekEndDate
    #        print bw.dict_NetworkTotalHoursByPeriod
    #        print bw.dict_NetworkTotalHoursByWeekEndDate
                      
            t2 = time.time()
            
            self.checkLoads()
            span = self.timeSpan(t1, t2)
            MsgDialog(None, msg = ("Load Complete.\n" + span), caption = "Message", style = wx.OK).msg()
            dlg.Destroy()       
                                
    def onReset(self, event):
        """ reset load buttons """
        
        '''re-initialize everything'''
        src.__init__(mappingDirectory, resourcesDirectory, mpmExports, namerun, reports, userguide, resources, mappingUserGuide, mappingResources)
        mcal.__init__()
        wcal.__init__()
        emp.__init__()
        mpm.__init__()
        bw.__init__(mcal, wcal, emp, mpm)
        
        '''enable disabled buttons'''
        if self.loadMPMFilesButton.Enabled == False:
            self.loadMPMFilesButton.Enable(True)
        if self.loadResourcesButton.Enabled == False:
            self.loadResourcesButton.Enable(True)
        if self.loadNamerunButton.Enabled == False:
            self.loadNamerunButton.Enable(True)
        
        '''disable enabled buttons'''
        for each in self.menuItemList:
            self.menu.Enable(each, enable = False)
        if self.openButton.Enabled == True:
            self.openButton.Enable(False)
        if self.generateNetworksButton.Enabled == True:
            self.generateNetworksButton.Enable(False)
            
    def onOpen(self, event):
        """ perform all calculations and open reporting subframe with MainFrame as parent"""
        
        #PROGRAM SUMMARY
        
        SubFrame(parent = self, title = "").reportMainFrame()
        self.Hide()
        
    def onExit(self, event):
        """close frame and exit"""
        #xl.quitExcel()
        self.Destroy()
        sys.exit(0)
    
    def onDeleteEmployee(self, event):
        ''' delete employee '''
        emp.editEmployee_list = []
        empList = emp.employee_name_dictionary.values()
        empList.sort()
        
        if len(emp.employee_list) <> 0:
            deleteList = MultiChoiceDialog(self, "Choose Employee(s) to Delete", "Delete Employee(s)", empList).select(empList)
            
            if deleteList != None:
                emp.deleteCurrentEmployee(deleteList)
                src.reloadEmployee(empFilename, emp.employee_list)
            else:
                pass
                #MsgDialog(self, "Must Load All Data before Editing", "No Data", style = wx.OK).msg()
    
    def onEditEmployee(self, event):
        ''' displays a list of employees to edit '''
        
        emp.editEmployee_list = []
        
        empList = emp.employee_name_dictionary.values()
        empList.sort()
        
        if len(emp.employee_list) <> 0:
            editList = MultiChoiceDialog(self, "Choose Employee(s) to Edit", "Edit Employee(s)", empList).select(empList)
            
            print editList
            if editList != None: 
                mpm.getCAMs()
                editEmpWindow = SubFrame(parent = self, title = "")
                editEmpWindow.editEmployeeMainFrame(editList, mpm.camList)
                self.Hide()
        else:
            #pass
            MsgDialog(self, "Must Load All Data before Editing", "No Data", style = wx.OK).msg()
            
    def onEditAccountingCalendar(self, event):
        print 'edit calendar'
    
    def onAddEmployee(self, event):
        ''' adds employee if missed in loading of namerun '''

        if len(emp.unknownEmployee_dictionary) > 0:                
            emp.invertUnknownEmployeeDictionary()
            emp.UnknownEmployeeList()
        
            emp.newEmployees_dictionary = {}
            
            employeeList = emp.unknownEmployee_list
            print employeeList
            unknownEmployees = self.displayUnknownEmployees(employeeList)
            
            if unknownEmployees != None:
                try:
                    if len(unknownEmployees) != 0:
                        '''USE EMPLOYEE LIST TO POPULATE THE PANEL'''
                        mpm.getCAMs() #fills up camList
        
                        newEmpWindow = SubFrame(parent = self, title = "")
                        newEmpWindow.newEmployeeMainFrame(unknownEmployees, mpm.camList)
                        self.Hide()
                    else:
                        pass
                except TypeError:
                    pass
    
    def onAddAccountingCalendar(self, event):
        print 'adding new calendar items'
    
    def onOpenTimesheet(self, event):
        ''' Opens timesheet folder '''
        os.startfile(resourcesDirectory)
        
    def onOpenMPMExports(self, event):
        ''' Opens mpm exports folder '''
        os.startfile(mpmExports)
    
    def onOpenNamerun(self, event):
        ''' opens namerun folder '''
        os.startfile(namerun)
    
    def onOpenReports(self, event):
        ''' Opens reports folder '''
        os.startfile(reports)
    
    def onHelpMain(self, event):
        ''' loads main instuctions frame '''
        title = "Main Instructions"
        filename = open(mainInstructions, 'r')
        msg = filename.read()
        simpleText(self, title, msg, 330, 800)
        filename.close()
        
    def onHelpMPM(self, event):
        ''' loads MPM Requirements frame '''
        title = "MPM Requirements"
        filename = open(mpmRequirements, 'r')
        msg = filename.read()
        simpleText(self, title, msg, 300, 400)
        filename.close()
        
    def onHelpBW(self, event):
        ''' loads BW Requirements frame '''
        title = "BW Requirements"
        filename = open(bwRequirements, 'r')
        msg = filename.read()
        simpleText(self, title, msg, 300, 400)
        filename.close()
        
    def onHelpWiki(self, event):
        '''opens timesheet wiki page'''
        try:
            webbrowser.open(tsaWiki, 0, True)
        except:
            MsgDialog(self, msg = "Cannot open wiki page\nurl: " & tsaWiki, caption = "Load Error", style = wx.OK).msg()
    
    def checkLoads(self):
        """ button checker to enable network generation button and to enable run open reporting button
        
            - when the resources and mpm data have been loaded, then generate networks will be enabled
            - when the resources, mpm, and bw data have been loaded, then reporting button will be enabled 
        """
       
        if self.loadResourcesButton.Enabled == False and self.loadMPMFilesButton.Enabled == False:
            self.generateNetworksButton.Enable(True)
            for each in self.menuItemList[0:2]:
                self.menu.Enable(each, True)
        if self.loadResourcesButton.Enabled == False and self.loadMPMFilesButton.Enabled == False and self.loadNamerunButton.Enabled == False:
            self.openButton.Enable(True)
            for each in self.menuItemList:
                self.menu.Enable(each, True)

#        elif self.loadResourcesButton.Enabled == False or self.loadMPMFilesButton.Enabled == False or self.loadNamerunButton.Enabled == False:
#            print "button off"
#        else:
#            print "all enabled"
    
    def displayCAMChoice(self, camList):
        '''displays a choice box containing the list of cams'''
        msg = "Please choose one or more CAM"
        title = "CAM Options"
        return MultiChoiceDialog(self, msg, title, camList).select(camList)
    
    def displayNetworks(self):
        '''displays a text box containing the list of networks'''
        directions = "1. Copy network list." + "\n2. Run Business Warehouse (BW) report.\n    IMPORTANT FIELDS:\n\tHome Cost Center\n\tNetwork #\n\tNetwork Description\n\tActivity Code\n\tShip Field\n\tSupp Fields 1-3\n\tEmployee ID\n\tEmployee Name\n\tHours Code\n\tWork and Weekending Date\n\tHours Valid" + "\n3. Export as csv."
        title = "Network Generator"
        text = mpm.generateNetworksString()
        #wx.MultiChoiceDialog(msg, title, self.generateNetworksString())
        #egui.textbox(msg, title, text, codebox)

        '''EMPTY NetworkList after string generation'''
        mpm.emptyNetworkList()

        self.generateNetworkFrame = TextFrame(None, title, directions, text, self.onContinue)
        
    def onContinue(self, event):
        """ event handler for network generation: closes the frame and opens the namerun directory """ 
        self.generateNetworkFrame.Close()
        src.openNamerunDirectory()

    def displayUnknownEmployees(self, employeeList):
        '''displays a choice box containing the list of employees'''
        
        msg = "Click OK to continue without a selection." + \
              "\nCancel to return to the previous screen."
        title = "Unknown Employees"
        return MultiChoiceDialog(self, msg, title, employeeList).select(employeeList)
    
    def timeSpan(self, t0, t1):
        return "%.2f" % float(t1-t0) + " seconds"
    
class SubFrame(wx.Frame):
    
    def __init__(self, parent = None, title = ""):
        """ overwrite wx.frame __init__ with id = 2"""
        wx.Frame.__init__(self, parent, 2, title)       

    def boldFont(self): return wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD)
    
    def resizeWH(self, button, width, height): return button.SetSizeWH(width, height)
    
    def textColor(self, button, color = "#000000"): return button.SetForegroundColour(color)

    def backgroundColor(self, button, color = "#000000"): return button.SetBackgroundColour(color)
        
    def reportMainFrame(self):
        """ creates buttons via the following:
                View All                    xls
                View Suspicious Charges     xls
                CAM Summary By Program      ppt
                CAM Summary By Network      ppt
                CAM Summary By Employee     xls
                
                            Close
        """
        
        self.SetTitle("TimeSheet Reports")
        frameWidth = 350
        frameHeight = 450
        framePosition = (frameWidth, frameHeight)
        self.SetSize(framePosition)
        id = wx.ID_ANY
        
        '''REPORT PANEL'''
        self.reportPanel = wx.Panel(self, id, style = wx.NO_BORDER)
        reportPanelWidth = frameWidth
#        reportPanelWidth = frameWidth - 100
        reportPanelHeight = frameHeight - 100
        reportPanelPosition = (reportPanelWidth, reportPanelHeight)
        self.reportPanel.SetSize(reportPanelPosition)
        self.reportPanel.SetBackgroundColour("#003366")    #("#008080")#("#003366")
        self.reportButtons = wx.FlexGridSizer(rows = 5, vgap = 30)
        
        width = 70
        height = 30
        gap = 65
        self.viewAllButton = wx.Button(self.reportPanel, id, "View All", (width, height))
        self.suspiciousChargesButton = wx.Button(self.reportPanel, id, "View Suspicious Charges Only", (width, height + (gap)))
        self.camProgramSummaryButton = wx.Button(self.reportPanel, id, "CAM Summary By Program", (width, height + (2*gap)))
        self.camNetworkSummaryButton = wx.Button(self.reportPanel, id, "CAM Summary By Network", (width, height + (3*gap)))
        self.camEmployeeSummaryButton = wx.Button(self.reportPanel, id, "CAM Summary By Employee", (width, height + (4*gap)))
        
        reportPanelButtonsList = [self.viewAllButton, self.suspiciousChargesButton, self.camEmployeeSummaryButton, self.camNetworkSummaryButton, self.camProgramSummaryButton]
        for eachButton in reportPanelButtonsList:
            self.resizeWH(eachButton, 200, 30)
            eachButton.SetBackgroundColour(wx.WHITE)
            eachButton.SetFont(self.boldFont())
            self.reportButtons.Add(eachButton)
        
#        # EXPORT PANEL
#        self.exportPanel = wx.Panel(self, id, style = wx.NO_BORDER)
#        exportPanelWidth = frameWidth - 250
#        exportPanelHeight = frameHeight - 100
#        exportPanelPosition = (exportPanelWidth, exportPanelHeight)      
#        self.exportPanel.SetSize(exportPanelPosition)
#        self.exportPanel.SetBackgroundColour("#003366")    #("#008080")#("#003366")
#        self.exportButtons = wx.FlexGridSizer(rows = 5, vgap = 30)
#        
#        self.exportViewAllButton = wx.Button(self.exportPanel, id, "xls", (width, height))
#        self.exportSuspiciousChargesButton = wx.Button(self.exportPanel, id, "xls", (width, height + (gap)))
#        self.exportProgramSummaryButton = wx.Button(self.exportPanel, id, "ppt", (width, height + (2*gap)))
#        self.exportNetworkSummaryButton = wx.Button(self.exportPanel, id, "ppt", (width, height + (3*gap)))
#        self.exportEmployeeSummaryButton = wx.Button(self.exportPanel, id, "xls", (width, height + (4*gap)))
#        
#        exportPanelButtonsList = [self.exportViewAllButton, self.exportSuspiciousChargesButton, self.exportProgramSummaryButton, self.exportNetworkSummaryButton, self.exportEmployeeSummaryButton]
#        for eachButton in exportPanelButtonsList:
#            self.resizeWH(eachButton, 30, 30)
#            eachButton.SetBackgroundColour(wx.WHITE)
#            eachButton.SetFont(self.boldFont())
#            self.exportButtons.Add(eachButton)
        
        '''CLOSE PANEL'''
        self.closePanel = wx.Panel(self, id, style = wx.SIMPLE_BORDER)
        closePanelWidth = frameWidth
        closePanelHeight = frameHeight - reportPanelHeight
        closePanelPosition = (closePanelWidth, closePanelHeight)
        self.closePanel.SetSize(closePanelPosition)
        self.closePanel.SetBackgroundColour(wx.WHITE)
        self.closeButtons = wx.FlexGridSizer(rows = 5, vgap = 30)

        position = (closePanelWidth/3.5, closePanelHeight/4)
        self.closeButton = wx.Button(self.closePanel, wx.ID_EXIT, "PREVIOUS SCREEN", position) #ID_EXIT = EXIT Program
        self.closeButton.SetSize((150, 30))
        self.backgroundColor(self.closeButton, "#800000")
        self.closeButton.SetFont(self.boldFont())
        self.textColor(self.closeButton, wx.WHITE)

        self.closeButtons.Add(self.closeButton)               

        '''BIND BUTTONS'''
        self.Bind(wx.EVT_BUTTON, self.onViewAll, self.viewAllButton)
        self.Bind(wx.EVT_BUTTON, self.onSuspiciousCharges, self.suspiciousChargesButton)
        self.Bind(wx.EVT_BUTTON, self.onProgramSummary, self.camProgramSummaryButton)
        self.Bind(wx.EVT_BUTTON, self.onNetworkSummary, self.camNetworkSummaryButton)
        self.Bind(wx.EVT_BUTTON, self.onEmployeeSummary, self.camEmployeeSummaryButton)
#        self.Bind(wx.EVT_BUTTON, self.onExportViewAll, self.exportViewAllButton)
#        self.Bind(wx.EVT_BUTTON, self.onExportSuspiciousCharges, self.exportSuspiciousChargesButton)
#        self.Bind(wx.EVT_BUTTON, self.onExportProgramSummary, self.exportProgramSummaryButton)
#        self.Bind(wx.EVT_BUTTON, self.onExportNetworkSummary, self.exportNetworkSummaryButton)
#        self.Bind(wx.EVT_BUTTON, self.onExportEmployeeSummary, self.exportEmployeeSummaryButton)
        self.Bind(wx.EVT_BUTTON, self.onExit, self.closeButton, wx.ID_EXIT)
        
        '''ENABLE/DISABLE BUTTONS'''
        self.camNetworkSummaryButton.Enable(False)
        self.camEmployeeSummaryButton.Enable(False)
        
        '''POSITION WITHIN FRAME'''
        self.CenterOnScreen()
        #self.exportPanel.Move((reportPanelWidth,0))
        self.closePanel.Move((0, reportPanelHeight))
        self.Show()

    def onViewAll(self, event): 
        '''provides an all-up view of the data. format:
            Program1   Network1    Emp1
                                   Emp2
                                   .
                                   .
                                   .
                       Network2    Emp1
                                   Emp2
                                   .
                                   .
                                   .
        '''
        
        ''' excel export '''
        t0 = time.time()
        
        '''OBTAIN ACCOUNTING MONTH NOW BEFORE CSV GENERATION'''
        filterMonth = MultiChoiceDialog(self, "Please choose an Accounting Month", "Filter Accounting Month", bw.list_Period).select(bw.list_Period)
        
        if filterMonth <> None:
            #REINSTANTIATE excel
#            xl.__init__()
            
            '''OBTAIN ACTIVITY FILTERS BEFORE RUNNING CSV GENERATION'''
            listActivities = self.insertSelectAll(bw.list_Activity)
            filterActivities = MultiChoiceDialog(self, "Please choose your Activity Codes", "Filter Activity Codes", listActivities).select(listActivities)
            
            '''OBTAIN COLORSCHEME; MULTICHOICEDIALOG RETURNS A LIST THEREFORE, WE HAVE TO TAKE OUT THE VALUE'''
            cellColor = MultiChoiceDialog(self, "Please choose a color theme", "Color Theme", xl.listColors()).select(xl.listColors())
            try:
                cellColor = cellColor[0]
            except: #works for both TypeError and IndexError
                cellColor = "None"
    
            pivotStyle = xl.dict_PivotStyles()[cellColor]
            try:
                pivotColor = xl.convertRGB(xl.dict_colorChoice()[cellColor])
            except KeyError:
                pivotColor = wx.WHITE
            
            xlfilename = self.dateString()+'NamerunAllUp'
            xlFile = src.createFilePath(reports, xlfilename + '.xlsx')
            csvFile = src.createFilePath(reports, self.dateString()+'NamerunAllUp.csv')
            
            '''SHOULD I WRITE THE CSV HERE?'''
            bw.toCSVFile(csvFile)        
            convertFile = src.convertCSVToXLSX(csvFile, xlFile)
            
#            print convertFile
            wb = xl.openExcelFile(convertFile)
            
            '''CREATE PIVOT TABLE; FIRST SET PARAMETERS'''
            # number of rows = length of namerunlist + 1, number of columns = length of 1 namerun
            colEnd = len(bw.namerun_list[0].header().split(","))
            datarange = xl.setRange(wb, xlfilename, cellStart = (1,1), cellEnd = (len(bw.namerun_list)+1, colEnd))
            datasource = "'" + xlfilename + "'!" + datarange.Address
#            print datarange.Address
            
            filters = ("Accounting Month",) #extra comma necessary to indicate tuple element
            filterValues = filterMonth
            columns = ("Weekend Date",) #extra comma necessary to indicate tuple element
            rows = ("Program", "Network", "Description", "TotalHours", "ETC", "Percent Spent", "Activity", "Employee CAM Code", "Employee Name", "Hours Code", "Charge Comment")
            subtotalrows = (False, True, False, False, False, False, False, False, False, False, False)
            sumvalue = "Sum of Hours Valid"
            sortfield = ()
    
            '''FUNCTION CALL BELOW RETURNS [worksheet, pivot table]'''
#            pivotws = xl.addPivot(wb, "'2012-09-24NamerunAllUp'!$A$1:$AE$130951", "PivotSummary", filters, filterValues, columns, rows, subtotalrows, sumvalue, sortfield)
#            print datasource
#            print "press enter..."
#            raw_input()
            pivotws = xl.addPivot(wb, datasource, "PivotSummary", filters, filterValues, columns, rows, subtotalrows, sumvalue, sortfield)
            pivotWorksheet = pivotws[0]
            pivotTable = pivotws[1]

            '''COMMENT CURRENT PERCENT SPENT'''
            xl.visible()
            
            strPercentSpent = "SHOULD BE: " + str(self.getPercentSpentToDate()) + " SPENT FOR " + str(wcal.getPeriod(wcal.getCurrentWeekEndingDate()))
            xl.addComment(pivotWorksheet, "$F$2", strPercentSpent, 30, 90, True)
         
            for row, TFValue in zip(rows, subtotalrows):
                xl.pivotSubTotal(pivotTable, row, TFValue)
            
            for filter, filterValue in zip(filters, filterValues):
                xl.pivotFilter(pivotTable, filter, filterValue)
            
            '''CHOOSE THE FILTERED ITEMS WE WANT TO SHOW'''
            try:
                if "Select All" not in filterActivities: 
                    filterItem = ("Activity",) #extra comma necessary to indicate tuple element
                    filterItemValues = filterActivities
                    xl.pivotItemFilter(pivotTable, filterItem[0], bw.list_Activity, filterItemValues)
            except:
                pass
            
            #xl.colorRange(pivotWorksheet.Cells, 2)  <== changes all cells to white
            xl.pivotTheme(pivotTable, pivotStyle)   
    
            '''APPLY CONDITIONAL FORMATTING'''
            '''NO ETC'''
            stringRange = '$K:$K'
            string = 'No ETC'
            interior = xl.convertRGB((192, 0, 0)) #RED
            fontColor = 2 #fontColor = xl.convertRGB((255, 255, 255)) #WHITE; need to use color index
            xl.conditionUsingString(pivotWorksheet, stringRange, string, interior, fontColor, 1)
            '''CHECK ACTIVITY CODE'''
            stringRange = '$K:$K'
            string = 'Check Activity Code'
            interior = xl.convertRGB((255, 231, 155)) #YELLOW/BROWN
            fontColor = 1 #fontColor = xl.convertRGB((0, 0, 0)) #BLACK; need to use color index
            xl.conditionUsingString(pivotWorksheet, stringRange, string, interior, fontColor, 2)
            '''SUBTOTALS'''
            formulaRange = 'B:K'
            curFormula = '=RIGHT($B1,5)="Total"'
            xl.conditionUsingFormula(pivotWorksheet, formulaRange, curFormula, 3)        
            '''PIVOT DATA FIELD'''
            #scopeFormula = '=RIGHT(OFFSET($A5,0,1),5)="Total"'
            scopeRange = '$L5'
            xl.conditionPivotScopeUsingFormula(pivotWorksheet, scopeRange, curFormula, 4)
            '''CHARGE COMMENT ROW'''
            rowFieldName = "'Charge Comment'[All]"
            xl.conditionPivotSelect(pivotTable, rowFieldName, "Good")
            '''PERCENT SPENT FORMAT'''
            rowFieldName2 = "'Percent Spent'[All]"
            xl.conditionPivotSelect(pivotTable, rowFieldName2, "Percent")
            
            
            xl.freezePane(pivotWorksheet, "$D$2")
            
            #xl.visible()
        else:
            MsgDialog(self, "Data cannot be generated without Accounting Month.\nPlease choose an Accounting Month.", "None Selected", wx.OK).msg()
        
        '''RECORD TIME'''
        t1 = time.time()
        span = self.timeSpan(t0, t1)
        MsgDialog(None, msg = ("Load Complete.\n" + span), caption = "Message", style = wx.OK).msg()
            
    def onSuspiciousCharges(self, event): 
        ''' provides a view of suspicious chargers.  same view as the All Up View, but of suspicious chargers only'''
        
        t0 = time.time()
        
        '''OBTAIN ACCOUNTING MONTH NOW BEFORE CSV GENERATION'''
        filterMonth = MultiChoiceDialog(self, "Please choose an Accounting Month", "Filter Accounting Month", bw.list_Period).select(bw.list_Period)
        
        if filterMonth <> None:
            
            '''OBTAIN ACTIVITY FILTERS BEFORE RUNNING CSV GENERATION'''
            listActivities = self.insertSelectAll(bw.list_Activity)
            filterActivities = MultiChoiceDialog(self, "Please choose your Activity Codes", "Filter Activity Codes", listActivities).select(listActivities)
            
            '''OBTAIN COLORSCHEME; MULTICHOICEDIALOG RETURNS A LIST THEREFORE, WE HAVE TO TAKE OUT THE VALUE'''
            cellColor = MultiChoiceDialog(self, "Please choose a color theme", "Color Theme", xl.listColors()).select(xl.listColors())
            try:
                cellColor = cellColor[0]
            except: #works for both TypeError and IndexError
                cellColor = "None"
    
            pivotStyle = xl.dict_PivotStyles()[cellColor]
            try:
                pivotColor = xl.convertRGB(xl.dict_colorChoice()[cellColor])
            except KeyError:
                pivotColor = wx.WHITE
            
    #        raw_input()
 
            xlfilename = self.dateString()+'SuspiciousCharges'
            xlFile = src.createFilePath(reports, xlfilename + '.xlsx')
            csvFile = src.createFilePath(reports, self.dateString()+'SuspiciousCharges.csv')    
    
            '''SHOULD I WRITE THE CSV HERE?'''
            bw.toCSVFile(csvFile)        
            convertFile = src.convertCSVToXLSX(csvFile, xlFile)
#            print convertFile
            wb = xl.openExcelFile(convertFile)
            
            '''CREATE PIVOT TABLE; FIRST SET PARAMETERS'''
            # number of rows = length of namerunlist + 1, number of columns = length of 1 namerun
            colEnd = len(bw.namerun_list[0].header().split(","))
            datarange = xl.setRange(wb, xlfilename, cellStart = (1,1), cellEnd = (len(bw.namerun_list)+1, colEnd))
            datasource = "'" + xlfilename + "'!" + datarange.Address
            
            #sourcedata = "NamerunAllUp" + "!" + str(xl.rangeAddress(datarange))        
            filters = ("Accounting Month",) #extra comma necessary to indicate tuple element
            filterValues = filterMonth
            columns = ("Weekend Date",) #extra comma necessary to indicate tuple element
            rows = ("Program", "Network", "Description", "Activity", "Employee CAM Code", "Employee Name", "Hours Code", "Charge Comment")
            subtotalrows = (False, True, False, False, False, False, False, False)
            sumvalue = "Sum of Hours Valid"
            sortfield = ()
    
            '''FUNCTION CALL BELOW RETURNS [worksheet, pivot table]'''
            pivotws = xl.addPivot(wb, datasource, "PivotSummary", filters, filterValues, columns, rows, subtotalrows, sumvalue, sortfield)
            pivotWorksheet = pivotws[0]
            pivotTable = pivotws[1]
         
            for row, TFValue in zip(rows, subtotalrows):
                xl.pivotSubTotal(pivotTable, row, TFValue)
            
            for filter, filterValue in zip(filters, filterValues):
                xl.pivotFilter(pivotTable, filter, filterValue)
            
            '''CHOOSE THE FILTERED ITEMS WE WANT TO SHOW'''
            try:
                if "Select All" not in filterActivities: 
                    filterItem = ("Activity",) #extra comma necessary to indicate tuple element
                    filterItemValues = filterActivities
                    xl.pivotItemFilter(pivotTable, filterItem[0], bw.list_Activity, filterItemValues)
            except:
                pass
    
            filterItem = ("Charge Comment",) #extra comma necessary to indicate tuple element
            filterItemValues = ["No ETC", "Check Activity Code", "No CAM Code"]
            xl.pivotItemFilter(pivotTable, filterItem[0], bw.list_ChargeComments, filterItemValues)
            
            #xl.colorRange(pivotWorksheet.Cells, 2)  <== changes all cells to white
            xl.pivotTheme(pivotTable, pivotStyle)   
    
            '''APPLY CONDITIONAL FORMATTING'''
            '''NO ETC'''
            stringRange = '$H:$H'
            string = 'No ETC'
            interior = xl.convertRGB((192, 0, 0)) #RED
            fontColor = 2 #fontColor = xl.convertRGB((255, 255, 255)) #WHITE; need to use color index
            xl.conditionUsingString(pivotWorksheet, stringRange, string, interior, fontColor, 1)
            '''CHECK ACTIVITY CODE'''
            stringRange = '$H:$H'
            string = 'Check Activity Code'
            interior = xl.convertRGB((255, 231, 155)) #YELLOW/BROWN
            fontColor = 1 #fontColor = xl.convertRGB((0, 0, 0)) #BLACK; need to use color index
            xl.conditionUsingString(pivotWorksheet, stringRange, string, interior, fontColor, 2)
            '''SUBTOTALS'''
            formulaRange = 'B:H'
            curFormula = '=RIGHT($B1,5)="Total"'
            xl.conditionUsingFormula(pivotWorksheet, formulaRange, curFormula, 3)        
            '''PIVOT DATA FIELD'''
            #scopeFormula = '=RIGHT(OFFSET($A5,0,1),5)="Total"'
            scopeRange = '$I5'
            xl.conditionPivotScopeUsingFormula(pivotWorksheet, scopeRange, curFormula, 4)
            '''CHARGE COMMENT ROW'''
            rowFieldName = "'Charge Comment'[All]"
            xl.conditionPivotSelect(pivotTable, rowFieldName, "Good")
            
            xl.visible()
        else:
            MsgDialog(self, "Data cannot be generated without Accounting Month.\nPlease choose an Accounting Month.", "None Selected", wx.OK).msg()
        
        '''RECORD TIME'''
        t1 = time.time()
        span = self.timeSpan(t0, t1)
        MsgDialog(None, msg = ("Load Complete.\n" + span), caption = "Message", style = wx.OK).msg()
    
    def onProgramSummary(self, event): 
        ''' program summary view for each CAM:
                            Current Month
            Program1    Total Actuals    ETC    %Spent
            Program2
        '''
        
        t0 = time.time()
        
        '''OBTAIN ACCOUNTING MONTH NOW BEFORE CSV GENERATION'''
        filterMonth = MultiChoiceDialog(self, "Please choose an Accounting Month", "Filter Accounting Month", bw.list_Period).select(bw.list_Period)
        
        if filterMonth <> None:

            '''OBTAIN ACTIVITY FILTERS BEFORE RUNNING CSV GENERATION'''
            listActivities = self.insertSelectAll(bw.list_Activity)
            filterActivities = MultiChoiceDialog(self, "Please choose your Activity Codes", "Filter Activity Codes", listActivities).select(listActivities)
            
            '''#OBTAIN COLORSCHEME; MULTICHOICEDIALOG RETURNS A LIST THEREFORE, WE HAVE TO TAKE OUT THE VALUE'''
            cellColor = MultiChoiceDialog(self, "Please choose a color theme", "Color Theme", xl.listColors()).select(xl.listColors())
            try:
                cellColor = cellColor[0]
            except: #works for both TypeError and IndexError
                cellColor = "None"
    
            pivotStyle = xl.dict_PivotStyles()[cellColor]
            try:
                pivotColor = xl.convertRGB(xl.dict_colorChoice()[cellColor])
            except KeyError:
                pivotColor = wx.WHITE
            
    #        raw_input()
            xlfilename = self.dateString()+'ProgramSummary'
            xlFile = src.createFilePath(reports, xlfilename + '.xlsx')
            csvFile = src.createFilePath(reports, self.dateString()+'ProgramSummary.csv')              
#            xlFile = src.createFilePath(reports, 'ProgramSummary.xlsx')
#            csvFile = src.createFilePath(reports, 'ProgramSummary.csv')
    
            '''SHOULD I WRITE THE CSV HERE?'''
            bw.toCSVFile(csvFile)        
            convertFile = src.convertCSVToXLSX(csvFile, xlFile)
            wb = xl.openExcelFile(convertFile)
            
            '''CREATE PIVOT TABLE; FIRST SET PARAMETERS'''
            # number of rows = length of namerunlist + 1, number of columns = length of 1 namerun
            colEnd = len(bw.namerun_list[0].header().split(","))
            datarange = xl.setRange(wb, xlfilename, cellStart = (1,1), cellEnd = (len(bw.namerun_list)+1, colEnd))
            datasource = "'" + xlfilename + "'!" + datarange.Address
            #sourcedata = "NamerunAllUp" + "!" + str(xl.rangeAddress(datarange))        
            filters = ("Accounting Month",) #extra comma necessary to indicate tuple element
            filterValues = filterMonth
            columns = ("Weekend Date",) #extra comma necessary to indicate tuple element
            rows = ("Program",)
            subtotalrows = (False,)
            sumvalue = "Sum of Hours Valid"
            sortfield = ()
    
            '''FUNCTION CALL BELOW RETURNS [worksheet, pivot table]'''
            pivotws = xl.addPivot(wb, datasource, "PivotSummary", filters, filterValues, columns, rows, subtotalrows, sumvalue, sortfield)
            pivotWorksheet = pivotws[0]
            pivotTable = pivotws[1]
         
    #        for row, TFValue in zip(rows, subtotalrows):
    #            xl.pivotSubTotal(pivotTable, row, TFValue)
            
            for filter, filterValue in zip(filters, filterValues):
                xl.pivotFilter(pivotTable, filter, filterValue)
            
            '''CHOOSE THE FILTERED ITEMS WE WANT TO SHOW'''
            try:
                if "Select All" not in filterActivities: 
                    filterItem = ("Activity",) #extra comma necessary to indicate tuple element
                    filterItemValues = filterActivities
                    xl.pivotItemFilter(pivotTable, filterItem[0], bw.list_Activity, filterItemValues)
            except:
                pass
    
            #xl.colorRange(pivotWorksheet.Cells, 2)  <== changes all cells to white
            xl.pivotTheme(pivotTable, pivotStyle)   
    
    #        # APPLY CONDITIONAL FORMATTING
    #        # NO ETC
    #        stringRange = '$H:$H'
    #        string = 'No ETC'
    #        interior = xl.convertRGB((192, 0, 0)) #RED
    #        fontColor = 2 #fontColor = xl.convertRGB((255, 255, 255)) #WHITE; need to use color index
    #        xl.conditionUsingString(pivotWorksheet, stringRange, string, interior, fontColor, 1)
    #        # CHECK ACTIVITY CODE
    #        stringRange = '$H:$H'
    #        string = 'Check Activity Code'
    #        interior = xl.convertRGB((255, 231, 155)) #YELLOW/BROWN
    #        fontColor = 1 #fontColor = xl.convertRGB((0, 0, 0)) #BLACK; need to use color index
    #        xl.conditionUsingString(pivotWorksheet, stringRange, string, interior, fontColor, 2)
    #        # SUBTOTALS
    #        formulaRange = 'B:H'
    #        curFormula = '=RIGHT($B1,5)="Total"'
    #        xl.conditionUsingFormula(pivotWorksheet, formulaRange, curFormula, 3)        
    #        # PIVOT DATA FIELD
    #        #scopeFormula = '=RIGHT(OFFSET($A5,0,1),5)="Total"'
    #        scopeRange = '$I5'
    #        xl.conditionPivotScopeUsingFormula(pivotWorksheet, scopeRange, curFormula, 4)
    #        # CHARGE COMMENT ROW
    #        rowFieldName = "'Charge Comment'[All]"
    #        xl.conditionPivotSelect(pivotTable, rowFieldName, "Good")
            
            xl.visible()
        else:
            MsgDialog(self, "Data cannot be generated without Accounting Month.\nPlease choose an Accounting Month.", "None Selected", wx.OK).msg()
    
        '''RECORD TIME'''
        t1 = time.time()
        span = self.timeSpan(t0, t1)
        MsgDialog(None, msg = ("Load Complete.\n" + span), caption = "Message", style = wx.OK).msg()
    
    def onNetworkSummary(self, event): 
        ''' network summary by program and by CAM:
                            Current Month
            Program1    Network1     Total Actuals    ETC    %Spent    Open/Closed
            Program2    Network1     Total Actuals    ETC    %Spent    Open/Closed

        '''
        
        print "network"
    
    def onEmployeeSummary(self, event): 
        ''' employee summary
            Employee1    Program1    Network1
                         Program2    Network1
            Employee2    Program1    Network1
                                     Network2
        '''
        
        emp.toHTMLFile("employeelist.html")
            
    def onExit(self, event):
        self.Destroy()
        self.Parent.Show()
        #DriverApp(False).MainLoop()
    
    def getPercentSpentToDate(self):
        ''' returns the string value for the percent spent'''
        curWeekEndingDate = wcal.getCurrentWeekEndingDate()
        #print curWeekEndingDate
        curPeriod = wcal.getPeriod(curWeekEndingDate)
        #print curPeriod
        totalHours = mcal.getPeriodHours(int(curPeriod))
        
#        val = float(wcal.getHoursToDate()/totalHours*100)
#        "%.2f" % val
#        return str(val) + "%"
        
        try:
            val = float(wcal.getHoursToDate()/totalHours*100)
            "%.2f" % val
            return str(val) + "%"
        except ValueError:
            return "Not Found in Acct Calendar"
            
    def newEmployeeMainFrame(self, employeeList = [], camList = []):
        ''' this method creates the GUI window to add new employees'''
        self.SetTitle("Add New Employees")
        frameWidth = 300
        increment = 10
        frameHeight = 150
        titlePanelHeight = 25
        totalHeight = frameHeight + titlePanelHeight
        self.SetSize((frameWidth, totalHeight))
        id = wx.ID_ANY
        
        '''CLOSE PANEL'''
        self.closePanel = wx.Panel(self, id, style = wx.SIMPLE_BORDER)
        closePanelHeight = 50
        self.closePanel.SetSize((frameWidth, closePanelHeight))
        self.closePanel.SetBackgroundColour(wx.WHITE) 
        self.closeButtons = wx.FlexGridSizer()
                
        position = (frameWidth/1.8, closePanelHeight/4)
        self.closeButton = wx.Button(self.closePanel, wx.ID_EXIT, "EXIT", position) #ID_EXIT = EXIT Program
        self.backgroundColor(self.closeButton, "#800000")
        
        position = (frameWidth/5.5, closePanelHeight/4)
        self.AddButton = wx.Button(self.closePanel, wx.ID_OPEN, "ADD", position) #ID_EXIT = EXIT Program
        self.backgroundColor(self.AddButton, "#003300")
        
        closePanelButtonsList = [self.closeButton, self.AddButton]
        for eachButton in closePanelButtonsList:
            eachButton.SetFont(self.boldFont())
            self.textColor(eachButton, wx.WHITE)
            self.closeButtons.Add(eachButton)
            
        '''BIND EXIT AND ADD BUTTONS TO EVENTS'''
        self.Bind(wx.EVT_BUTTON, self.onCloseNewEmployee, self.closeButton, wx.ID_EXIT)
        self.Bind(wx.EVT_BUTTON, self.onAdd, self.AddButton)
        
        '''LIST PANEL'''
        self.listPanel = scroll.ScrolledPanel(self, id, style = wx.TAB_TRAVERSAL | wx.SUNKEN_BORDER)
        listPanelHeight = frameHeight - closePanelHeight
        self.listPanel.SetSize((frameWidth-10, listPanelHeight))
        self.listPanel.SetBackgroundColour("#003366")        
        #self.panelBoxes = wx.FlexGridSizer(cols = 2, rows = 1)
        self.panelBoxes = wx.BoxSizer(wx.HORIZONTAL)
                
        emptyText = wx.StaticText(self.listPanel, id, "AddNew")
        emptyText.SetForegroundColour("#003366")
        self.empName = "Name"
        self.employeeList = employeeList
        position = (20, increment)
        self.empBox = wx.ComboBox(self.listPanel, id, self.empName, position, choices = self.employeeList, style = wx.CB_DROPDOWN)
        self.camCode = "CAM"
        position = (170, increment)
        self.camBox = wx.ComboBox(self.listPanel, id, self.camCode, position, choices = camList, style = wx.CB_DROPDOWN)
        boxList = [emptyText, self.empBox, self.camBox]
        
        for eachBox in boxList:
            self.panelBoxes.Add(eachBox, 0, wx.ALIGN_CENTER|wx.ALL)
            
        '''BIND COMBO BOXES TO EVENTS:'''
        self.Bind(wx.EVT_TEXT, self.onSelectCAM, self.camBox)
        self.Bind(wx.EVT_TEXT, self.onSelectName, self.empBox)
        
        self.listPanel.SetSizer(self.panelBoxes)
        self.listPanel.Layout()
        self.listPanel.SetupScrolling()
        
        #POSITION WITHIN FRAME
        self.CenterOnScreen()
        self.closePanel.Move((0, listPanelHeight))
        self.checkAdd()
        self.Show()

    def editEmployeeMainFrame(self, employeeList = [], camList = []):
        ''' this method creates the GUI window to edit current employees'''
        self.SetTitle("Edit Current Employees")
        frameWidth = 300
        increment = 10
        frameHeight = 150
        titlePanelHeight = 25
        totalHeight = frameHeight + titlePanelHeight
        self.SetSize((frameWidth, totalHeight))
        id = wx.ID_ANY
        
        #CLOSE PANEL
        self.closePanel = wx.Panel(self, id, style = wx.SIMPLE_BORDER)
        closePanelHeight = 50
        self.closePanel.SetSize((frameWidth, closePanelHeight))
        self.closePanel.SetBackgroundColour(wx.WHITE) 
        self.closeButtons = wx.FlexGridSizer()
                
        position = (frameWidth/1.8, closePanelHeight/4)
        self.closeButton = wx.Button(self.closePanel, wx.ID_EXIT, "EXIT", position) #ID_EXIT = EXIT Program
        self.backgroundColor(self.closeButton, "#800000")
        
        position = (frameWidth/5.5, closePanelHeight/4)
        self.AddButton = wx.Button(self.closePanel, wx.ID_OPEN, "EDIT", position) #ID_EXIT = EXIT Program
        self.backgroundColor(self.AddButton, "#003300")
        
        closePanelButtonsList = [self.closeButton, self.AddButton]
        for eachButton in closePanelButtonsList:
            eachButton.SetFont(self.boldFont())
            self.textColor(eachButton, wx.WHITE)
            self.closeButtons.Add(eachButton)
            
        #BIND EXIT AND ADD BUTTONS TO EVENTS
        self.Bind(wx.EVT_BUTTON, self.onCloseEditEmployee, self.closeButton, wx.ID_EXIT)
        self.Bind(wx.EVT_BUTTON, self.onEdit, self.AddButton)
        
        #LIST PANEL
        self.listPanel = scroll.ScrolledPanel(self, id, style = wx.TAB_TRAVERSAL | wx.SUNKEN_BORDER)
        listPanelHeight = frameHeight - closePanelHeight
        self.listPanel.SetSize((frameWidth-10, listPanelHeight))
        self.listPanel.SetBackgroundColour("#003366")        
        #self.panelBoxes = wx.FlexGridSizer(cols = 2, rows = 1)
        self.panelBoxes = wx.BoxSizer(wx.HORIZONTAL)
                
        emptyText = wx.StaticText(self.listPanel, id, "EditCurrent")
        emptyText.SetForegroundColour("#003366")
        self.empName = "Name"
        self.employeeList = employeeList
        position = (20, increment)
        self.empBox = wx.ComboBox(self.listPanel, id, self.empName, position, choices = self.employeeList, style = wx.CB_DROPDOWN)
        self.camCode = "CAM"
        position = (170, increment)
        self.camBox = wx.ComboBox(self.listPanel, id, self.camCode, position, choices = camList, style = wx.CB_DROPDOWN)
        boxList = [emptyText, self.empBox, self.camBox]
        
        for eachBox in boxList:
            self.panelBoxes.Add(eachBox, 0, wx.ALIGN_CENTER|wx.ALL)
            
        #BIND COMBO BOXES TO EVENTS:
        self.Bind(wx.EVT_TEXT, self.onSelectCAM, self.camBox)
        self.Bind(wx.EVT_TEXT, self.onSelectName, self.empBox)
        
        self.listPanel.SetSizer(self.panelBoxes)
        self.listPanel.Layout()
        self.listPanel.SetupScrolling()
        
        #POSITION WITHIN FRAME
        self.CenterOnScreen()
        self.closePanel.Move((0, listPanelHeight))
        self.checkAdd()
        self.Show()
    
    def onSelectCAM(self, event):
        '''captures the text the user chooses'''
        #print event.GetString()
        self.camCode = event.GetString()
        self.checkAdd()
        
    def onSelectName(self, event):
        '''captures the text the user chooses'''
        #print event.GetString()
        self.empName = event.GetString()
        self.checkAdd()
        
    def checkAdd(self):
        ''' method to enable or disable the addButton'''
        
        if self.empName != "Name" and self.camCode != "CAM":
            self.AddButton.Enable(True)
        else:
            self.AddButton.Enable(False)

    def onCloseNewEmployee(self, event):
        ''' this method will scrap the window from the buffer; 
            it will then check if there are new employees stored in the dictionary; if so, 
            then it will append it the data and reload the Employee class and Namerun class'''
        
        self.Destroy()
        
        ''''check if the new employee dictionary contains any data, if so add the new data to source''' 
        if len(emp.newEmployees_dictionary) != 0:
            src.reloadEmployee(empFilename, emp.employee_list)
        
#        ''''reload employee and namerun sources'''
#        if src.copyResources(mappingResources, resources) != False:
#            self.Parent.loadEmployeeFile()
#            self.Parent.loadNamerunMappings()
        
        self.Parent.Refresh()
        self.Parent.Show()
    
    def onCloseEditEmployee(self, event):
        ''' scrap current window, reload employee data, show parent'''
        
        self.Destroy()
        src.reloadEmployee(empFilename, emp.employee_list)
        self.Parent.Refresh()
        self.Parent.Show()
    
    def onAdd(self, event):
        ''' this method will call an employees class method and store the information'''
        
        if self.empName != 'Name' and self.camCode != 'CAM': 
            emp.storeNewEmployee(self.empName, self.camCode)
            self.refreshAddNewEmp()
            self.checkAdd()
        else:
            MsgDialog(self, msg = 'code "CAM" not allowed', caption = 'Message', style = wx.OK).msg()
        
    def onEdit(self, event):
        ''' 1. look up the employee's ID via name??? <== not implemented right now
            2. change his cam to the new cam code
        '''
        emp.editCurrentEmployee(self.empName, self.camCode)
        self.refreshAddNewEmp()
        self.checkAdd()
    
    def refreshAddNewEmp(self):
        ''' this method 'refreshes' the window by removing the added name from the list, 
            updating the list, resetting the default values of self.empName and self.camCode
            '''
        self.employeeList.remove(self.empName) #remove added name from the list
        self.empName = "Name" #reset to default value
        self.camCode = "CAM" #reset to default value
        self.empBox.Clear() #clear combo box
        self.empBox.AppendItems(self.employeeList) #add updated list
        if len(self.employeeList) == 0: #if no more employees to add, clear cam list
            self.camBox.Clear()
        
    def timeSpan(self, t0, t1): return "%.2f" % float(t1-t0) + " seconds"
    
    def chooseMonth(self, listArg): return MultiChoiceDialog(self, "Please choose an Accounting Month", "Filter Accounting Month", listArg).select(listArg)
        
    def chooseActivityCodes(self, listArg): return MultiChoiceDialog(self, "Please choose your Activity Codes", "Filter Activity Codes", listArg).select(listArg)
    
    def chooseColor(self):
        #obtain colorscheme; MultiChoiceDialog returns a list therefore, we have to take out the value
        cellColor = MultiChoiceDialog(self, "Please choose a color theme", "Color Theme", xl.listColors()).select(xl.listColors())
        return cellColor[0]  
    
    def insertSelectAll(self, listarg):
        """inserts a Select All option to a list"""
        
        if "Select All" not in listarg:
            listarg.insert(0, "Select All")
        return listarg

    def dateString(self): return str(datetime.date.today())
###################################
# RUN GUI
###################################

if __name__ == "__main__":
    DriverApp(False).MainLoop()
