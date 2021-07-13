'''
Created on Apr 9, 2012

@author: Joselle Abagat, Daniel McDonald
'''

import os, shutil, re, csv, wx, stat
import win32com.client as win32
import easygui as egui
from clsBWNamerun import *
from clsCAM import *
from clsContract import *
from clsEmployee import *
from clsETC import *
from clsMonthlyCalendar import *
from clsNetwork import *
from clsWeeklyCalendar import *
from clsGUI import *

class ReloadResources(object):
    
    def __init__(self, mappingDir, mainDir, mpmExportsDir, namerunDir, reportDir, guideDir, resourcesDir, mappingUserGuide, mappingResources):
        self.mainDirectory = mainDir
        self.mappingDirectory = mappingDir
        self.currentDirectory = os.getcwd()
        self.resourcesDirectory = resourcesDir
        self.mpmExports = mpmExportsDir
        self.namerun = namerunDir
        self.reports = reportDir
        self.userguide = guideDir
        self.mappingUserGuide = mappingUserGuide
        self.mappingResources = mappingResources
        self.directoryCheck = self.checkDirectories()
        
        self.converted = []
        self.notConverted = []
                
    def checkDirectories(self):
        """
        checks the following directories:   MAPPING = must be on a mapped drive
                                            resources = creates it if it doesn't exist
                                            MPMexports = creates it if it doesn't exist
                                            Namerun = creates it if it doesn't exist
        """
        
        title = "Directory Check"
        ok_button ="OK"
        dirList = [self.mainDirectory, self.resourcesDirectory, self.mpmExports, self.namerun, self.reports, self.userguide]
        if not self.testDirectory(self.mappingDirectory) == True:
            msg = "Check that Directory:\n\n" + os.path.abspath(self.mappingDirectory) + "\n\nexists or is mapped correctly."
            egui.msgbox(msg, title, ok_button, None, None)

        for each in dirList:
            if not self.testDirectory(each) == True:
                os.makedirs(each, 0777)
                os.chmod(each, stat.S_IWGRP)
                os.chmod(each, stat.S_IWRITE)
        
#        if not self.testDirectory(self.mainDirectory) == True:
#            os.mkdir(self.mainDirectory)
#            os.chmod(self.mainDirectory, stat.S_IWGRP)
#        if not self.testDirectory(self.resourcesDirectory) == True:
#            os.mkdir(self.resourcesDirectory)
#            os.chmod(self.resourcesDirectory, stat.S_IWGRP)
#            #msg = "Check that Directory:\n\n" + os.path.abspath(self.resourcesDirectory) + " \n\nexists or is mapped correctly."
#            #egui.msgbox(msg, title, ok_button, None, None)
#            #print("Check that Directory: " + os.path.abspath(self.resourcesDirectory) + " exists or is mapped correctly.")
#        if not self.testDirectory(self.mpmExports) == True:
#            os.mkdir(self.mpmExports)
#            os.chmod(self.mpmExports, stat.S_IWGRP)
#            #msg = "Import all MPM Project Files In this folder:\n\n" + os.path.abspath(self.mpmExports)
#            #egui.msgbox(msg, title, ok_button, None, None)
#            #print("Check that Directory: " + os.path.abspath(self.mpmExports) + " exists or is mapped correctly.")
#        if not self.testDirectory(self.namerun) == True:
#            os.mkdir(self.namerun)
#            os.chmod(self.namerun, stat.S_IWGRP)
#        if not self.testDirectory(self.reports) == True:
#            os.mkdir(self.reports)
#            os.chmod(self.reports, stat.S_IWGRP)
#        if not self.testDirectory(self.userguide) == True:
#            os.mkdir(self.userguide)
#            os.chmod(self.userguide, stat.S_IWGRP)
        
    def testDirectory(self, directoryString):
        """test if the directory is valid"""
        
        title = "Server Mapping Check"
        ok_button ="OK"
        try:
            #check if it exists
            if os.path.isdir(directoryString) == True:
                return True
            else:
                return False
        except WindowsError:
            msg = 'map server: "\\rscakgh\ghshare" to the J Drive'
            egui.msgbox(msg, title, ok_button, None, None)
            #print('map server: "\\rscakgh\ghshare" to the J Drive')
            return False
        except IOError:
            msg = 'map server: "\\rscakgh\ghshare" to the J Drive'
            egui.msgbox(msg, title, ok_button, None, None)
            #print('map server: "\\rscakgh\ghshare" to the J Drive')
            return False
        except:
            return False

    def copyResources(self, sourceDir, destDir):
        """
        load/reload constant mapping files to local resources folder
        """
        
        if self.testDirectory(sourceDir) == True:    
            for each_file in os.listdir(sourceDir):
                try:
                    fullpath = os.path.join(sourceDir, each_file)
                    shutil.copy2(fullpath, destDir)
                except IOError:
                    
                    MsgDialog(None, msg = each_file + " is open.\nPlease close and reload", caption = "Close Resource File(s)", style = wx.OK).msg()
                    return False
    
    def copyResource(self, sourceFullPath, destDir):
        """
        load/reload one file to a destination folder
        """
        try:
            shutil.copy2(sourceFullPath, destDir)
        except IOError:
            MsgDialog(None, msg = sourceFullPath + " is open.\nPlease close and reload", caption = "Close File", style = wx.OK).msg()
        
    def convertAllMPMFiles(self):
        """
        goes to the MPM folder in resources and compiles all project files and creates 
        ETCs.csv in the resources folder
        """
        
        extensionList = self.getExtension(self.mpmExports) 
        
        #check if there are files in the folder        
        if len(extensionList) == 0:
            msg = "EMPTY DIRECTORY\nImport all MPM Project Files in this folder:\n\n" + os.path.abspath(self.mpmExports)
            egui.msgbox(msg, "Directory Check", "OK", None, None)
            #open folder
            os.startfile(self.mpmExports)
            return False
        else:
            self.convertAllXLS2CSV(self.mpmExports)
            return True
         
    def convertXLS2CSV(self, directoryString, fileName):
        ''' converts one excel file to csv; does not delete the excel file'''
        
        excel = win32.DispatchEx("Excel.Application")

        xlFilePath = self.createFilePath(directoryString, fileName)
        xlWorkBook = excel.Workbooks.Open(xlFilePath)
        
        fileNameAndExtension = os.path.splitext(fileName)
        createCSVFile = fileNameAndExtension[0] + ".csv"
        CSVFilePath = self.createFilePath(directoryString, createCSVFile)
        
        xlWorkBook.SaveAs(CSVFilePath, FileFormat = 24) # 24 represents xlCSVMSDOS        
        xlWorkBook.Close(False)
        
        excel.Quit()
        
    def convertAllXLS2CSV(self, directoryString): 
        '''converts a MS Excel file to csv w/ the same name in the same directory 
        and deletes the excel file'''
        
        try:
            excel = win32.DispatchEx('Excel.Application')
            excel.DisplayAlerts = False
            
            self.converted = []
            self.notConverted = []
            sheetcount = 0
            for each_file in os.listdir(directoryString):

                fileNameAndExtension = os.path.splitext(each_file)                
                
                #print each_file
                #print fileNameAndExtension[1].lower()
                
                xlFilePath = self.createFilePath(directoryString, each_file)
                if fileNameAndExtension[1].lower() != ".csv":  
                    
                    xlWorkBook = excel.Workbooks.Open(xlFilePath)
                    
                    sheetcount = xlWorkBook.Sheets.Count
                    if sheetcount == 1:
                        self.converted.append(xlFilePath)
                        
                        #fileDir, fileName = os.path.split(aFile)    
                        createCSVFile = fileNameAndExtension[0] + ".csv"
                        CSVFilePath = self.createFilePath(directoryString, createCSVFile)
                        
                        xlWorkBook.SaveAs(CSVFilePath, FileFormat =24) # 24 represents xlCSVMSDOS
                        xlWorkBook.Close(False)
                        self.deleteFile(directoryString, each_file)
                    else:
                        xlWorkBook.Close(False)
                        self.notConverted.append(each_file)
                else:
                    if xlFilePath not in self.converted:
                        self.converted.append(xlFilePath)
            
            excel.DisplayAlerts = True    
            excel.Quit()
                    
        except:
            print ">>>>>>> FAILED to convert " + each_file + " to CSV!"
    
    def convertCSVToXLSX(self, csv, xl):
        """converts to xlsx filetype; xl MUST end in xlsx!"""
        excel = win32.DispatchEx('Excel.Application')
        try:
            wb = excel.Workbooks.Open(csv)
        except:
            sys.exit()
        
        excel.DisplayAlerts = False
        
        while True:
            try:
                wb.SaveAs(xl, FileFormat = 51)
                break
            except:
                xl = self.revUp(xl)
                wb.SaveAs(xl, FileFormat = 51)
                
        address = wb.Path + "\\"  + wb.Name                
        wb.Close (False)
            
        excel.Visible = True
        excel.DisplayAlerts = True
        excel.Quit()
        
        return address
    
    def revUp(self, filestring):
        """ revs up a document """
        
        #split between name and extension
        splitxl = os.path.splitext(filestring)
        filename = splitxl[0]
        try:
            # check if the last two digits of the filename are integers
            rev = int(filename[-2:])
            newrev = rev + 1
            # if it's less than 10, then add a leading 0
            if len(str(newrev)) < 2:
                return filename[:-2] + "0" + str(newrev) + splitxl[1]
            else:
                return filename[:-2] + str(newrev) + splitxl[1]    
        # if value error, then it means that it's the original file and we want to go to rev 1
        except ValueError:
            filename = filename + "01"
            return filename + splitxl[1]
      
    def createFilePath(self, directoryString, fileName):
        ''' creates the file path '''
        return os.path.abspath(os.path.join(directoryString, fileName))
    
    def deleteFile(self, directoryString, fileName):
        ''' delete a file '''
        os.remove(self.createFilePath(directoryString, fileName))
        
    def checkNamerunDirectory(self):
        """checks if BW Namerun exists or if it's a csv file"""
        
        extensionList = self.getExtension(self.namerun)
        if len(extensionList) == 0:
            title = "Directory Check"
            msg = "EMPTY DIRECTORY\nSave Business Warehouse Namerun in this folder:\n(BW export must be .csv format)\n\n" + os.path.abspath(self.namerun)
            MsgDialog(None, msg, title, style = wx.OK).msg()
            #egui.msgbox(msg, "Directory Check", "OK", None, None)
            #open folder
            self.openNamerunDirectory()
        else:
            if len(extensionList) == 1 and extensionList[0] == "csv":
                return True
            else:
                egui.msgbox("Business Warehouse Namerun must be in csv format", "Check Namerun", "Ok", None, None)
    
    def openNamerunDirectory(self):
        """opens Namerun Folder"""
        os.startfile(self.namerun)
    
    def getNamerunFiles(self):
        '''checks for multiple namerun files and returns a list containing the full path of the files'''
        filenames = self.getFileNames(self.namerun)
        if len(filenames) == 1:
            return [self.createFilePath(self.namerun, filenames[0])]
        else:
            return FileDlg(self.namerun).select()

#            title = "Directory Check"
#            msg = "Multiple files found.\nContinue running using multiple nameruns?\n\tYES = Load ALL Files\n\tNO = Choose one or multiple Files"
#            
#            option = MsgDialog(None, msg, title, style = wx.YES_NO | wx.YES_DEFAULT).yesNoSelect()
#            
#            #option = egui.ynbox("Multiple files found.\Continue running using multiple nameruns?", "Directory Check", ("Yes", "No"), None)
#            if option == 0:
#                return FileDlg(self.namerun).select()
#            if option == 1:
#                filenamePaths = []
#                for each_filename in filenames:
#                    filenamePaths.append(self.createFilePath(self.namerun, each_filename))
#                return filenamePaths
            
    def getMPMfiles(self):
        '''returns a list containing the full path of the MPM files'''
        
        filenamePaths = []
        for each_filename in self.getFileNames(self.mpmExports):
            filenamePaths.append(self.createFilePath(self.mpmExports, each_filename))
        return filenamePaths
    
    def compileCSVFiles(self, sourceDir, destDir, fileName):
        """ compiles csv files from the source directory into one csv file stored
        in the destination directory with a filename from "fileName"
        assuming they all have the same header"""
        
        # join directory and file
        etcFilePath = self.createFilePath(destDir, fileName)
        # open/overwrite or create the file
        etcFile = csv.writer(open(etcFilePath, 'w+'))
        
        headerRow = None
#        collectData = []
        
        for each_file in os.listdir(sourceDir):
            mpmFilePath = self.createFilePath(sourceDir, each_file)
            mpmFile = open(mpmFilePath, 'r')
            
            readMPMFile = csv.reader(mpmFile, )
            headerRow = readMPMFile[0]
            
            for each_line in readMPMFile:
                etcFile.writerows(each_line)
    
    def getExtensionString(self, fileName):
        """ obtains extension of a file and returns it as a string"""
        
        fileExtRegEx = re.compile(r'[^.]*.(\w*)$', re.I)
        fileExt = fileExtRegEx.findall(fileName)
        
        return (fileExt[0].lower())
        
    def getExtension(self, directoryString):
        """ creates a list of extensions that's found in the directory """
        
        extensions = []
        
        for root, dirs, files, in os.walk(directoryString):
            for each_file in files:
                
                fileExt = self.getExtensionString(each_file)
                
                if not fileExt in extensions:
                    extensions.append(fileExt)
                    
        return extensions

    def getFileNames(self, directoryString):
        """ returns a list of the filenames in the directory """
        filenames = []
        
        for each_file in os.listdir(directoryString):
            filenames.append(each_file)

        return filenames
    
    def checkNetworkStatus(self):
        """
        checks the status of the network we currently have on file and checks the POP in order to update
        whether it is closed or open
        """
        pass
    
    def checkForNewNetworks(self):
        """
        checks for new Networks in ETCs.csv that's not in networks.csv 
        """
        pass

    def editEmployee(self, empID, firstName = None, lastName = None, camCode = None, functionalHomeRoom = None):
        """
        via EmpGUI, checks employee's data and edits the new information
        """
        
        ID = None
        firstName = None
        lastName = None
        empID  = None
        camCode  = None
        functionalHomeRoom  = None        
        
        pass
        
    def reloadEmployee(self, empFilename, empList):
        ''' 1. create a tempfile in mapping folder
            2. go through empList* and add to tempfile
            3. replace original file
            4. replaces the copy in the local folder 
            5. delete tempfile
            
            * empList is a list of the class instance Employee and therefore has its attributes
            '''
        
        #set filepaths
        tempFilename = empFilename + ".tmp"
        tempPath = self.createFilePath(self.mappingResources, tempFilename)
        resourceFilePath = self.createFilePath(self.resourcesDirectory, empFilename)
        sourceFilePath = self.createFilePath(self.mappingResources, empFilename)

        #CREATE TEMPFILE
        tempfile  = open(tempPath, "w")
        #WRITE HEADERS and DATA
        tempfile.write("//First Name,Last Name,EmpNo,CAM,FHR\n")
        for emp in empList:
            tempfile.write(emp.EmpFirstName + "," + emp.EmpLastName + "," + emp.EmpNo + "," + emp.CAMCode + "," + emp.FuncHomeRoom + "\n")
        
        tempfile.close()
        
        #delete old employee file, then rename tempfile to the old filename
        os.remove(sourceFilePath)
        self.renameFile(tempPath, sourceFilePath)
        MsgDialog(None, "Reset and Start from Step 1 to reflect changes", "Message", style = wx.OK).msg()
    
    def renameFile(self, origFile, newFile):
        ''' origFile = must be full path (directory + filename)
            newFile = must be full path (directory + filename)'''
        
        os.rename(origFile, newFile)
    
        
#open(os.path.join(mappingDirectory, 'CAMS.csv'), 'w')