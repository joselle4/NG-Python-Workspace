'''
Created on Jul 26, 2012

@author: Joselle Abagat
'''

import os, shutil, re, csv, sys
import win32com.client as win32
import win32api
import itertools
import pythoncom
import tempfile
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


class Excel(object):
    '''
    classdocs
    '''

    def __init__(self):
               
        win32.gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, 6) #Microsoft Excel 12.0 Object Library
        win32.gencache.EnsureModule('{2A75196C-D9EB-4129-B803-931327F72D5C}', 0, 2, 8) #Microsoft Active X Data Objects 2.8 Library
        win32.gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 4) #Microsoft Office 12.0 Object Library
        
        self.xl = win32.DispatchEx('Excel.Application')
        self.win32c = win32.constants
        
#        print "constants:\n"
#        print self.win32c.__dicts__
                                    
    def addExcelWorkbook(self, fullpath):
        """add excel workbook given name"""
        wb = self.xl.Workbooks.Add()
        self.saveWorkbook(wb, fullpath)
    
    def addWorksheet(self, wb, name):    
        """ adds and returns worksheet given name"""
        ws = wb.Worksheets.Add()
        ws.Name = name
        return ws
    
    def saveWorkbook(self, wb, fullpath):
        """saves workbook and moves up revision if it already exists"""
        while True:
            try:
                wb.SaveAs(fullpath, self.win32c.xlNormal,"","",False)
                wb.Saved = True
            except:
                revUpFile = self.revUp(fullpath)
                wb.SaveAs(revUpFile, self.win32c.xlNormal,"","",False)
                wb.Saved = True
                
    def setWorkbook(self, fullpath): 
        #assuming it's already open, we only want the filename, not the filepath
        self.xl.Workbooks(fullpath)
        pass
    
    def setWorksheet(self, wb, wsName): return wb.Worksheets(wsName) #SIMILAR TO VBA POINTING TO OBJECT
    
    def setRange(self, wb, wsName, cellStart=(), cellEnd=()): 
        """ similar to how VBA points to object"""
        #print wb.Name
        #print wsName
        ws = self.setWorksheet(wb, wsName)
        firstRange = ws.Cells(cellStart[0],cellStart[1])
        lastRange = ws.Cells(cellEnd[0], cellEnd[1])
        return ws.Range(firstRange, lastRange)

#    def setRange(self, wb, wsName, cellStart=(), cellEnd=()): 
#        """ similar to how VBA points to object"""
#        print wb.Name
#        print wsName
#        ws = self.setWorksheet(wb, wsName)
#        firstRange = ws.Cells(cellStart[0],cellStart[1])
#        lastRange = ws.Cells(cellEnd[0], cellEnd[1])
#        return ws.Range(firstRange, lastRange)
    
    def rangeAddress(self, range):
        """returns address of range"""
        return range.Address
    
    def getRows(self): pass
    
    def getColumns(self): pass
    
    def falseAlerts(self): self.xl.DisplayAlerts = False
    
    def trueAlerts(self): self.xl.DisplayAlerts = True
            
    def quitExcel(self): self.xl.Application.Quit()
    
    def openExcelFile(self, fullpath): return self.xl.Workbooks.Open(fullpath)
    
    def visible(self): self.xl.Visible = True
    
    def addPivot(self, wbk, sourcedata, pivotName, filters = (), filterstring = "", columns = (), rows = (), subtotalTF = (), sumvalue = (), sortfield = ""):
        """ assumes workbook is already open """
        
#        self.__init__()
#        print self.win32c.__dicts__
        
#        win32.gencache.Rebuild()
        
        psheet = self.addWorksheet(wbk, pivotName)
        destination = psheet.Cells(3,1)
        
#        win32c = win32.constants
        tablecount = itertools.count(1)
        pivotName = "pivotName%d"%tablecount.next()
        pAdd = wbk.PivotCaches().Add(SourceType = self.win32c.xlDatabase, SourceData = sourcedata)
        
        pCreate = pAdd.CreatePivotTable(TableDestination=destination, TableName=pivotName, ReadData=pythoncom.Missing, DefaultVersion=self.win32c.xlPivotTableVersion10) #@UndefinedVariable
        
        #NO NEED TO USE DUE TO CLEANER GROUPED FOR LOOP BELOW    
#        for i,each in enumerate(rows):
#            print each
#            pCreate.PivotFields(each).Orientation = win32c.xlRowField
#            pCreate.PivotFields(each).Position = i + 1
#        
#        for i,each in enumerate(columns):
#            print each
#            print i+1
#            pCreate.PivotFields(each).Orientation = win32c.xlColumnField
#            pCreate.PivotFields(each).Position = i + 1
#        
#        for i,each in enumerate(filters):
#            pCreate.PivotFields(each).Orientation = win32c.xlPageField
#            pCreate.PivotFields(each).Position = i + 1
            
        # in this for loop: fieldlist will be the filters, columns, and rows
        for fieldlist, fieldc in ( (filters, self.win32c.xlPageField),
                                   (columns, self.win32c.xlColumnField),
                                   (rows, self.win32c.xlRowField)
                                   ):
            for i,each in enumerate(fieldlist):
                pCreate.PivotFields(each).Orientation = fieldc
                pCreate.PivotFields(each).Position = i + 1
        
        psheet.PivotTables(pivotName).AddDataField(
            pCreate.PivotFields(sumvalue[7:]), sumvalue, self.win32c.xlSum)
                
        return [psheet,pCreate]
        
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
    
    def printError(self, errorNumber): win32api.FormatMessage(errorNumber) #PRINT ERROR NUMBER
        
    def colorRange(self, range, colorIndex): range.Interior.ColorIndex = colorIndex #CHANGE CELL COLOR
    
    def addComment(self, ws, cellAddress, textComment, height, width, show = True):
        ''' add comment to a selected cell'''
        
        ws.Activate
        ws.Range(cellAddress).AddComment()
        ws.Range(cellAddress).Comment.Text(textComment)
        ws.Range(cellAddress).Comment.Visible = show
        ws.Range(cellAddress).Comment.Shape.Height = height
        ws.Range(cellAddress).Comment.Shape.Width = width
        
    def freezePane(self, ws, cellAddress):
        ''' freeze pane'''
        ws.Range(cellAddress).Select
        self.xl.ActiveWindow.FreezePanes = True
    
    def pivotSubTotal(self, pTable, pField, OnOff = False): 
        """removes subtotals"""
        tf = (OnOff, OnOff, OnOff, OnOff, OnOff, OnOff, OnOff, OnOff, OnOff, OnOff, OnOff, OnOff)
        pTable.PivotFields(pField).Subtotals = tf
        
    def pivotFilter(self, pTable, pField, filterString): pTable.PivotFields(pField).CurrentPage = filterString #CREATES FILTER FOR PIVOT
    
    def pivotItemFilter(self, pTable, pField, allValues, filterValues):
        """ filters items in the pivot table values """
        falseValues = []
        for eachItem in allValues:
            if eachItem not in filterValues:
                falseValues.append(eachItem)
        
        falseValues = list(set(falseValues)) #remove dupliCates
        for each in falseValues:              
            # this is for activity codes only
            try:
                int(each)
                if each[:2] == "00":
                    each = each[2:]
            except:
                pass
            pTable.PivotFields(pField).PivotItems(each).Visible = False

    def pivotItemFilterReverse(self, pTable, pField, allValues, filterValues):
        """ filters items in the pivot table values in reverse"""
        for each in filterValues:              
            # this is for activity codes only
            try:
                int(each)
                if each[:2] == "00":
                    each = each[2:]
            except:
                pass
            pTable.PivotFields(pField).PivotItems(each).Visible = False
    
    def pivotTheme(self, pTable, styleString): pTable.TableStyle2 = styleString #ASSIGN STYLE
    
    def conditionUsingFormula(self, ws, range, formula, conditionNumber):
        """ 
            http://www.ablebits.com/office-addins-blog/2011/05/23/excel-conditional-formatting-pivottables/
            
            works by adding a new conditional format provided the formula, then converts color choice and assigns cell color
        """
        
        ws.Range(range).FormatConditions.Add(Type=self.win32c.xlExpression, Formula1=formula)
        color = self.convertRGB(self.dict_colorChoice()["beige"])
        ws.Range(range).FormatConditions(conditionNumber).Interior.Color = color
        
    def conditionUsingString(self, ws, range, string, interiorColor, fontColor, conditionNumber):
        """
            adds a new conditional format and edits the cell and font color
            need to use colorindex command for font
        """

        ws.Range(range).FormatConditions.Add(Type=self.win32c.xlTextString, String=string, TextOperator=self.win32c.xlContains)
        ws.Range(range).FormatConditions(conditionNumber).Interior.Color = interiorColor
        ws.Range(range).FormatConditions(conditionNumber).Font.ColorIndex = fontColor

    def conditionUsingCellValue(self, ws, range, cellValue, xlOperator, colorFormat):
        #  = "xlGreater"
        stringFormula = '"=' + str(cellValue) + '"'
        ws.Range(range).FormatConditions.Add(Type=self.win32c.xlCellValue, Operator=xlOperator, Formula=stringFormula)
        conditionNumber = ws.Range(range).FormatConditions.Count
        ws.Range(range).FormatConditions(conditionNumber).SetFirstPriority
        
        ws.Range(range).FormatConditions(1).Font.TintAndShade = colorFormat
    
    def conditionPivotScopeUsingCellValue(self, ws, range, min, max, conditionNumber):
        
        formula1 = '"=' + str(min) + '"'
        formula2 = '"=' + str(max) + '"'
        color = self.convertRGB(self.dict_colorChoice()["beige"])
#        ws.Range(range).FormatConditions.Add(Type=self.win32c.xlCellValue, Operator=self.win32c.xlBetween, Formula1=formula1, Formula2=formula2)
        ws.Range(range).FormatConditions.Add(Type=self.win32c.xlCellValue, Operator=self.win32c.xlBetween, Formula1=min, Formula2=max)
#        conditionNumber = ws.Range(range).FormatConditions.Count
#        ws.Range(range).FormatConditions(conditionNumber).SetFirstPriority
        
        ws.Range(range).FormatConditions(conditionNumber).Interior.Color = color
        ws.Range(range).FormatConditions(conditionNumber).ScopeType = self.win32c.xlFieldsScope

    def conditionPivotScopeUsingFormula(self, ws, range, formula, conditionNumber):
        """
            allows conditional formatting of data field by using field scope option; 
            in this way, format will stay even if you refresh the pivot
        """
        color = self.convertRGB(self.dict_colorChoice()["beige"])
        ws.Range(range).FormatConditions.Add(Type=self.win32c.xlExpression, Formula1=formula)
        conditionNumber = ws.Range(range).FormatConditions.Count
        ws.Range(range).FormatConditions(conditionNumber).Interior.Color = color
        ws.Range(range).FormatConditions(conditionNumber).ScopeType = self.win32c.xlDataFieldScope        

    
    def conditionPivotSelect(self, pivotTable, pivotRowName, styleString):
        """ vba command: pt.PivotSelect Name:="", Mode:=xlLabelOnly, UseStandardName:=True """
        
        pivotTable.PivotSelect(Name=pivotRowName, Mode=self.win32c.xlLabelOnly, UseStandardName=True)
        self.xl.Selection.Style = styleString
    
    def dict_PivotStyles(self):
        return {"grey":"PivotStyleMedium1",
                "blue":"PivotStyleMedium2",
                "red":"PivotStyleMedium3",
                "green":"PivotStyleMedium4",
                "purple":"PivotStyleMedium5",
                "aqua":"PivotStyleMedium6",
                "orange":"PivotStyleMedium7",
                "None":"None"}
    
    def dict_colorChoice(self):
        """
            name        R    G    B
            grey      216    216    216
            beige     221    217    195
            blue      197    217    241
            red       242    221    220
            green d   215    228    188
            green l   234    241    221
            green b   146    208    80
            purple    204    192    218
            aqua      182    221    232
            orange    255    192    0
        """
        return {"grey":(216,216,216),
                "beige":(221,217,195),
                "blue":(197,217,241),
                "red":(242,221,220),
                "green":(234,241,221),
                "purple":(204,192,218),
                "aqua":(182,221,232),
                "orange":(255,192,0)}
    
        
    
    def convertRGB(self, (R,G,B)):
        """ rgb = (#, #, #); converts value to hexadecimal """
        return (256 * 256 * B)   + G * 256 + R 

    def listColors(self): 
        """ lists colors """
        colors = self.dict_PivotStyles().keys()
        colors.sort()
        return colors
    
    
#    ActiveSheet.PivotTables("pivotName1").PivotSelect "'Percent Spent'[All]", _
#        xlLabelOnly, True
#    Selection.Style = "Percent"
#    Range("E7").Select
#    ActiveSheet.PivotTables("pivotName1").PivotCache.Refresh