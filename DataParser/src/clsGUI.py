'''
Created on May 22, 2012

@author: Joselle Abagat
'''

import easygui, wx, sys, os, re, win32com.client, csv, shutil

wildcard = "CSV files (*.csv)|*.csv|" \
           "Excel files (*.xls, *.xlsx, *.xlsm)|*.xls|*.xlsx|*.xlsm|" \
           "Text files (*.txt)|*.txt" \
           "All files (*.*)|*.*"

class App(wx.App):
    def __init__(self, redirect = False, filename = None, useBestVisual = True, clearSigInt = True):
        wx.App.__init__(self, redirect, filename, useBestVisual, clearSigInt)

class FileDialog(wx.FileDialog):
    
    def __init__(self):
        self.dlg = wx.FileDialog.__init__(self,wx.Panel,
                                 message="Choose a file",
                                 defaultDir=os.getcwd(),
                                 defaultFile="",
                                 wildcard=wildcard,
                                 style=wx.OPEN | wx.MULTIPLE | wx.CHANGE_DIR
                                 )
    
        if self.dlg.ShowModal() == wx.ID_OK:
            paths = self.dlg.GetPaths() #an array            
            
            for path in paths:
                print path
            
        self.dlg.Destroy()

class ComboBox(wx.Panel):
    ''' this is a dynamic combo box panel that allows for a variable size, entries, and combo boxes 
        PARAMETERS:
            parent
            static text
            list for combo box
            increment for distance between combo boxes
            handler ==> WE WANT THE EVENT HANDLER TO BE CONTROLLED FROM
                        THE DRIVER MODULE SO THIS CLASS CAN REMAIN INDEPENDENT.
                        WE CAN'T JUST USE A SIMPLE RETURN FUNCTION HERE BECAUSE
                        OF THE POSSIBILITY OF MULTIPLE VARIABLES
            '''
    
    def __init__(self, parent = None):
        
        self.id = wx.ID_ANY
        wx.Panel.__init__(self, parent, self.id)
    
    def comboBox(self, selection1 = [], defaultValue1 = "", handler1 = None, position = (0,0)):
        #1st combo box
        cb1 = wx.ComboBox(self, self.id, defaultValue1, position, choices = selection1, style = wx.CB_DROPDOWN) #| wx.TE_PROCESS_ENTER | wx.TE_PROCESS_TAB)
        self.Bind(wx.EVT_TEXT, handler1, cb1, self.id)
            
    def combo2Boxes(self, selection1 = [], defaultValue1 = "", handler1 = None, selection2 = [], defaultValue2 = "", handler2 = None, increment = 0.0):
        #1st combo box
        position = (20, increment)
        cb1 = wx.ComboBox(self, self.id, defaultValue1, position, choices = selection1, style = wx.CB_DROPDOWN) #| wx.TE_PROCESS_ENTER | wx.TE_PROCESS_TAB)
        self.Bind(wx.EVT_TEXT, handler1, cb1, self.id)
        
        #2nd combo box
        position = (100, increment)
        cb2 = wx.ComboBox(self, self.id, defaultValue2, position, choices = selection2, style = wx.CB_DROPDOWN) #| wx.TE_PROCESS_ENTER | wx.TE_PROCESS_TAB)
        self.Bind(wx.EVT_TEXT, handler1, cb2, self.id)        
    
    def comboTextAndBox(self, text = "", selection = [], defaultValue = "", increment = 0.0, handler = None):
        
        position = (100, increment)
        self.staticText = wx.StaticText(self, self.id, text, position, style = wx.ALIGN_RIGHT)
        self.staticText.SetForegroundColour(wx.WHITE)
        self.staticText.SetFont(self.boldFont())
        
        position = (20, increment)
        cb = wx.ComboBox(self, self.id, defaultValue, position, choices = selection, style = wx.CB_DROPDOWN) #| wx.TE_PROCESS_ENTER | wx.TE_PROCESS_TAB)
        self.Bind(wx.EVT_TEXT, handler, cb, self.id)
        self.Bind(wx.EVT_TEXT, handler, self.staticText, self.id)

    def boldFont(self): return wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD)
    
class MultiChoiceDialog(wx.MultiChoiceDialog):
    
    def __init__(self, parent = None, msg = "", title = "", listArg = []):
        """ constructs multichoice dialog box with parameters: parent, msg, title, and list of choices """
        wx.MultiChoiceDialog.__init__(self, parent, msg, title, listArg)
    
    def select(self, listArg = []):
        ''' returns a list of choices or exits program'''
        if self.ShowModal() == wx.ID_OK:
            selections = self.GetSelections()
            strings = [listArg[x] for x in selections]
            self.Destroy()
            return strings
        else:
            self.Destroy()
#            sys.exit(0)

class MultiChoiceDialogSub(wx.MultiChoiceDialog):
    """ same functionality as the first multichoicedialog but does not exit program upon quitting """
    def __init__(self, parent = None, msg = "", title = "", listArg = []):
        """ constructs multichoice dialog box with parameters: parent, msg, title, and list of choices """
        wx.MultiChoiceDialog.__init__(self, parent, msg, title, listArg)
    
    def select(self, listArg = []):
        ''' returns a list of choices or exits program'''
        if self.ShowModal() == wx.ID_OK:
            selections = self.GetSelections()
            strings = [listArg[x] for x in selections]
            self.Destroy()
            return strings
        else:
            self.Destroy()
#            sys.exit(0)

class MsgDialog(wx.MessageDialog):
    def __init__(self, parent = None, msg = "", caption = "", style = wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL):
        wx.MessageDialog.__init__(self, parent, msg, caption, style)
        self.SetTitle(caption)
    
    def yesNoSelect(self):
        result = self.ShowModal()
        if result == wx.ID_YES:
            self.Destroy()
            return 1
        elif result == wx.ID_NO:
            self.Destroy()
            return 0
        else:
            self.Destroy()
            sys.exit(0)
    
    def msg(self):
        self.ShowModal()
        self.Destroy()

class TextFrame(wx.Frame):
    def __init__(self, parent = None, title = "", msg = "", content = "", handler = None):
        wx.Frame.__init__(self, parent, wx.ID_ANY, title)
        
        self.SetTitle(title)
        id = wx.ID_ANY
        panel = wx.Panel(self, id)
        multiLabel = wx.StaticText(panel, id, msg, pos = wx.Point(0,0))
        multiLabel.SetFont(self.boldFont())
        multiText = wx.TextCtrl(panel, id, content, pos = wx.Point(0,250), size = (380, 100), style = wx.TE_MULTILINE|wx.TE_PROCESS_ENTER)
        multiText.SetInsertionPoint(0)
        
        button = wx.Button(panel, 10, "Continue", pos = wx.Point(150,370))
        #self.Bind(wx.EVT_BUTTON, self.onContinue, button)
        self.Bind(wx.EVT_BUTTON, handler, button)
        button.SetDefault()
        button.SetSize(button.GetBestSize())
        
        #sizer = wx.FlexGridSizer(cols = 1, rows = 10, hgap = 6, vgap = 6)
        #sizer.AddMany([(multiLabel, wx.EXPAND | wx.ALIGN_CENTER), 
        #               (multiText, wx.EXPAND | wx.ALIGN_CENTER), 
        #               (button, wx.EXPAND | wx.ALIGN_CENTER),
        #               ])
        
        #panel.SetSizer(wx.BoxSizer(wx.VERTICAL | wx.ALL))
        #panel.SetSizer(sizer)
        #panel.SetAutoLayout(True)
        #panel.Center()
        
        self.SetSizeWH(400, 430)
        
        self.Show()   

    def boldFont(self): return wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD)
    
    def onContinue(self, event): self.Close()
    
    def onClose(self, event):
        self.Destroy()
        sys.exit(0)
        
#class ModifyData(wx.Frame):
#    
#    def __init__(self, parent = None, title = "", *handlersList, **namedListsDict):
#        pass