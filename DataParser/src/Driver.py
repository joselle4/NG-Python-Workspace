'''
Created on Sep 11, 2012

@author: Joselle Abagat
'''

from clsGUI import *
from clsDataParse import *



class DriverApp(wx.App):
    """ need wx.app in order to create frames and subframes"""
    
    def __init__(self, redirect = False, filename = None, useBestVisual = True, clearSigInt = True):
        """ overwrite wx.app instance """
        wx.App.__init__(self, redirect, filename, useBestVisual, clearSigInt)
        
    def OnInit(self):
        """ create on initialize method to run the data loading main frame when class instance is called"""
        drive()
        return True

def drive():
    fdlg = FileDialog()

if __name__ == "__main__":
    DriverApp(False).MainLoop()
#    fldg = FileDialog()