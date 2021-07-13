'''
Created on Mar 15, 2012
Basic GUI to interact with classes in this project.
@author: Jesus Medrano
'''
import wx
import os


#############
# GUI portion
#############
   
class StartFrame(wx.Frame):
    
    def __init__(self, parent, id):
        wx.Frame.__init__(self, parent, id, "Global Hawk Time Sheet Analyzer", (-1, -1))
        
        # Create Status Bar
        status=self.CreateStatusBar()
        self.SetStatusText("Database: %s" % "Some stuff")
        
        # Create the Menu Bar
        menubar = wx.MenuBar()
        
        # Create the menu items
        file_menu = wx.Menu()
        edit_menu = wx.Menu()
        view_menu = wx.Menu()
        report_menu = wx.Menu()
        
        # Create sub menus
        open_submenu = wx.Menu()
        
        
        # Create IDs for each of the menu items
        self.ID_FILE_LOAD_MPEX = wx.NewId()
        self.ID_FILE_OPEN_DB = wx.NewId()
        self.ID_FILE_BACKUP = wx.NewId()
        self.ID_SEARCH = wx.NewId()
        self.ID_ASSEMBLY_VIEW = wx.NewId()
        
        
        # Add Menu item to the File Menu
        open_submenu.Append(self.ID_FILE_LOAD_MPEX,"Load from &Name Run", "Load an Name Run export.")
        open_submenu.Append(self.ID_FILE_OPEN_DB,"Open &Database", "Opens an existing Mass Properties Database")
        file_menu.AppendSubMenu(open_submenu, '&Open')
        file_menu.Append(self.ID_FILE_BACKUP,"&Backup", "Backup current Database")
        file_menu.Enable(self.ID_FILE_BACKUP, False)
        file_menu.AppendSeparator()
        file_menu.Append(wx.ID_EXIT,"E&xit", "Exits the program")
        
        # Add Menu item to the Edit Menu
        
        
        # Add Menu item to the View Menu
        view_menu.Append(self.ID_SEARCH, "&Search")
        
        
        # Events
        wx.EVT_MENU(self, self.ID_FILE_LOAD_MPEX, self.load_MPEX)
        wx.EVT_MENU(self, self.ID_FILE_OPEN_DB, self.ToDo)
        wx.EVT_MENU(self, wx.ID_EXIT, self.exitProgram)
        
        # Add to Menu Bar
        menubar.Append(file_menu, "&File")
        menubar.Append(edit_menu, "&Edit")
        menubar.Append(view_menu, "&View")
        menubar.Append(report_menu, "&Reports")
        
        self.SetMenuBar(menubar)
        
    
    def load_MPEX(self, evt):
        dlg = wx.FileDialog(self, "Choose a file", os.getcwd(), "", "*.*", wx.OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            mypath = os.path.basename(path)

            try:
                             
                self.SetStatusText("Database: %s" % path)
                
                # If successfully loaded 
                dlg = wx.MessageDialog(self, 'Successfully loaded %s' % path, 'MPEX Successfully Loaded',
                             wx.OK | wx.ICON_INFORMATION)
                dlg.ShowModal()
                dlg.Destroy()
            except:
                print "Could not load " + path
                
            finally:
                dlg.Destroy()          
        
            
    def ToDo(self, evt):
        """
        A general purpose "we'll do it later" dialog box
        """
        dlg = wx.MessageDialog(self, 'Not Yet Implemented!', 'ToDo',
                             wx.OK | wx.ICON_INFORMATION)
        dlg.ShowModal()
        dlg.Destroy()
        
    def exitProgram(self, evt):
        """
        Exits the program
        """
        self.Close(True)


def main():
    """Main function to be ran on execution of this file"""
    #_app = SplashApp(redirect=True, filename="program.log") #Use for release Non Debug
    _app = SplashApp() # Debug version0
    _app.MainLoop()    
    

###
# Splash Screen
###



class SplashScreen(wx.SplashScreen):
    """
    Creates a splash screen widget.
    """
    
    def __init__(self, parent=None):
        _bitmap= wx.Image(name = '../../../../resources/global_hawk_qr_code.png').ConvertToBitmap()
        _splashStyle = wx.SPLASH_CENTRE_ON_SCREEN | wx.SPLASH_TIMEOUT
        _splashDuration = 2000 # milliseconds
        
        wx.SplashScreen.__init__(self, _bitmap, _splashStyle, _splashDuration, parent)
        self.Bind(wx.EVT_CLOSE, self.OnExit)
        
        # Stuff to do during splash screen
        
        # Load the local database
        
        
        wx.Yield()

    def OnExit(self, evt):
        self.Hide()
        
        # StartFrame is the main frame.
        _startFrame = StartFrame(parent=None, id=-1)
        _startFrame.Show()
        #self.SetTopWindow(_startFrame)
        
        
        # The program will freeze without this line
        evt.Skip()
        
class SplashApp(wx.App):
    def OnInit(self):
        _splash = SplashScreen()
        _splash.Show()
        
        return True


if __name__ == '__main__':
    main()