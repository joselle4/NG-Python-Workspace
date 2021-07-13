'''
Created on Aug 20, 2012

@author: Joselle Abagat
'''

from distutils.core import setup
import win32com.client as win32
import py2exe
import os

def recursive_remove(dir_name):
    for root, dirs, files in os.walk(dir_name, topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))
    #os.remove(dir_name)

#################################################################################

recursive_remove("dist")
recursive_remove("build")

cur = os.getcwd()
dir = os.path.abspath(os.path.join(cur, "Driver.py"))

msoExcel12 = ('{00020813-0000-0000-C000-000000000046}', 0, 1, 6)
msoActiveX = ('{2A75196C-D9EB-4129-B803-931327F72D5C}', 0, 2, 8)
msOffice12 = ('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 4)

setup(windows=[dir], options = {"py2exe": {"typelibs": 
                                 [msoExcel12,
                                  msoActiveX,
                                  msOffice12
                                  ]
                                 }
                      }
      )

# DEBUGGING MODE
#setup(console=[dir], options = {"py2exe": {"typelibs": 
#                                 [('{00020813-0000-0000-C000-000000000046}', 0, 1, 6),
#                                  ('{2A75196C-D9EB-4129-B803-931327F72D5C}', 0, 2, 8)
#                                  ]
#                                 }
#                      }
#      )

#setup(console=[dir])

#setup(console=[dir], options={'py2exe':{'skip_archive':True}})

#CMD: curdir>python filename.py py2exe
#CMD on win32com\client>python makepy.py -i