'''
Created on Sep 11, 2012

@author: g73666
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
            
recursive_remove("dist")
recursive_remove("build")

cur = os.getcwd()
filename = "Driver.py"
dir = os.path.abspath(os.path.join(cur, filename))

setup(windows=[dir])