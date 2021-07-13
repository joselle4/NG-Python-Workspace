'''
Created on Mar 30, 2012

@author: joselle4
'''
import re
import os

class Extensions(object):
    '''
    classdocs
    '''

    def __init__(self, directory_string):
        '''
        Constructor
        '''
        self.list_extensions = self.mapFileSizesToExtension(directory_string)

    def mapFileSizesToExtension(self, directory_string):

        extension_dictionary = {}
    
        for root, dirs, files in os.walk(directory_string):
            for eachfile in files:
                
                #need to account for filenames like: "blah 2.12.2012.txt"
                getExtension = re.compile(r'[^.]*.(\w*)$',re.I)
                extension = getExtension.findall(eachfile)
                extension = extension[0]
                
                try:
                    file_size = int(os.path.getsize(os.path.join(root, eachfile)))
                except WindowsError:
                    file_size = 0
                    
                if not extension in extension_dictionary:
                    extension_dictionary[extension] = [file_size]
                else:
                    extension_dictionary[extension].append(file_size)
        
        return extension_dictionary
    
    def getNumberOfFiles(self, file_extension): 
        try:
            return len(self.list_extensions[file_extension])
        except KeyError:
            print("The extension is not a key in the dictionary")
    
    def getMinFileSize(self, file_extension):
        try:
            return min(self.list_extensions[file_extension])
        except TypeError:
            pass
        except KeyError:
            print("The extension is not a key in the dictionary")
    
    def getMaxFileSize(self, file_extension):
        try:
            return max(self.list_extensions[file_extension])
        except TypeError:
            pass
        except KeyError:
            print("The extension is not a key in the dictionary")
            
    def getAveFileSize(self, file_extension):
        try:
            return sum(self.list_extensions[file_extension])/self.getNumberOfFiles(file_extension)
        except TypeError:
            pass
        except KeyError:
            print("The extension is not a key in the dictionary")
        except ZeroDivisionError:
            print("There are 0 files with that extension")
    
    def RunReport(self):
        for each_key in self.list_extensions:
            s = "Extension: '" + str(each_key) + "'" + \
                "\n\tNumber of Files: " + str(self.getNumberOfFiles(each_key)) + \
                "\n\tMax file size: " + str(self.getMaxFileSize(each_key)) + " bytes" + \
                "\n\tMin file size: " + str(self.getMinFileSize(each_key)) + " bytes" + \
                "\n\tAverage file size: " + str(self.getAveFileSize(each_key)) + " bytes"
            print(s)
        return s

filename = "..\\..\\..\\"
print(os.path.abspath(filename))
ext = Extensions(filename)
ext.RunReport()
