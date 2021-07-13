'''
Created on Apr 3, 2012

@author: Joselle Abagat
'''

from collections import *

class Odict(UserDict):
    
    def __setitem__(self, key, item): 
        keyValueOrder = "Entry " + str(len(self.data) + 1)
        self.data[(keyValueOrder, key)] = item
    
    def okeys(self):
        
        keyOrder = []
        sortedKeys = []
        
        for each_key in self.data.keys():
            try:
                keyOrder.append(each_key[0])
            except KeyError:
                print("No key found")
        
        keyOrder.sort()
        
        for each_item in keyOrder:
            for each_key in self.data.keys():
                if each_key[0] == each_item:
                    sortedKeys.append(each_key[1])
        
        return sortedKeys
    
#sample run
dict1 = Odict({1:"one", 2:"two", 10:"ten"})
print(dict1)
dict1[7]="seven"
print(dict1)
dict1[4] = "four"
print(dict1)
print(dict1.okeys())