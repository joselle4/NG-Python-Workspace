from collections import *

class Odict(UserDict):
        
#    def __init__(self, dict=None, order=None, **kwargs):
#        self.data = {}
#        if dict is not None:
#            self.update(dict)
#        if len(kwargs):
#            self.update(kwargs)
#        
#        self.order = {}
#        if order is not None:
#            self.update(order)
    
    def __setitem__(self, key, item): 
        keyValueOrder = "entry " + str(len(self.data) + 1)
        #keyValueOrder = len(self.data) + 1
        #self.data[key] = item               
        self.data[(keyValueOrder, key)] = item
        #self.data[keyValueOrder] = (key, item)

    
    def okeys(self):
        
        keyOrder = []
        sortedKeys = []
        
        for each_key in self.data.keys():
            try:
                keyOrder.append(each_key[0])
                #keyOrder.append(self.data[each_key])
            except KeyError:
                print("No key found")
        
        keyOrder.sort()
        
        for each_item in keyOrder:
            for each_key in self.data.keys():
                if each_key[0] == each_item:
                    sortedKeys.append(each_key[1])
        
        return sortedKeys
    

dict1 = Odict({1:"one", 2:"two", 10:"ten"})
print(dict1)
dict1[7]="seven"
print(dict1)
dict1[4] = "four"
print(dict1)
print(dict1.okeys())

#    def __init__(self, listarg):
#        self.userdict_dictionary = UserDict(listarg)
#        self.our_dictionary = self.createdict(listarg)
#
#    def createdict(self, listarg):
#        dict = {}
#        for each in listarg:
#            if not each in dict:
#                dict[each[0]] = each[1]
#        return dict
#    
#    def __setitem__(self, key, value):
#        
#        order = {}
#        if not key in order:
#            order[key] = value
#        else:
#            order[key].append(value)
#            
#        return order
#
#    def okeys(self):
#
#        pass
#
#    def __str__(self):
#        return str(self)
#
#
#filepath = "C:\Documents and Settings\G73080\My Documents\programming\python\data.csv"
#
#def datalist(filepath):
#    readfile = open(filepath, 'r')
#    listdata = []
#    for each in readfile:
#        item = each.split('\t')
#        if not item in listdata:
#            listdata.append(item)
#    return listdata
#
#data = datalist(filepath)
#dict1 = UserDict(data)
#print(data)
#data.append(["HXX","YSSELS"])
#odict = Odict(data)
#print(odict)
