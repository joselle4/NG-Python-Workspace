'''
Created on Apr 3, 2012

@author: Joselle Abagat
'''

from collections import *

class Ulist(UserList):

    def __add__(self, other):
        if isinstance(other, UserList):
            return self.__class__(self.data + other.data)
        elif isinstance(other, type(self.data)):
            return self.__class__(self.data + other)
        else:
            return self.__class__(self.data + list(other))

    def append(self, item):
        if not item in self.data:
            self.data.append(item)

    def extend(self, other):
        if isinstance(other, UserList):
            for item in other.data:
                if not item in self.data:
                    self.data.extend(item)
        else:
            for item in other:
                if not item in self.data:
                    self.data.extend(item)

mylist = Ulist(["this", "is","a", "python"])
print(mylist)
mylist.append("class")
print(mylist)
mylist.append("class")
print(mylist)
mylist.append(["python", "for", "new", "learners"])
print(mylist)
#strings are iterable
mylist.extend(["a", "fun", "python", "class"])
print(mylist)
