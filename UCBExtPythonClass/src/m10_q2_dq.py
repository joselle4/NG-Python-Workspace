from collections import *

class Ulist(UserList):

    def __add__(self, other):
        extra = []
        if isinstance(other, UserList):
            for item in other.data:
                if not item in self.data:
                    extra.append(item)
            return self.__class__(self.data + extra)
        elif isinstance(other, type(self.data)): #this section may not be needed
            for item in other:
                if not item in self.data:
                    extra.append(item)
            return self.__class__(self.data + extra)
        else:
            for item in other:
                if not item in self.data:
                    extra.append(item)
            return self.__class__(self.data + extra)

    def append(self, item):
        if not item in self.data:
            self.data.append(item)

    def extend(self, other):
        if isinstance(other, UserList):
            for item in other.data:
                if not item in self.data:
                    self.data.extend([item])
        else:
            for item in other:
                if not item in self.data:
                    self.data.extend([item])
