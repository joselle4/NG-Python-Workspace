'''
Created on Feb 20, 2012

@author: Jesus Medrano (medraje1) 
'''

CAMS = []

class CAM(object):
    '''
    classdocs
    '''
        
    def __init__(self, ID, CAMCode, CAMFirstName, CAMLastName):
        
        self.ID = ID
        self.CAMCode = CAMCode
        self.CAMFirstName = CAMFirstName
        self.CAMLastName = CAMLastName
    
    def __str__(self):
        
        s = "\nCAM:" + \
            "\n\tID: " + str(self.ID) + \
            "\n\tCAM Code: " + self.CAMCode + \
            "\n\tFirst Name: " + self.CAMFirstName + \
            "\n\tLast Name: " + self.CAMLastName
            
        return s
    
def readCSVCAM(filePath):
    
    f = open(filePath, "r")
    
    for line in f:
        if not line.startswith("//"):
            values = line.split(",")
            CAMS.append(CAM(values[0], values[1], values[2], values[3]))

    

readCSVCAM("CAMS.csv")

for c in CAMS:
    print(c)

        