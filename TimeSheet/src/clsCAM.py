'''
Created on Feb 21, 2012

@author: Joselle Abagat
'''

class CAMs(object):
    '''
    classdocs
    '''
    
    def __init__(self, filepath):
        '''
        
        '''
        self.cam_list = self.parseFile(filepath)
        self.camOwnerDictionary = self.camOwnerMapping()
        
    def parseFile(self, filename):
        '''
        ' FUNCTION TO LOAD CAMS; need to be in CSV format"
        '''
        
        file = open(filename,'r')
        CAMlist = []
        
        for CAMLine in file:
            if not CAMLine.startswith('//'):
                CAMItem = CAMLine.split(',')
                CAMlist.append(CAM(CAMItem[0].strip(), CAMItem[1].strip(), CAMItem[2].strip(), CAMItem[3].strip()))
        
        return CAMlist
    
    def camOwnerMapping(self):
        
        dict = {}
        
        for cam in self.cam_list:
            if not cam.CAMCode in dict:
                dict[cam.CAMCode] = cam.CAMFullName
        
        return dict
    
    def getCAMOwner(self, cam_code):
        
        try:
            return self.camOwnerDictionary[cam_code]
        except KeyError:
            print(cam_code + " is not listed in the raw data.")

class CAM(object):
    '''
    classdocs
    '''

    def __init__(self, ID, CAMCode, CAMFirstName, CAMLastName):
        '''
        Constructor: needs the following parameters:
            ID, CAM Code, CAM's First Name, CAM's Last Name
        '''
        
        self.ID = ID
        self.CAMCode = CAMCode
        self.CAMFirstName = CAMFirstName
        self.CAMLastName = CAMLastName
        self.CAMFullName = self.CAMFirstName + " " + self.CAMLastName
    
    def __str__(self):
        
        s = "\nCAM:" + \
            "\n\tID: " + str(self.ID) + \
            "\n\tCAM Code: " + self.CAMCode + \
            "\n\tFirst Name: " + self.CAMFirstName + \
            "\n\tLast Name: " + self.CAMLastName
            
        return s

    def __eq__(self,test_case):
        """
        Overloads the equals operators.  Ensures that all instance variables
        are equal.
        """
        
        if self.CAMCode != test_case.CAMCode:
            return False
        if self.CAMFirstName != test_case.CAMFirstName:
            return False
        if self.CAMLastName != test_case.CAMLastName:
            return False
        if self.ID != test_case.ID:
            return False
        if self.CAMFullName != test_case.CAMFullName:
            return False
        
        return True
        
#    def ListCAMOwner(self):
#        '''
#        returns a list in the form [CAM Code, CAM Name]
#        '''
#        
#        lCAMOwner = [self.CAMCode, self.CAMFullName]
#        
#        return lCAMOwner
#    
#    def CallCAMOwner(self,loaded_file):
#        '''
#        Function to call the list of CAMs
#        '''
#        
#        listCAM = []
#        
#        for item in loaded_file:
#            if not self.ListCAMOwner(item) in listCAM:
#                listCAM.append(self.ListCAMOwner(item))
#        
#        return listCAM