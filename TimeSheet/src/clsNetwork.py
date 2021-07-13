'''
Created on Feb 21, 2012

@author: Joselle Abagat
'''

#

# so the networks class will be updated not via a network file, but via an mpm file

class Networks(object):
    '''
    classdocs
    '''

    def __init__(self, filepath):
        self.network_list = self.parseFile(filepath)
        self.networkToContract_dict = self.mapNetworkToContract()
        self.networkToActivity_dict = self.mapNetworkToActivity()
        self.networkToCAM_dict = self.mapNetworkToCAMCode()
        self.networkToDescription_dict = self.mapNetworkToDescription()
        self.networkToPOP_dict = self.mapNetworkToPOP()
        
    def parseFile(self, filepath):
        ''' FUNCTION TO LOAD list of NETWORKS in csv format '''
        
        filename = open(filepath,'r')
        Networklist = []
        
        for NetworkLine in filename:
            if not NetworkLine.startswith('//'):
                NetworkItem = NetworkLine.split(',')
                Networklist.append(Network(NetworkItem[0].strip(), NetworkItem[1].strip(), NetworkItem[2].strip(), NetworkItem[3].strip(), NetworkItem[4].strip(), NetworkItem[5].strip(), NetworkItem[6].strip(), NetworkItem[7].strip(), NetworkItem[8].strip()))
        
        return Networklist
    
    def mapNetworkToContract(self):
        """ returns a dictionary of the form: {Network1: Contract Name, ..., Networkn: Contract Name} """
        
        dict = {}
        
        for item in self.network_list:
            if not item.NetworkNo in dict:
                dict[item.NetworkNo] = item.Contract
        
        return dict
    
    def getContractFromNetwork(self, network_number):
        """
        returns the contract name for a given network number """
        
        try:
            return self.networkToContract_dict[network_number]
        except KeyError:
            print(network_number + "not found. Please add to raw data")
    
    def mapNetworkToPOP(self):
        """ returns a dictionary of the form: {Network1: YYYYMM, ..., Networkn: YYYYMM} """
        
        ######################
        # NEED A WAY TO OBTAIN POP FOR EACH NETWORK; POSSIBILITY: ETC UPDATE FILES (GET LATEST PERIOD VALUE)
        ######################
        dict = {}
        
        for item in self.network_list:
            if not item.NetworkNo in dict:
                dict[item.NetworkNo] = int(item.POP)
        
        return dict
    
    def getPOPfromNetwork(self, network_number):
        """ returns the period for a given network number """
        
        try:
            return self.networkToPOP_dict[network_number]
        except KeyError:
            print(network_number + "not found. Please add to raw data")
    
    #INCOMPLETE
    def mapNetworkToCAMCode(self):
        """ returns a dictionary of the form: {Network1: [CAM1, CAM2], ..., Networkn: [CAM1, CAM2]} """
        
        dict = {}
        
        for item in self.network_list:
            if not item.NetworkNo in dict:
                dict[item.NetworkNo] = [item.CAMCode]
            else:
                if not "HXX" in dict[item.NetworkNo]:
                    dict[item.NetworkNo].append(item.CAMCode)
                elif item.CAMCode != "HXX":
                    dict[item.NetworkNo].append(item.CAMCode)
        
        return dict
    
    
    # INCOMPLETE
    def mapNetworkToActivity(self):
        """ returns a directory of the form: {Network1: [0030, H0PM, etc], ... Networkn: [xxxx, xxxx, ...]} """
        
        ####################
        # NEED A WAY TO OBTAIN ALL NETWORK-ACTIVITY COMBINATION; POSSIBILITY IS VIA BW NAMERUN
        #####################
        dict = {}
        
        for item in self.network_list:
            if not item.NetworkNo in dict:
                dict[item.NetworkNo] = [item.Activity]
            else:
                dict[item.NetworkNo].append(item.Activity)
        
        return dict
    
    def mapNetworkToDescription(self):
        """ returns a dictionary of the form: {Network1: Contract Name, ..., Networkn: Contract Name} """
        
        dict = {}
        
        for item in self.network_list:
            if not item.NetworkNo in dict:
                dict[item.NetworkNo] = item.Description
        
        return dict
    
    def getDescriptionFromNetwork(self, network_number):
        """ returns the description of the network """
        
        try:
            self.networkToDescription_dict[network_number]
        except KeyError:
            print(network_number + "not found. Please add to raw data")
            
    
    
    
class Network(object):
    '''
    classdocs
    '''
    
    def __init__(self, ID, Contract, CAMCode, NetworkNo, Activity, Description, PseudoWCC, POP, Status):
        '''
        Constructor
        '''
        
        self.ID = ID
        self.NetworkNo = NetworkNo
        self.Contract = Contract
        self.CAMCode = CAMCode
        self.PseudoWCC = PseudoWCC
        self.Status = Status
        self.POP = POP
        self.Activity = Activity
        self.Description = Description
        
    def __str__(self):
        s = "\nNetwork:" + \
            "\n\tID: " + str(self.ID) + \
            "\n\tContract: " + str(self.Contract) + \
            "\n\tCAM Code: " + str(self.CAMCode) + \
            "\n\tCharge Number: " + str(self.NetworkNo) + \
            "\n\tActivity Code: " + str(self.Activity) + \
            "\n\tDescription: " + str(self.Description) + \
            "\n\tPseudo Work Cost Center: " + str(self.PseudoWCC) + \
            "\n\tPOP: " + str(self.POP) + \
            "\n\tStatus: " + str(self.Status)
        
        return s
    
    def __eq__(self, test_case):
        """
        Overloads the equals operators.  Ensures that all instance variables
        are equal.
        """
             
        if self.Activity != test_case.Activity:
            return False
        if self.CAMCode != test_case.CAMCode:
            return False
        if self.Contract != test_case.Contract:
            return False
        if self.Description != test_case.Description:
            return False
        if self.ID != test_case.ID:
            return False
        if self.NetworkNo != test_case.NetworkNo:
            return False
        if self.POP != test_case.POP:
            return False
        if self.PseudoWCC != test_case.PseudoWCC:
            return False
        if self.Status != test_case.Status:
            return False
        
        return True
        
#    def MapNetworkToContract(self):
#        '''
#        returns a list of Network and Contract in the form [Contract, Network]
#        '''
#        
#        lNetworkToContract = [self.Contract, self.NetworkNo]
#        
#        return lNetworkToContract
#        
#    def MapNetworkToCAM(self):
#        '''
#        returns a list of Network and CAMCode in the form [CAMCode, Network]
#        '''
#        lNetworkToContract = [self.CAMCode, self.NetworkNo]
#        
#        return lNetworkToContract
#    
#    def MapNetworkToStatus(self):
#        '''
#        returns a list of Network and CAMCode in the form [CAMCode, Network]
#        '''
#        lNetworkToStatus = [self.NetworkNo, self.Status]
#        
#        return lNetworkToStatus
#    
#    def NetworkActivityPair(self):
#        '''
#        returns a list of Network and Activity Code in the form [Network, Activity, NetworkActivity]
#        '''
#        
#        lNetworkActivity = [self.NetworkNo, self.Activity, self.NetworkNo + self.Activity]
#        
#        return lNetworkActivity
#
#    def CallNetworkContract(self,loaded_file):
#        """
#        function that returns a mapping between Contracts and Networks
#        """
#        
#        lNetworkContract = []
#        
#        for each_item in loaded_file:
#            if not Network.MapNetworkToContract(each_item) in lNetworkContract:
#                lNetworkContract.append(Network.MapNetworkToContract(each_item))
#        
#        return lNetworkContract
#
#    def CallNetworkCAM(self,loaded_file):
#        """
#        Function that returns a mapping between networks and camc odes
#        """
#        
#        lNetworkCAM = []
#        
#        for each_item in loaded_file:
#            if not Network.MapNetworkToCAM(each_item) in lNetworkCAM:
#                lNetworkCAM.append(Network.MapNetworkToCAM(each_item))
#        
#        return lNetworkCAM
#    
#    def CallNetworkActivity(self, loaded_file):
#        """
#        Function that returns a list mapping between network and activity codes
#        """
#        lNetworkActivity = []
#        
#        for each_item in loaded_file:
#            if not Network.NetworkActivityPair(each_item) in lNetworkActivity:
#                lNetworkActivity.append(Network.NetworkActivityPair(each_item))
#        
#        return lNetworkActivity
#    
#    def CallNetworkStatus(self, loaded_file):
#        """
#        Function that returns a list mapping between the network and its status
#        """
#        
#        lNetworkStatus = []
#        
#        for each_item in loaded_file:
#            if not Network.MapNetworkToStatus(each_item) in lNetworkStatus:
#                lNetworkStatus.append(Network.MapNetworkToStatus(each_item))
#        
#        return lNetworkStatus