'''
Created on Feb 21, 2012

@author: Joselle Abagat
'''

# this class will allow me to manipulate what i've passed in the ContractEntry class as a whole matrix (list of lists)
class Contracts(object):
    '''
    class docs
    '''
    
    def __init__(self, filepath):
        self.contract_list = self.parseFile(self, filepath)
    
    def parseFile(self, filepath):
        '''
        ' FUNCTION TO LOAD CONTRACTS; need to be in CSV format"
        '''
             
        filename = open(filepath,'r')
        Contractlist = []
        
        for ContractLine in filename:
            if not ContractLine.startswith('//'):
                ContractItem = ContractLine.split(',')
                Contractlist.append(Contract(ContractItem[0].strip(), ContractItem[1].strip()))
        
        return Contractlist

class Contract(object):
    '''
    classdocs
    '''

    def __init__(self, ID, Contract):
        '''
        Constructor
        '''
        self.ID = ID
        self.Contract = Contract
    
    def __str__(self):
        
        s = "\nContract:" + \
            "\n\tID: " + str(self.ID) + \
            "\n\tContract: " + str(self.Contract)
        
        return s
    
    def __eq__(self, test_case):
        """
        Overloads the equals operators.  Ensures that all instance variables
        are equal.
        """
        
        if self.Contract != test_case.Contract:
            return False
        if self.ID != test_case.ID:
            return False
        
        return True
        
#    def ContractDictionary(self):
#        '''
#        Returns a Dictionary of ID and Contract in the form {ID: Contract}
#        '''
#        Contractdict = {self.ID: self.Contract}
#        
#        return Contractdict
#
#    def ContractItem(self):
#        '''
#        ' Returns a list of the Contracts in the form [Contract]
#        '''
#        
#        lContract = self.Contract
#        
#        return lContract
#    
#    def ContractList(self):
#        '''
#        ' Returns a list of ID and Contracts in the form [ID, Contract]
#        '''
#        
#        lContract = [self.ID, self.Contract]
#        return lContract
#    
#    def CallContractsDict(self, loaded_file):
#        '''
#        ' CALL CONTRACT DICTIONARY
#        '''
#        
#        lContract = []
#        dictContract = {}
#        
#        for item in loaded_file:
#            
#            lContract = self.ContractList(item)
#            #print(lContract)
#            dictContract[int(lContract[0])] = lContract[1]
#        
#        #print(dictContract.items())
#        return dictContract
#    
#    def CallContracts(self,loaded_file):
#        '''
#        Returns a list of all the contracts: [Contract Name]
#        '''
#        
#        lContracts = []
#        
#        for each_item in loaded_file:
#            if not self.ContractItem(each_item) in lContracts:
#                lContracts.append(self.ContractItem(each_item))
#        
#        return lContracts