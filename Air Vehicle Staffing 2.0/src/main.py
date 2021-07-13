'''
Created on Mar 14, 2012

@author: G73080
'''

from Proposal import *

class AllProposalData(object):
    '''
    classdocs
    '''
     
    def __init__(self):
        '''
        Constructor
        '''
        self.rawDataPath = "../RawProposalData.csv"
        
        self.proposalList = []
    
    def ImportFromRawData(self):
        try:
            data = open(self.rawDataPath, "r")
        except IOError, e:
            print("Error opening raw data file at %s: %s" % (self.rawDataPath, e[1]))
            exit(0)
        for line in data:
            fields = line.split("|") # delimiter is "|", separate data based on that
            fields[6] = fields[6][:-1] # get rid of newline character
            if fields[0] != "ProposalName":
                id = 0
                for proposal in self.proposalList:
                    if proposal.proposalName == fields[0]:
                        id = proposal.ID
                        self.NewTask(fields, id)
                if id == 0:
                    self.NewProposal(fields)
                
    def NewProposal(self, fields):
        IDnum = 0
        for proposal in self.proposalList:
            if proposal.ID > IDnum:
                IDnum = proposal.ID
        IDnum = IDnum + 1
        newProposal = Proposal(IDnum, fields[0], fields[1])
        newTask = Task(fields[2], fields[3], fields[4], fields[5], fields[6])
        newProposal.AddTask(newTask)
        self.proposalList.append(newProposal)
        
    def NewTask(self, fields, propID):
        for proposal in self.proposalList:
            if proposal.ID == propID:
                newTask = Task(fields[2], fields[3], fields[4], fields[5], fields[6])
                proposal.AddTask(newTask)
                
    def ProposalReport(self):
        for proposal in self.proposalList:
            print(proposal)

    def TaskReport(self):
        for proposal in self.proposalList:
            print(proposal)
            proposal.ShowReport()
                
newData = AllProposalData()
newData.ImportFromRawData()

newData.TaskReport()