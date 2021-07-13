'''
Created on May 9, 2012

@author: G73080
'''

class Proposal(object):
    '''
    classdocs
    '''
     
    def __init__(self, ID, proposalName, contract):
        '''
        Constructor
        '''
        self.ID = ID
        self.proposalName = proposalName
        self.contract = contract
        
        self.tasks = []
    
    def AddTask(self, task):
        if isinstance(task, Task):
            task.PropID = self.ID
            self.tasks.append(task)
        else:
            print("Error: Invalid argument for AddTask - args must be Tasks")
        
    def ShowReport(self):
        for task in self.tasks:
            print task

    def __str__(self):
        s = "Proposal ID " + str(self.ID) + ": " + str(self.proposalName) + " (" + str(self.contract) + ")"
        return s
    
class Task(object):
    '''
    classdocs
    '''

    def __init__(self, CAM, cost_center, task_title, month, hours):
        '''
        Constructor
        '''
        self.CAM = CAM
        self.cost_center = cost_center
        self.task_title = task_title
        self.month = month
        self.hours = hours
        
        self.PropID = 0
        self.skillAlloc = {}
        
    def __str__(self):
        s = "CAM: " + str(self.CAM) + " - " + str(self.cost_center) + ": " + str(self.task_title) + " (" + str(self.month) + ") Hours: " + str(self.hours)
        return s