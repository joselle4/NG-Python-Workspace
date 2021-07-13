'''
Created on Oct 31, 2012

@author: g73666
'''

class MPM(object):
    '''
    MPM Data Parser
    '''


    def __init__(self):
        '''
        Constructor
        '''
        pass

class MPMEntry(object):
    ''' field constructor:
    PROJECT WBS ID  description ALIAS   E   CAM CHARGE #    Network resp    perf    YYYYMM  
    BCWS HRS/UTS    ACT HRS/UTS BCWP HRS/UTS    ETC HRS/UTS 
    BCWS PRIME  BCWP PRIME  ACT PRIME   ETC PRIME   
    BCWS T DLRS BCWP T DLRS ACT T DLRS  ETC T DLRS  
    EV % COMPL  clin    clin    E   SCE
    '''
    
    def __init__(self, contract, resp, wbsid, description, cec, perf, clin, networkno, period, BCWShrs, BCWPhrs, ACThrs, ETChrs, elem):
        '''
        Constructor
        '''
        
        self.contract = contract
        self.resp = resp
        self.wbsid = wbsid
        self.description = description
        self.cec = cec
        self.perf = perf
        self.clin = clin
        self.networkno = networkno
        self.period = period
        self.BCWShrs = BCWShrs
        self.BCWPhrs = BCWPhrs
        self.ACThrs = ACThrs
        self.ETChrs = ETChrs
        self.elem = elem
    
    