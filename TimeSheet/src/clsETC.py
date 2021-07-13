'''
Created on Feb 23, 2012

@author: Joselle Abagat
'''

import csv, re, sys, os
import easygui as egui
import wx
from clsGUI import *

class ETC(object):
    '''
    class docs
    '''
    
    def __init__(self):
        
        self.etc_list = []
        self.networkList = []
        self.camList = []
        self.unknownNetworkList = []
        self.dict_networkToContract = {} #self.mapNetworkToContract()
        self.dict_networkToPOP = {} #self.mapNetworkToPOP()
        self.dict_networkToWBS = {} #self.mapNetworkToWBS()
        self.dict_networkToCAM = {} #self.mapNetworkToCAM()        
        self.dict_networkToCAMPeriod = {} #self.mapNetworkToCAMPeriod()
        self.dict_networkCAMPeriodETC = {} #self.mapNetworkToETC()
        self.dict_networkDescription = {} #self.mapNetworkToDescription()
    
    def parseFile(self, filepath):
        '''
        Loads file into instantation function
        '''
        
        filename = open(filepath, 'r')
        read_csv = csv.reader(filename, dialect = "excel", quotechar = '"', delimiter = ",")
        
        file_header = read_csv.next()
        #print(file_header)
        contract_index = self.SearchTitle(file_header, "PROJECT")
        resp_index = self.SearchTitle(file_header, "RESP") 
        wbs_index = self.SearchTitle(file_header, "WBS ID")
        desc_index = self.SearchTitle(file_header, "DESCRIPTION")
        cec_index = self.SearchTitle(file_header, "CEC")
        perf_index = self.SearchTitle(file_header, "PERF")
        clin_index = self.SearchTitle(file_header, "CLIN")
        network_index = self.SearchTitle(file_header, "CHARGE")
        period_index = self.SearchTitle(file_header, "YYYYMM")
        bcws_index = self.SearchTitle(file_header, "BCWS HRS/UTS")
        bcwp_index = self.SearchTitle(file_header, "BCWP HRS/UTS")
        act_index = self.SearchTitle(file_header, "ACT HRS/UTS")
        etc_index = self.SearchTitle(file_header, "ETC HRS/UTS")
        elem_index = self.MatchTitle(file_header, "E")
        
#        if readfile[0] in readfile:
#            item = readfile[0].split(',')
#            contract_index = self.SearchTitle(item, "PROJECT")
#            resp_index = self.SearchTitle(item, "RESP") 
#            wbs_index = self.SearchTitle(item, "WBS ID")
#            desc_index = self.SearchTitle(item, "DESCRIPTION")
#            cec_index = self.SearchTitle(item, "CEC")
#            perf_index = self.SearchTitle(item, "PERF")
#            clin_index = self.SearchTitle(item, "CLIN")
#            network_index = self.SearchTitle(item, "CHARGE")
#            period_index = self.SearchTitle(item, "YYYYMM")
#            bcws_index = self.SearchTitle(item, "BCWS HRS/UTS")
#            bcwp_index = self.SearchTitle(item, "BCWP HRS/UTS")
#            act_index = self.SearchTitle(item, "ACT HRS/UTS")
#            etc_index = self.SearchTitle(item, "ETC HRS/UTS")
#            elem_index = self.MatchTitle(item, "E")
        
        for item in read_csv:
#            if not line == readfile[0]:
#                item = line.split(',')
            contract = self.FindInList(item, contract_index)
            resp = self.FindInList(item, resp_index)
            wbs = self.FindInList(item, wbs_index)
            desc = self.FindInList(item, desc_index)
            cec = self.FindInList(item, cec_index)
            perf = self.FindInList(item, perf_index)
            clin = self.FindInList(item, clin_index)
            network = self.FindInList(item, network_index)
            period = self.FindInList(item, period_index)
            bcws = self.FindInList(item, bcws_index)
            bcwp = self.FindInList(item, bcwp_index)
            act = self.FindInList(item, act_index)
            etc = self.FindInList(item, etc_index)
            elem = self.FindInList(item, elem_index)
            
            if network != '':
                self.etc_list.append(ETCEntry(contract, resp, wbs, desc, cec, perf, clin, network, period, bcws, bcwp, act, etc, elem))
                #print(contract + '\n' + resp + '\n' + wbs + '\n' + desc + '\n' + cec + '\n' + perf + '\n' + clin + '\n' + network + '\n' + period + '\n' + bcws + '\n' + bcwp + '\n' + act + '\n' + etc + '\n' + elem)
        
        filename.close()
        
    def FindInList(self, listArgument, index):
        """
            method to strip a list and assign an element of the list to a variable; 
            if it's a type error, it will return ""
        """
        
        try:
            return listArgument[index].strip()
        except TypeError:
            return ""
        except IndexError:
            return ""    
    
    def MatchTitle(self, lMatch, Title):
        '''
        returns the position of the title in the list of the CSV file
        '''
        
        ListTitle = []
        
        for item in lMatch:
            ListTitle.append(item.strip())
        
        index = 0
        for item in ListTitle:
            
            if item == Title:
                return index
            index = index + 1
        
        return None
            
    def SearchTitle(self, lMatch, Title):
        '''
        returns the position of the title in the list of the CSV file
        '''
        
        ListTitle = []
        
        for item in lMatch:
            ListTitle.append(item.strip())
        
        index = 0
        for item in ListTitle:
            regex = re.compile(Title, re.I)
            findre = regex.search(item)
            
            if not findre == None:
                return index
            index = index + 1
        
        return None
    
    def getCAMs(self):
        ''' generates list of cams'''
        for each_entry in self.etc_list:
            if each_entry.RESP != " " and len(each_entry.RESP) == 3 and each_entry.RESP != '\n':
                if each_entry.RESP not in self.camList:
                    self.camList.append(each_entry.RESP)
                    self.camList.sort(cmp=None, key=None, reverse=False)

        
    def getNetworks(self, camChoices):
        '''will generate list of networks for BW namerun of the form: Network1; Network2; Network3; ...'''
        # will write into file and open that file for the planners to use; if the file does not exist, create a new
        # one, otherwise, provide an option for the planner to use the existing list of networks (say for a new month, etc)
        
        for each_entry in self.etc_list:
            try:
                for cam in camChoices:
                    if each_entry.RESP == cam:
                        if each_entry.NetworkNo != None and each_entry.NetworkNo != "" and each_entry.NetworkNo != "\n" and each_entry.NetworkNo != " " and len(each_entry.NetworkNo) == 9:
                            if not each_entry.NetworkNo in self.networkList:
                                self.networkList.append(each_entry.NetworkNo)
            except TypeError:
                egui.msgbox("No CAM(s) selected", "Program Terminated", "OK", None, None)
                sys.exit(0)
        
    def generateNetworksString(self):
        '''will generate list of networks for BW namerun of the form: Network1; Network2; Network3; ...'''
        
        BWstring = ""
        for eachNetwork in self.networkList:
            BWstring = BWstring + str(eachNetwork) + "; "
        
        return BWstring
    
    def emptyNetworkList(self): self.networkList = []
    
    def mapNetworkToContract(self):
        """
        will update dict_networkToContract dictionary of the form {networkNo1: Contract name, ..., networkNon: Contract name}
        """
        
        for each_entry in self.etc_list:
            if not each_entry.NetworkNo in self.dict_networkToContract:
                self.dict_networkToContract[each_entry.NetworkNo] = each_entry.Contract
    
    def mapNetworkToPOP(self):
        """
        will update self.dict_networkToPOP dictionary with the very last period associated with the given network number 
        (it will assume it is the end of the POP)
        """
        
        for each_entry in self.etc_list:
            if not each_entry.NetworkNo in self.dict_networkToPOP:
                try:
                    self.dict_networkToPOP[each_entry.NetworkNo] = int(each_entry.Period)
                except ValueError:
                    print(each_entry.Period + " is not a period")
            else:
                # loop through all the existing dictionary keys
                # check that we are comparing key to key
                # if so, compare the periods and replace with the larger value
                for network_key in self.dict_networkToPOP.keys():
                    if network_key == each_entry.NetworkNo:
                        try:
                            if self.dict_networkToPOP[each_entry.NetworkNo] < int(each_entry.Period):
                                self.dict_networkToPOP[each_entry.NetworkNo] = int(each_entry.Period)
                        except ValueError:
                            print "Period: '" + each_entry.Period + "'"
    
    def mapNetworkToWBS(self):
        """
        updates self.dict_networkToWBS dictionary of the form: {
                                                                    network1: [WBS1, WBS2, ],
                                                                    network2: [WBS1, WBS2, ],
                                                                    ...
                                                                }
        """
        
        for each_entry in self.etc_list:
            if not each_entry.NetworkNo in self.dict_networkToWBS:
                self.dict_networkToWBS[each_entry.NetworkNo] = [each_entry.WBSID]
            else:
#                print "WBS-NETWORK DICT: " 
#                print self.dict_networkToWBS[each_entry.NetworkNo]
                if each_entry.WBSID not in self.dict_networkToWBS[each_entry.NetworkNo]:
                    self.dict_networkToWBS[each_entry.NetworkNo].append(each_entry.WBSID)
    
    def mapNetworkToCAM(self):
        """
                                                                   lv-1     
        updates self.dict_networkToCAM dictionary of the form: { network1: [CAM1, CAM2],
                                                                 network2: [CAM1, CAM2, CAM3],
                                                                            ...}
                                                                ...
                                                               }
        """
        
        for each_entry in self.etc_list:
            if not each_entry.NetworkNo in self.dict_networkToCAM:
                self.dict_networkToCAM[each_entry.NetworkNo] = [each_entry.RESP]
            else:
                if each_entry.RESP not in self.dict_networkToCAM[each_entry.NetworkNo]:
                    self.dict_networkToCAM[each_entry.NetworkNo].append(each_entry.RESP)
    
    # we want to use this to create a dictionary within a dictionary
    def mapNetworkToCAMPeriod(self):
        """
                                                                   lv-1     lv-2
        updates self.dict_networkToCAM dictionary of the form: { network1: {CAM1: [],
                                                                           CAM2: [],},
                                                                network2: {CAM1: [],
                                                                            ...}
                                                                ...
                                                               }
        """
        
        for each_entry in self.etc_list:
            if not each_entry.NetworkNo in self.dict_networkToCAMPeriod:
                self.dict_networkToCAMPeriod[each_entry.NetworkNo] = {each_entry.RESP: []}
            else:
                if each_entry.RESP not in self.dict_networkToCAMPeriod[each_entry.NetworkNo]:
                    # does not really matter if RESP level2 key is repeated since keys are unique.  
                    # but will check anyway for efficiency
                    self.dict_networkToCAMPeriod[each_entry.NetworkNo][each_entry.RESP] = []
    
    # we have to return something for this one since it'll be overwritten by self.mapNetworkToCAM anyway
    def mapNetworkToETC(self):
        """
                                             lv-1     lv-2    lv-3
        returns a dictionary of the form: {network1: {CAM1: {yyyymm1: hrs, 
                                                             yyyymm2: hrs}, 
                                                      CAM2: {yyyymm1: hrs, 
                                                             yyyymm2: hrs}}, 
                                           network2: {CAM1: {yyyymm1: hrs, 
                                                             yyyymm2: hrs}, 
                                                      CAM2: {yyyymm1: hrs, 
                                                             yyyymm2: hrs}}, 
                                                             ..., }
        creates a 3:1 lookup to the ETC value
        period keys: dictionary[network][camcode].keys()
        hrs: dictionary[network][camcode][period]
        
        note: this is a one:one:one:one match between network:cam:period:hrs
        """
        
        if self.dict_networkCAMPeriodETC == {}:
            self.dict_networkCAMPeriodETC = self.dict_networkToCAMPeriod
        else:
            # check if the network is not one the tier I keys.  if it's not, then add the network number as a key
            # and the contents of that key from self.dict_networkToCAM as the value
            for each_entry in self.etc_list:
                if not each_entry.NetworkNo in self.dict_networkCAMPeriodETC:
                    self.dict_networkCAMPeriodETC[each_entry.NetworkNo] = self.dict_networkToCAMPeriod[each_entry.NetworkNo] 
        
        # now go through each and 
        for each_entry in self.etc_list:
            if self.dict_networkCAMPeriodETC[each_entry.NetworkNo][each_entry.RESP] == []:
                # add lv-3
                try:
                    self.dict_networkCAMPeriodETC[each_entry.NetworkNo][each_entry.RESP] = {int(each_entry.Period): float(each_entry.ETChrs)}
                except ValueError:
                    print("Cannot convert " + str(each_entry.Period) + " and/or " + str(each_entry.ETChrs) + "to numbers.  Check raw data")
            else:
                # if it's not empty, add only the value of the level 3           
                try:
                    self.dict_networkCAMPeriodETC[each_entry.NetworkNo][each_entry.RESP][int(each_entry.Period)] = float(each_entry.ETChrs)
                except TypeError:
                    print(self.dict_networkCAMPeriodETC[each_entry.NetworkNo])
                    print(self.dict_networkCAMPeriodETC[each_entry.NetworkNo][each_entry.RESP])
                except ValueError:
                    print("Cannot convert " + str(each_entry.Period) + " and/or " + str(each_entry.ETChrs) + "to numbers.  Check raw data")
            #print each_entry.NetworkNo + " " + each_entry.RESP + " " + each_entry.Period
            #print self.dict_networkCAMPeriodETC[each_entry.NetworkNo][each_entry.RESP][int(each_entry.Period)]
            
    def mapNetworkToDescription(self):
        """
        returns a dictionary of the network numbers and their descriptions
        """
        
        for each_entry in self.etc_list:
            if not each_entry.NetworkNo in self.dict_networkDescription:
                self.dict_networkDescription[each_entry.NetworkNo] = each_entry.Description
        
    def getContractFromNetwork(self, network_number):
        """
        returns the contract associated with the given network
        """
        
        try:
            return self.dict_networkToContract[network_number]
        except KeyError:
            #print("Unknown Contract: " + network_number + " not found in the ETC file.")
            if network_number not in self.unknownNetworkList:
                self.unknownNetworkList.append(network_number)
            return "Not in MPM Export"

    #we want to use this method to write into the network source file and add new networks and/or edit incorrect POPs
    def getPOPfromNetwork(self, network_number):
        """
        returns the end of the period of performance associated with the given network
        """
        
        try:
            return self.dict_networkToPOP[network_number]
        except KeyError:
            #print("Unknown POP. " + network_number + " not found in the ETC file.")
            if network_number not in self.unknownNetworkList:
                self.unknownNetworkList.append(network_number)
            return "Not in MPM Export"

    def getETCFromNetworkCAMAndPeriod(self, network_number, cam_code, period):
        """
        returns the hours of etc given the network number, cam, and the period
        NOTE: certain network cams may contain multiple cam_codes therefore, need to split and sum up ETCs
        """
        camlist = []
        
 
        if len(cam_code) < 4: 
            try:
                return self.dict_networkCAMPeriodETC[network_number][cam_code][period]
            except KeyError:
                print "No ETC found for " + network_number + " " + cam_code + " on " + str(period)
                if network_number not in self.unknownNetworkList:
                    self.unknownNetworkList.append(network_number)                
        else:
            camlist = cam_code.split("; ")
            if len(camlist) != 1:
                sumETC = 0
                for cam in camlist:
                    try:
                        sumETC = self.dict_networkCAMPeriodETC[network_number][cam][period] + sumETC
                    except KeyError:
                        print "No ETC found for " + network_number + " " + cam + " on " + str(period)
                        if network_number not in self.unknownNetworkList:
                            self.unknownNetworkList.append(network_number)                    
                return sumETC
            else:
                try:
                    return self.dict_networkCAMPeriodETC[network_number][cam_code][period]
                    print self.dict_networkCAMPeriodETC[network_number][cam_code][period]
                except KeyError:
                    print "No ETC found for " + network_number + " " + cam_code + " on " + str(period)
                    if network_number not in self.unknownNetworkList:
                        self.unknownNetworkList.append(network_number)
            #return 0
    
    def getDescriptionFromNetwork(self, network_number):
        """
        returns the description of a network number
        """
        
        try:
            return self.dict_networkDescription[network_number]
        except KeyError:
            #print("Cannot find description for " + network_number + ". Network number does not exist in the raw data")
            if network_number not in self.unknownNetworkList:
                self.unknownNetworkList.append(network_number)
            return "Not in MPM Export"
    
    def getCAMsFromNetwork(self, network_number):
        """
        returns the CAM or CAMs under the network number
        """
        
        try: 
            #If there are multiple cams in one network, convert the list into a string
            camCount = len(self.dict_networkToCAM[network_number])
            if camCount == 1:
                return self.dict_networkToCAM[network_number][0]
            else:
                camString = self.dict_networkToCAM[network_number][0]
                for count in range(0,camCount):
                    count = count + 1
                    if count < camCount:
                        camString = camString + "; " + self.dict_networkToCAM[network_number][count]
                return camString
        except KeyError:
            if network_number not in self.unknownNetworkList:
                self.unknownNetworkList.append(network_number)
            return "Not in MPM Export"
    
    def getWBSFromNetwork(self, network_number):
        """
        returns the WBS(s) under the network number
        """
        
        try:
            wbsCount = len(self.dict_networkToWBS[network_number])
            if wbsCount == 1:
                return self.dict_networkToWBS[network_number][0]
            else:
                wbsString =  self.dict_networkToWBS[network_number][0]
                for count in range(0, wbsCount):
                    count = count + 1
                    if count < wbsCount:
                        wbsString = wbsString + "; " + self.dict_networkToWBS[network_number][count]
                return wbsString
        except KeyError:
            if network_number not in self.unknownNetworkList:
                self.unknownNetworkList.append(network_number)
            return "Not in MPM Export"
    
class ETCEntry(object):
    '''
    requires 
    '''

    def __init__(self, Contract, RESP, WBSID, Description, CEC, PERF, CLIN, NetworkNo, Period, BCWShrs, BCWPhrs, ACThrs, ETChrs, elem):
        '''
        Constructor
        '''
        
        self.Contract = Contract
        self.RESP = RESP
        self.WBSID = WBSID
        self.Description = Description
        self.CEC = CEC
        self.PERF = PERF
        self.CLIN = CLIN
        self.NetworkNo = NetworkNo
        self.Period = Period
        self.BCWShrs = BCWShrs
        self.BCWPhrs = BCWPhrs
        self.ACThrs = ACThrs
        self.ETChrs = ETChrs
        self.elem = elem
    
    def __str__(self):
        
        s = "\nMPM Data" + \
            "\n\tContract: " + str(self.Contract) + \
            "\n\tResponsible CAM: " + str(self.RESP) + \
            "\n\tWBS ID: " + str(self.WBSID) + \
            "\n\tDescription: " + str(self.Description) + \
            "\n\tCEC: " + str(self.CEC) + \
            "\n\tPerforming: " + str(self.PERF) + \
            "\n\tCLIN: " + str(self.CLIN) + \
            "\n\tNetwork Number: " + str(self.NetworkNo) + \
            "\n\tPeriod: " + str(self.Period) + \
            "\n\tBCWS: " + str(self.BCWShrs) + \
            "\n\tBCWP: " + str(self.BCWPhrs) + \
            "\n\tACT: " + str(self.ACThrs) + \
            "\n\tETC: " + str(self.ETChrs) + \
            "\n\tElement: " + str(self.elem)
            
        return s
    
    def __eq__(self, test_case):
        """
        Overloads the equals operators.  Ensures that all instance variables
        are equal.
        """
        
        if self.ACThrs != test_case.ACThrs:
            return False
        if self.BCWPhrs != test_case.BCWPhrs:
            return False
        if self.BCWShrs != test_case.BCWShrs:
            return False
        if self.CEC != test_case.CEC:
            return False
        if self.CLIN != test_case.CLIN:
            return False
        if self.Contract != test_case.Contract:
            return False
        if self.Description != test_case.Description:
            return False
        if self.elem != test_case.elem:
            return False
        if self.ETChrs != test_case.ETChrs:
            return False
        if self.NetworkNo != test_case.NetworkNo:
            return False
        if self.PERF != test_case.PERF:
            return False
        if self.Period != test_case.Period:
            return False
        if self.RESP != test_case.RESP:
            return False
        if self.WBSID != test_case.WBSID:
            return False
        
        return True
    
#    def MapNetworkToContract(self):
#        '''
#        returns a list of the form [Contract, Network No]
#        '''
#        lMapWBSToContract = [self.Contract, self.NetworkNo]
#        
#        return lMapWBSToContract
#    
#    def MapNetworkToCAM(self):
#        '''
#        returns a list of the form [CAM Code, Network No]
#        '''
#        lMapWBSToCAM = [self.RESP, self.NetworkNo]
#        
#        return lMapWBSToCAM
#    
#    def MapNetworkToETC(self):
#        '''
#        returns a list of the form [Network No, Period, ETC]
#        '''
#        
#        lMapWBSToETC = [self.NetworkNo, self.Period, self.ETChrs]
#        
#        return lMapWBSToETC
#        
#    def ReportData(self):
#        '''
#        returns a list of the form [Contract, Responsible CAM, Network Number, Period, ETC]
#        '''
#        
#        lReportData = [self.Contract, self.Resp, self.NetworkNo, self.Period, self.ETChrs]
#        
#        return lReportData
#    
#    def CallNetworkContract(self, loaded_file):
#        """
#        function returns list mapping between network numbers and the contract
#        """
#        
#        lmapping = []
#        
#        for item in loaded_file:
#            if not self.MapNetworkToContract(item) in lmapping:
#                lmapping.append(self.MapNetworkToContract(item))
#        
#        return lmapping
#    
#    def CallNetworkCAM(self, loaded_file):
#        """
#        This function returns a list mapping Networks to CAM
#        """
#        
#        lmapping = []
#        
#        for item in loaded_file:
#            if not self.MapNetworkToCAM(item) in lmapping:
#                lmapping.append(self.MapNetworkToCAM(item))
#        
#        return lmapping
#    
#    def CallNetworkToETC(self, loaded_file):
#        """
#        This functions returns a list of the networks and their ETCs
#        """
#        
#        lmapping = []
#        
#        for item in loaded_file:
#            if not self.MapNetworkToETC(item) in lmapping:
#                lmapping.append(self.MapNetworkToETC(item))
#        
#        return lmapping
#    
#    def CallReportData(self, loaded_file):
#        """
#        This function returns the required fields needed from the ETC data
#        """
#        
#        lmapping = []
#        
#        for item in loaded_file:
#            if not self.ReportData(item) in lmapping:
#                lmapping.append(self.ReportData(item))
#        
#        return lmapping
