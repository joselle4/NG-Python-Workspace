'''
Created on Mar 8, 2012

@author: G73080
'''
import urllib2

class CAM(object):
    '''
    classdocs
    '''
    
    def __init__(self, ID, camCode, firstName, lastName):
        '''
        Constructor
        '''
        self.ID = ID
        self.camCode = camCode
        self.firstName = firstName
        self.lastName = lastName
        
def rapid_download(propID):
    '''doesn't work'''
#   url = "http://rapidreports01.northgrum.com/WQ.ASP?PH=1&PERSONID=2863&PROPID=" + str(propID) + "&RPTCD=EToolV1&OUTPUTTYPE=TABLE&SOURCEAPP=EXCEL&COLNAMES=1"
    url = "http://dicakite132859.northgrum.com/apps"

    try:
        f = urllib2.urlopen(url)
    except IOError, e:
        print("%s: Couldn't connect to %s\n" % (e, url))
        exit(0)
    source_code = str(f.read())
    f.close()
    file = open("RAPID_ID_" + str(propID) + "download.csv","w")
    file.write(source_code)
    file.close()
    
rapid_download(1234)