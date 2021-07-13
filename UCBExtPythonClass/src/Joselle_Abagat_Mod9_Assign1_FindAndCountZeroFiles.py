'''
Created on Mar 16, 2012

@author: Joselle Abagat
'''

import os

def read_dir(basedir):

    files = os.listdir(basedir)

    counter = 0

    for root, dirs, files in os.walk(basedir):
        for eachfile in files:
            if os.path.getsize(os.path.join(root, eachfile)) == 0:
                counter += 1
                print("empty file: " + eachfile)
    print("There are " + str(counter) + " zero length files.")

filedir = "..//..//"
print(os.path.abspath(filedir))
read_dir(filedir)


