'''
Created on Mar 17, 2012

@author: joselle4
'''

import re
import urllib.request
import sys

def count_images_in_url(url):
    
    try:
        f = urllib.request.urlopen(url)
    except IOError:
        sys.stderr.write("no connection")
        sys.exit
    contents = str(f.read())    
    f.close()

    getimage = re.compile(r'''< *img +src *=*["'](.+?)["']''', re.I)
    image = getimage.findall(contents)
    
    counter = 0
    for eachimage in image:
        counter += 1
        print("Image " + str(counter) + ": " + eachimage)
    
    imagelen = len(image)
    print("Total number of images in " + url + ": " + str(imagelen))
    
    return None

site = 'http://www.barnesandnoble.com'
count_images_in_url(site)    