import urllib
import bz2
import os

print "Downloading the latest data from 'https://www.fuzzwork.co.uk/dump/' ..."
zipfile_name = 'invTypes.xls.bz2'
urllib.urlretrieve('https://www.fuzzwork.co.uk/dump/latest/invTypes.xls.bz2',zipfile_name)

print "Decompressing the 'bz2' file ..."
zipfile = bz2.BZ2File(zipfile_name)
data = zipfile.read()
newfile_name = zipfile_name[:-4]
newfile = open(newfile_name,'wb').write(data)

print "Removing the 'bz2' file ..."
os.remove(zipfile_name)
