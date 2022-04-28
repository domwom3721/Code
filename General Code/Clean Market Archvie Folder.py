#We will use this to clear the area, neighborhood, and market archvie folders. NOT the legacy archive but the current archvies
#Author: Mike Leahy
#Date: 9/24/2021
import os


market_archive_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Market','Archive','2021 Q4') 
assert os.path.exists(market_archive_root)

#Loop through the folders in the archive folder and delete any empty ones
for i in range(10):
    for (dirpath, dirnames, filenames) in os.walk(market_archive_root):
        if dirnames == [] and filenames == [] and dirpath != market_archive_root:
            print('Deleting ',dirpath)
            os.rmdir(dirpath)

#After we cleaned up the archive folders, we can delete the files in the main area and hood folders so we can preserve the folder structure
market_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Market') 
assert os.path.exists(market_root)


for (dirpath, dirnames, filenames) in os.walk(market_root):
    if ('Archive'   in dirpath):
        continue
    for file in filenames:
        if ('2022 Q1' in file): #skip files that are for the most current quarter
            continue
        print('Deleting',os.path.join(dirpath,file))
        os.remove(os.path.join(dirpath,file))
