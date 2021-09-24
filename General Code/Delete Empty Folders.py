#We will use this to clear the area and neighborhood archvie folders. NOT the legacy archive but the current archvies
#Author: Mike Leahy
#Date: 9/24/2021
import os

area_archive_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Area','Archive','2021 Q2') 
assert os.path.exists(area_archive_root)


hood_archive_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Neighborhood','Archive','2021') 
assert os.path.exists(hood_archive_root)


for i in range(10):

    for (dirpath, dirnames, filenames) in os.walk(area_archive_root):
        if dirnames == [] and filenames == [] and dirpath != area_archive_root :
            print('Deleting ',dirpath)
            os.rmdir(dirpath)


    for (dirpath, dirnames, filenames) in os.walk(hood_archive_root):
        if dirnames == [] and filenames == [] and dirpath != hood_archive_root:
            print('Deleting ',dirpath)
            os.rmdir(dirpath)