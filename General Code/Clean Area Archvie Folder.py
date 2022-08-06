#We will use this to clear the area archive folders. NOT the legacy archive but the current archvies
#Date: 9/24/2021
import os

area_archive_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Area','Archive','2022 Q1') 
assert os.path.exists(area_archive_root)


#Loop through the folders in the archive folder and delete any empty ones
for i in range(10):
    for (dirpath, dirnames, filenames) in os.walk(area_archive_root):
        if dirnames == [] and filenames == [] and dirpath != area_archive_root :
            pass
            print('Deleting ',dirpath)
            os.rmdir(dirpath)


#After we cleaned up the archive folders, we can delete the files in the main area folders so we can preserve the folder structure
area_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Area') 
assert os.path.exists(area_root)


for (dirpath, dirnames, filenames) in os.walk(area_root):
    if 'Archive' in dirpath:
        continue
    for file in filenames:
        if ('2022 Q2' in file):
            continue
        print('Deleting',file)
        os.remove(os.path.join(dirpath,file))
        
