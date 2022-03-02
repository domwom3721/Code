#We will use this to clear the  neighborhood archive folders. NOT the legacy archive but the current archvies
#Author: Mike Leahy
#Date: 9/24/2021
import os


hood_archive_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Neighborhood','Archive') 
assert os.path.exists(hood_archive_root)


#Loop through the folders in the archive folder and delete any empty ones
for i in range(10):
    for (dirpath, dirnames, filenames) in os.walk(hood_archive_root):
        if dirnames == [] and filenames == [] and dirpath != hood_archive_root:
            print('Deleting ',dirpath)
            # os.rmdir(dirpath)
    


#After we cleaned up the archive folders, we can delete the files in the main area and hood folders so we can preserve the folder structure
# hood_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Market','Archive','Legacy Archive',statevar,quar_var1) 
hood_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Neighborhood') 
assert os.path.exists(hood_root)


for (dirpath, dirnames, filenames) in os.walk(hood_root):    
    if 'Archive' in dirpath:
        continue
    for file in filenames:
        print('Deleting: ',os.path.join(dirpath,file))
        # os.remove(os.path.join(dirpath,file))

