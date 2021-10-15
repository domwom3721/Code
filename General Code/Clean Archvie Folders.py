#We will use this to clear the area, neighborhood, and market archvie folders. NOT the legacy archive but the current archvies
#Author: Mike Leahy
#Date: 9/24/2021
import os

# area_archive_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Area','Archive','2021 Q2') 
# assert os.path.exists(area_archive_root)


# hood_archive_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Neighborhood','Archive','2021') 
# assert os.path.exists(hood_archive_root)

market_archive_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Market','Archive','2021 Q2') 
assert os.path.exists(market_archive_root)


for i in range(10):
    # for (dirpath, dirnames, filenames) in os.walk(area_archive_root):
    #     if dirnames == [] and filenames == [] and dirpath != area_archive_root :
    #         print('Deleting ',dirpath)
    #         os.rmdir(dirpath)


    # for (dirpath, dirnames, filenames) in os.walk(hood_archive_root):
    #     if dirnames == [] and filenames == [] and dirpath != hood_archive_root:
    #         print('Deleting ',dirpath)
    #         os.rmdir(dirpath)
    
    for (dirpath, dirnames, filenames) in os.walk(market_archive_root):
        if dirnames == [] and filenames == [] and dirpath != market_archive_root:
            print('Deleting ',dirpath)
            # os.rmdir(dirpath)

#After we cleaned up the archive folders, we can delete the files in the main area and hood folders so we can preserve the folder structure
# hood_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Neighborhood') 
# assert os.path.exists(hood_root)

# area_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Area') 
# assert os.path.exists(area_root)

market_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Market Analysis','Market') 
assert os.path.exists(market_root)



# for (dirpath, dirnames, filenames) in os.walk(hood_root):
#     if 'Archive' in dirpath:
#         continue
#     for file in filenames:
#         print('Deleting',file)
#         # os.remove(file)

# for (dirpath, dirnames, filenames) in os.walk(area_root):
#     if 'Archive' in dirpath:
#         continue
#     for file in filenames:
#         print('Deleting',file)
#         # os.remove(file)

for (dirpath, dirnames, filenames) in os.walk(market_root):
    if 'Archive' in dirpath:
        continue
    for file in filenames:
        print('Deleting',file)
        # os.remove(file)
