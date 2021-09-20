#Get Maps
import os
from os import walk
import shutil

cd = r'C:\Users\Michael Leahy\Dropbox (Bowery)\Market Analysis\Market'


#Get list of folders in market folder
market_folders = []
for (dirpath, dirnames, filenames) in walk(cd):
    market_folders.extend(dirnames)
    break

for folder in market_folders: #loop through each folder in the market folder
    if folder == '2019 Q3' or 'rchive'in folder:
        continue
    print('')
    print('The Folder is ------------',folder, '------------')
    print('')
    new_archive_folder = os.path.join(cd,folder,'Archive For Cleaning')
    print('Making a new foler: ', new_archive_folder)
    print('')
    #os.mkdir(os.path.join(cd,folder,'Archive For Cleaning')) #Create Archvie Folder
    
    #Get list of files in each folder
    files = []
    for (dirpath, dirnames, filenames) in walk(os.path.join(cd,folder)):
        files.extend(filenames)
        break

    old_files = []
    #Loop through each folder and collect the relevant documents in a list (which ones may have a map)
    for file in files:
        if ('2020Q4') not in file and ('2020 Q4') not in file and ('2021') not in file:
            old_files.append(file)

      
    #Loop through relevant files only
    for old_file in old_files:
        print('')
        print('The file is ', old_file)
        print('')
        current_path = os.path.join(cd,folder,old_file)
        new_path     = os.path.join(cd,folder,'Archive For Cleaning',old_file)
        print('It is currently here ',current_path) #current file location
        print('We are moving it here ', new_path)     #Where we are moving file to
        
        #shutil.move(current_path, new_path)


