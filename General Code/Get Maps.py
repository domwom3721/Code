#Get Maps
import os
import docx2txt
from os import walk

cd = r'C:\Users\Michael Leahy\Dropbox (Bowery)\Market Analysis\Neighborhood'
map_output =  r'C:\Users\Michael Leahy\Desktop\Maps'


#Get list of folders in market folder
market_folders = []
for (dirpath, dirnames, filenames) in walk(cd):
    market_folders.extend(dirnames)
    break

for folder in market_folders: #loop through each folder in the market folder
    if folder == '2019 Q3':
        continue
    print('The Folder is ',folder)
    
    #Get list of files in each folder
    files = []
    for (dirpath, dirnames, filenames) in walk(os.path.join(cd,folder)):
        files.extend(filenames)
        break

    recent_report = []
    #Loop through each folder and collect the relevant documents in a list (which ones may have a map)
    for file in files:
        if ('Q') in file and ('docx' in file) and (folder != '2019 Q3'):
            recent_report.append(file)

      
    #Loop through relevant files only
    for relevant_file in recent_report:
        relevant_file_clean = relevant_file.replace('.docx','')
        try:
            text = docx2txt.process(os.path.join(cd,folder,relevant_file), os.path.join(map_output)) #Read images from word doc
            os.rename(os.path.join(map_output,'image1.png'), os.path.join(map_output, relevant_file_clean+'.png')) #rename image1
        except Exception as e: print(e)

