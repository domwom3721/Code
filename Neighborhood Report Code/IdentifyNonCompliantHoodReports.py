#Summary: The research team has been getting some feedback from reviewers that some of the language in our neighborhood reports
#         uses words that are not in compliance with USPAP rules. We have compiled a list of these words.
#         This script reads all 2022 existing hood reports and finds the files that have any of the offending words
#         It then exports a csv with a list of these files

#Date: 03/15/2022
#Author: Mike Leahy

import os
import pandas as pd
from docx import Document

#Define file paths
main_output_location           =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research', 'Market Analysis','Neighborhood')     
data_location                  =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects', 'Research Report Automation Project','Data', 'Neighborhood Reports Data')

#Read in list of banned words as df
banned_words_df                = pd.read_csv(os.path.join(data_location,'Compliance Information','BlacklistWords.csv')) #read in crosswalk file
bad_files                      = []
banned_words                   = []
bad_paragraphs                 = []


#Loop through the neighborhood report directory
for (dirpath, dirnames, filenames) in os.walk(main_output_location):
        
        #Skip the archived folders
        if  ('Archive' in dirpath):
            continue

        for file in filenames:

            #Skip non report documents
            if '.docx' not in file:
                continue
            
            full_path = os.path.join(dirpath,file)
            document = Document(full_path)

            for paragraph in document.paragraphs:
                if full_path in bad_files:
                    break

                for banned_word in banned_words_df['Banned Words']:
                    if ((banned_word) in paragraph.text) or  ((banned_word.lower()) in paragraph.text) or  ((banned_word).title() in paragraph.text) or  ((banned_word.upper()) in paragraph.text) :

                        bad_files.append(full_path)
                        banned_words.append(banned_word)
                        bad_paragraphs.append( paragraph.text)
                        break
                        



bad_files_df = pd.DataFrame({'Non-Compliant Files': bad_files,
                            'Offending Word':       banned_words,
                            'Offending Paragraph':  bad_paragraphs,
                            }
                           )

bad_files_df.to_csv(os.path.join(data_location, 'Compliance Information','Non Compliant Files.csv'))