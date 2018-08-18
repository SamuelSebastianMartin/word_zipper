#! /usr/bin/env/ python3

'''This is a learning script for zipping and unzipping .docx files.'''

import os
import shutil
import zipfile

#  Setup.

if 'unzipped' in os.listdir():
    shutil.rmtree('unzipped')  # Removes any leftover folder from old crash.
os.makedirs('unzipped')

#  Unzip document.
zip_ref = zipfile.ZipFile('essay.docx', 'r')
zip_ref.extractall('unzipped')
zip_ref.close()


#with open('unzipped/word/document.xml') as f:
#    text = f.read()
#print(text)



#  Rezip document
shutil.make_archive('new_document', 'zip', 'unzipped')
if 'new_document.docx' in os.listdir():
    os.remove('new_document.docx')
os.rename('new_document.zip', 'new_document.docx')

#  Pulldown.

shutil.rmtree('unzipped')  # Removes full folder.
