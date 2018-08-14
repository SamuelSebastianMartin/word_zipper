#! /usr/bin/env/ python3

'''This is a learning script for zipping and unzipping .docx files.'''

import os
import shutil
import zipfile

#  Setup

os.makedirs('unzipped')

zip_ref = zipfile.ZipFile('essay.docx', 'r')
zip_ref.extractall('unzipped')
zip_ref.close()

with open('unzipped/word/document.xml') as f:
    text = f.read()
#  Pulldown

shutil.rmtree('unzipped')  # Removes full folder.
