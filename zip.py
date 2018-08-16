#! /usr/bin/env/ python3

'''This is a learning script for zipping and unzipping .docx files.'''

import os
import shutil
import zipfile

#  Setup.

os.makedirs('unzipped')

#  Unzip document.
zip_ref = zipfile.ZipFile('essay.docx', 'r')
zip_ref.extractall('unzipped')
zip_ref.close()

def list_tree(path):
    '''Taken from: https://stackoverflow.com/
    questions/9727673/list-directory-tree-structure-in-python'''
    for root, dirs, files in os.walk(path):
        level = root.replace(path, '').count(os.sep)
        indent = ' ' * 4 * (level)
        print('{}{}/'.format(indent, os.path.basename(root)))
        subindent = ' ' * 4 * (level + 1)
        for f in files:
            print('{}{}/'.format(subindent, f)

list_tree('unzipped')

with open('unzipped/word/document.xml') as f:
    text = f.read()
print(text)

#  Pulldown.

shutil.rmtree('unzipped')  # Removes full folder.
