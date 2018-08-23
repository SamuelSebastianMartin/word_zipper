#! /usr/bin/env/ python3

'''This is a learning script for zipping and unzipping .docx files.'''

import os
import shutil
import xml.etree.ElementTree as ET
import zipfile


def setup():
    '''Cleans up old attempts, and makes empty dir to unzip into.'''
    if 'unzipped' in os.listdir():
        shutil.rmtree('unzipped')
    os.makedirs('unzipped')


def doc_unzip():
    '''unzips a file called "essay.docx".'''
    zip_ref = zipfile.ZipFile('essay.docx', 'r')
    zip_ref.extractall('unzipped')
    zip_ref.close()


def parse():
    tree = ET.parse('unzipped/word/document.xml')
    root = tree.getroot()
    print('root.tag = ', root.tag)
    print('root.attrib = ', root.attrib)
    return tree


def experiment_zone(tree):
    '''This function is for playing with to see what is possible.'''
    os.system('echo " " >> unzipped/word/document.xml')
    #with open('unzipped/word/document.xml') as f:
    #    text = f.read()
    #    text.remove(' xml:space="preserve"')


def rezip():
    '''Takes any changes, and creates a new .docx file.'''
    shutil.make_archive('new_document', 'zip', 'unzipped')
    if 'new_document.docx' in os.listdir():
        os.remove('new_document.docx')
    os.rename('new_document.zip', 'new_document.docx')

def pulldown():
    '''Clears up any dirs and files created during this process.'''
#    shutil.rmtree('unzipped')  # Removes full folder.


def main():
    setup()
    doc_unzip()
    tree = parse()
    experiment_zone(tree)
    rezip()
    pulldown()


if __name__ == '__main__':
    main()
