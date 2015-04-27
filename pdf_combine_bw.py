#!/usr/bin/env python2.7
#
# Walks through a folder and all subfolders:
#  - Finds pdfs
#  - Combines pdfs into a single file
#  - Scales all PDFs to the same size
#  - Creates one bookmark per file, provides zoom level for whole pdf
#  - Sorts files on the basis that the filename is:
#    DWG Number _  sheet number _ revision 
#  - Sort order is DWG Number ascending followed by sheet number ascending.

# TODO:
# - OCR the output (perhaps with tesseract-ocr-3.02-win32-portable.zip, http://documentup.com/virantha/pypdfocr/)
# - presently abandoned as too difficult/time consuming.

__author__ = "Daniel Mulholland"
__copyright__ = "Copyright 2015, Daniel Mulholland"
__credits__ = ["Whoever wrote ghostscript, python!"]
__license__ = "GPL"
__version__ = "0.4"
__maintainer__ = "Daniel Mulholland"
__email__ = "dan.mulholland@gmail.com"

import sys
import time
import os
import subprocess as sub
import re
import zipfile
import glob
import time

INPUT_FOLDER = "in"
BASE_PATH = os.path.join(os.path.dirname(os.path.realpath(__file__)),INPUT_FOLDER)
GS_BINARY_PATH = r'C:\Program Files\gs\gs9.15\bin\gswin64c.exe'
PDF_EXTENSION = '.pdf'
KEEP_FILES = ['combined.pdf', 'pdf_combine_bw.py', 'pdfmerge.py', 'pdf_combine_bw.pyc', 'pdfmerge.pyc']
TOC_FILE = 'bookmarks_toc.ps'

def unzip(source_filename, dest_dir):
    # unzip everything and keep a list of the elements to process
    extracteditems = []
    
    with zipfile.ZipFile(source_filename) as zf:
        for member in zf.infolist():
            zf.extract(member, dest_dir)
            if member.filename[-1:] == '/': 
                extracteditems.append(member.filename)
    return extracteditems
    
def real_start(zip_file_name):
    # create a list / array
    file_list = []
    list_to_compile = []

    # walk through all folders in script folder
    for folder, subs, files in os.walk(BASE_PATH):
        for filename in files:
            # only look at pdfs based on the extension
            if filename[-3:] == 'pdf' and filename not in KEEP_FILES:
                file_list.append([folder, filename])
    
    sortable = []
    for f in file_list:         
      splitted = f[1].split('_')
      splitted.insert(0, os.path.join(f[0],f[1]))
      sortable.append(splitted)
    
    # this ensures that e.g. sheet 4A is near sheet 4 when ordered
    # this ensures that peculiarity in drawing file name case does not impact ordering
    sortable.sort(key=lambda row: (row[1].upper() + re.sub("[^0-9]", "", row[2]).zfill(4)), reverse=False)

    chosen_name = os.path.join(BASE_PATH,zip_file_name[0:-4]+PDF_EXTENSION)
    
    merge_and_create_bookmarks(sortable, chosen_name)

    # create ps file to allow processing of colour directives
    output_name_ps = "\"" + chosen_name[0:-3]+ 'ps'  + "\""
    ps_output_options = "-dNOPAUSE -dBATCH -sDEVICE=ps2write"    
    output_command = " ".join([GS_BINARY_PATH, ps_output_options, "-sOutputFile=" + output_name_ps, "\"" + chosen_name +  "\""])
    p = sub.Popen(output_command, stdout=sub.PIPE, stderr=sub.PIPE, cwd=BASE_PATH)
    output, errors = p.communicate()
    #print output, errors
    if errors != "": print errors
    
    # convert file to pdf and process colour
    # http://superuser.com/questions/200378/converting-a-pdf-to-black-white-with-ghostscript
    ps_output_options = " ".join(["-o " + "\"" +  chosen_name + "\"", "-sDEVICE=pdfwrite", "\"" + os.path.join(BASE_PATH,TOC_FILE) +  "\"", "-c \"/setrgbcolor{0 mul 3 1 roll 0 mul 3 1 roll 0 mul 3 1 roll 0 mul add add setgray} def\"", "-f " + output_name_ps])    
    output_command = " ".join([GS_BINARY_PATH, ps_output_options])
    p = sub.Popen(output_command, stdout=sub.PIPE, stderr=sub.PIPE, cwd=BASE_PATH)
    output, errors = p.communicate()
    #print output, errors
    if errors != "": print errors
    
    do_cleanup()
    
def merge_and_create_bookmarks(sortable, output_name):
    # create bookmarks
    f = open(os.path.join(BASE_PATH,TOC_FILE), 'w')
    str_files= ""
    for index,item in enumerate(sortable):
        f.write('[/Page ' + str(index + 1) + ' /View [ /FitH ] /Title ( ' + item[1].upper() + ' Sh ' + item[2].upper() + ' ) /OUT pdfmark' + "\n")
        str_files += " " + "\"" + os.path.join(BASE_PATH,item[0])  + "\"" 
    f.close()    
    
    # merge all pdfs
    #A3 = 11.7 x 16.5 in 3510x4950
    ps_output_options = "-dNOPAUSE -dBATCH -sDEVICE=pdfwrite -dPDFFitPage -r300x300 -g4950x3510"
    output_command = " ".join([GS_BINARY_PATH, ps_output_options, "-sOutputFile=" + "\"" + output_name + "\"", '-f ' + str_files])

    p = sub.Popen(output_command, stdout=sub.PIPE, stderr=sub.PIPE, cwd=BASE_PATH)
    output, errors = p.communicate()
    #print output, errors
    if errors != "": print errors

def do_cleanup():
    # TODO: this should just go through sortable and add files in cleanup
    for root, dirs, files in os.walk(BASE_PATH, topdown=False):
        for name in files:        
            if os.path.basename(name) not in KEEP_FILES:
              # print("Would remove: " + name)
              os.remove(os.path.join(root, name))
        for name in dirs:        
            os.rmdir(os.path.join(root, name))

def init_start():
    # keep a list of what we have so we don't delete them
    for f in glob.glob(os.path.join(BASE_PATH,"*.zip")):
        KEEP_FILES.append(os.path.basename(f)) # the file at present
        KEEP_FILES.append(os.path.basename(f)[0:-4] + PDF_EXTENSION) # the future PDF
    
    # time the overall operation
    total_start = time.time()
    # extract and process each zip file
    for file in glob.glob(os.path.join(BASE_PATH,"*.zip")): 
        # show user what is being worked on
        print('Found zip: ' + file)
        # time the operation
        start = time.time()
        
        # extract zip file
        unzip(file, BASE_PATH)    
        # carry out merging, resizing, scaling, colour operations
        real_start(file)
       
        end = time.time()
        print("Elapsed time: " + "%.2f" % (end-start))

    total_end=time.time()
    print("TOTAL elapsed time: " + "%.2f" % (total_end-total_start))
    
        
if __name__ == "__main__":
    init_start()
