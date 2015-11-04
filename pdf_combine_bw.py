#!/usr/bin/env python2.7
#
# Takes a ZIP file from a right-click menu in explorer or a command line argument:
#  - Combines pdfs into a single file
#  - Scales all PDFs to the same size
#  - Creates one bookmark per file, provides zoom level for whole pdf
#  - Sorts files on the basis that the filename is:
#    DWG Number _  sheet number _ revision 
#  - Sort order is DWG Number ascending followed by sheet number ascending.

# TODO:
# - OCR the output (perhaps with tesseract-ocr-3.02-win32-portable.zip, http://documentup.com/virantha/pypdfocr/)
# - presently abandoned as too difficult/time consuming.
# - bring back previous functionality

__author__ = "Daniel Mulholland"
__copyright__ = "Copyright 2015, Daniel Mulholland"
__credits__ = ["Whoever wrote ghostscript, python!"]
__license__ = "GPL"
__version__ = "0.5"
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
import tempfile
import datetime

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
    
def real_start(temporary_folder, zip_file_name):
    # create a list / array
    file_list = []
    list_to_compile = []

    # walk through all folders in script folder
    for folder, subs, files in os.walk(temporary_folder):
        for filename in files:
            # only look at pdfs based on the extension
            if filename[-3:] == 'pdf':
                file_list.append([folder, filename])
    
    sortable = []
    for f in file_list:         
      splitted = f[1].split('_')
      splitted.insert(0, os.path.join(f[0],f[1]))
      sortable.append(splitted)
    
    # this ensures that e.g. sheet 4A is near sheet 4 when ordered
    # this ensures that peculiarity in drawing file name case does not impact ordering
    sortable.sort(key=lambda row: (row[1].upper() + re.sub("[^0-9]", "", row[2]).zfill(4)), reverse=False)

    final_location_and_name = zip_file_name[0:-4]+PDF_EXTENSION
    chosen_name = os.path.join(temporary_folder,os.path.basename(zip_file_name)[0:-4]+PDF_EXTENSION)

    merge_and_create_bookmarks(sortable, chosen_name, temporary_folder)

    # create ps file to allow processing of colour directives
    output_name_ps = "\"" + chosen_name[0:-3]+ 'ps'  + "\""
    ps_output_options = "-dNOPAUSE -dBATCH -sDEVICE=ps2write"    
    output_command = " ".join([GS_BINARY_PATH, ps_output_options, "-sOutputFile=" + output_name_ps, "\"" + chosen_name +  "\""])

    p = sub.Popen(output_command, stdout=sub.PIPE, stderr=sub.PIPE, cwd=temporary_folder)
    output, errors = p.communicate()
    #print output, errors
    if errors != "": print errors
    
    # convert file to pdf and process colour
    # http://superuser.com/questions/200378/converting-a-pdf-to-black-white-with-ghostscript

    ps_output_options = " ".join(["-o " + "\"" +  final_location_and_name + "\"", "-sDEVICE=pdfwrite", "\"" + os.path.join(temporary_folder,TOC_FILE) +  "\"", "-c \"/setrgbcolor{0 mul 3 1 roll 0 mul 3 1 roll 0 mul 3 1 roll 0 mul add add setgray} def\"", "-f " + output_name_ps])    
    output_command = " ".join([GS_BINARY_PATH, ps_output_options])

    p = sub.Popen(output_command, stdout=sub.PIPE, stderr=sub.PIPE, cwd=temporary_folder)
    output, errors = p.communicate()
    #print output, errors
    if errors != "": print errors
    
    # remove the ps file
    os.remove(output_name_ps)
    
def merge_and_create_bookmarks(sortable, output_name, temporary_folder):
    # create bookmarks
    f = open(os.path.join(temporary_folder,TOC_FILE), 'w')
    str_files= ""
    for index,item in enumerate(sortable):
        f.write('[/Page ' + str(index + 1) + ' /View [ /FitH ] /Title ( ' + item[1].upper() + ' Sh ' + item[2].upper() + ' ) /OUT pdfmark' + "\n")
        str_files += " " + "\"" + os.path.join(temporary_folder,item[0])  + "\"" 
    f.close()    
    
    # merge all pdfs
    #A3 = 11.7 x 16.5 in 3510x4950
    ps_output_options = "-dNOPAUSE -dBATCH -sDEVICE=pdfwrite -dPDFFitPage -r300x300 -g4950x3510"
    output_command = " ".join([GS_BINARY_PATH, ps_output_options, "-sOutputFile=" + "\"" + output_name + "\"", '-f ' + str_files])

    p = sub.Popen(output_command, stdout=sub.PIPE, stderr=sub.PIPE, cwd=temporary_folder)
    output, errors = p.communicate()
    #print output, errors
    if errors != "": print errors  

def init_start(temporary_folder, zip_file):
    # keep a list of what we have so we don't delete them
    for f in glob.glob(os.path.join(sys.argv[1],"*.zip")):
        KEEP_FILES.append(os.path.basename(f)) # the file at present
        KEEP_FILES.append(os.path.basename(f)[0:-4] + PDF_EXTENSION) # the future PDF
    
    # time the overall operation
    total_start = time.time()
    # extract and process each zip file
    
    # time the operation
    start = time.time()
    
    # extract zip file
    unzip(zip_file, temporary_folder)    
    
    # carry out merging, resizing, scaling, colour operations
    real_start(temporary_folder, zip_file)
       
    end = time.time()
    print("Elapsed time: " + "%.2f" % (end-start))

    total_end=time.time()
    print("TOTAL elapsed time: " + "%.2f" % (total_end-total_start))
       
if __name__ == "__main__":
    
    # show name of file being processed
    print('Found zip: ' + sys.argv[1])

    # make folder with date time.
    date_and_time = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    # create temporary folder
    tempfolder =  tempfile.mkdtemp(suffix=date_and_time, prefix='tmp', dir=None)
    # now begin!
    init_start(tempfolder, sys.argv[1])