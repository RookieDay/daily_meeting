#!/usr/bin/env python
#-*- coding:utf-8 -*-

# doc2pdf.py: python script to convert doc to pdf with bookmarks!
# Requires Office 2007 SP2
# Requires python for win32 extension

import sys, os, shutil
from win32com.client import Dispatch, constants, gencache

app = Dispatch("Word.Application")
# set UI un-visible, no warning
app.Visible = False
app.DisplayAlerts = False

# convert .doc file to .pdf file
def doc2pdf(input, output):
    #w = Dispatch("Word.Application")
    global app
    try:
        doc = app.Documents.Open(input, ReadOnly = 1)
        doc.ExportAsFixedFormat(output, constants.wdExportFormatPDF,
          Item = constants.wdExportDocumentWithMarkup, CreateBookmarks = constants.wdExportCreateHeadingBookmarks)
        return 0
    except:
        return 1
    finally:
        app.Documents.Close(constants.wdDoNotSaveChanges)

# Generate all the support we can.
def GenerateSupport():
  # enable python COM support for Word 2007
  # this is generated by: makepy.py -i "Microsoft Word 12.0 Object Library"
  gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)

# convert files in directory and sub-directories
def walk_directory(directory):
    total = 0
    for root, dirs, files in os.walk(directory):
        count = 0
        # make directory
        pdf_file_dir = os.path.join(root, "doc2pdf")
        if not "doc2pdf" in dirs:
            os.mkdir(pdf_file_dir)

        for name in files:
            if name.split('.')[-1] == "docx":
                pdf_file = os.path.join(pdf_file_dir, name.split('.')[0]+".pdf")
                doc_file = os.path.join(root, name)
                print(name.split('.')[0])
                if not doc2pdf(doc_file, pdf_file):
                    count = count + 1
        if count == 0:
            os.rmdir(pdf_file_dir)
        else:
            total = total + count
    return total

# convert a exact file and the target path is optional
# the dafault target path is in the same directory with source .doc file
def convert_one_file(filepath, target=None):
    if not os.path.isabs(filepath):
        path = os.path.abspath(filepath)
    else:
        path = filepath
    if target:
        if not os.path.isabs(target):
            target = os.path.abspath(target)
        return doc2pdf(path, target)
    pdf_file = os.path.splitext(path)[0] + ".pdf"
    return doc2pdf(path, pdf_file)

def main():
    GenerateSupport()
    # if len(sys.argv) == 2:
    #     if os.path.isdir(sys.argv[1]):
    #         return walk_directory(sys.argv[1])
    #     elif os.path.isfile(sys.argv[1]):
    #         return convert_one_file(sys.argv[1])
    # elif len(sys.argv) == 3 and os.path.isfile(sys.argv[1]):
    #     return convert_one_file(sys.argv[1], sys.argv[2])
    path = r'C:\Users\GL\Desktop\2019年中期策略会\00ok\2019-06-10 策略会 上市公司 排表\上市公司排表\上市公司排表结果'
    return walk_directory(path)

if __name__=='__main__':
    rc = main()
    print(rc)
    app.Quit(constants.wdDoNotSaveChanges)
    sys.exit(rc)