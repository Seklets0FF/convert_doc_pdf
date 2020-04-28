# -*- coding: utf-8 -*-
import os
import comtypes.client
import glob

wdFormatPDF = 17

list_files_doc = []
list_files_docx = []

def start():
    for f in glob.glob('*.doc'):
        list_files_doc.append(f)

    for f in list_files_doc:
        convert(4, f)

    for f in glob.glob('*.docx'):
        list_files_docx.append(f)

    for f in list_files_docx:
        convert(5, f)


def convert(fformat, f):
    print(f)
    new_name = f[:len(f) - fformat] + '.pdf'
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(os.path.abspath(f))
    doc.SaveAs(os.path.abspath(new_name), FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


if __name__ == '__main__':
    start()
