from docx2pdf import *
from pdf2image import *
from os import  chdir, remove, listdir


DPI = 500
PATH = '../poppler-0.67.0/bin'


chdir('docx')
d_docx=listdir()
for docx in d_docx:
    convert(docx ,'../pdf',keep_active=False)


chdir('../pdf')
d_pdf = listdir()
for pdf in d_pdf:
    n=0
    for pic in convert_from_path(pdf,DPI,fmt='png',poppler_path=PATH):
        filename = '../img/' + pdf[0:-4:1] + '({}).png'.format(n)
        pic.save(filename, 'png')
        n+=1



