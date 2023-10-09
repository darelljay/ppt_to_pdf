import os.path
import os
from os import path
import aspose.slides as slides
import shutil
path_dir = 'C:/Users/User/Desktop/ppt_to_pdf/ppt'
pptFile_name = ''


file_list = os.listdir(path_dir)
while pptFile_name == '':
    pptFile_name = input('Enter the name of the ppt file your trying to convert (ex: 123.pptx): ')
    if file_list.__contains__(pptFile_name)==False or pptFile_name.endswith('.pptx') == False:
        print('Please check if the ppt file you typed in is in the ppt folder if it is please check if it is a ppt file')
        pptFile_name = ''
        continue   
    else:
        pdfName = input('Please enter the name of the pdf file: ')
        pres = slides.Presentation(path_dir+'/'+pptFile_name)
        pres.save(pdfName+".pdf", slides.export.SaveFormat.PDF) 
        shutil.move('C:/Users/User/Desktop/ppt_to_pdf/'+pdfName+'.pdf','C:/Users/User/Desktop/ppt_to_pdf/pdf')

# print(len(file_list))
# # print(str(path.isfile('ppt_to_pdf1.py')))


# if len(file_list) != 0:
#     print('hello')