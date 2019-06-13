import os, PyPDF2, re
from PyPDF2 import PdfFileReader
from os import listdir
import os.path
from os.path import isfile, join
import warnings
import tempfile
import PIL.Image
from wand.image import Image
from wand.color import Color
import requests
import glob
import skimage
import pytesseract
import cv2
import numpy
import string
import shutil
from word2number import w2n
import PythonMagick
from openpyxl import Workbook
import base64
import openpyxl

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\\tesseract.exe"

warnings.filterwarnings("ignore")



useless_2 = glob.glob('cache/*')
for f in useless_2:
    os.remove(f)
    
useless_3 = glob.glob('img/*')
for f in useless_3:
    os.remove(f)
useless_4 = glob.glob('spdf/*')
for f in useless_4:
    os.remove(f)

useless_5 = glob.glob('_____EDIT THIS______(JPG)/*')
for f in useless_5:
    os.remove(f)

useless_6 = glob.glob('rimg/*')
for f in useless_6:
    os.remove(f)

def ocr(filename):  
    
    text = pytesseract.image_to_string(PIL.Image.open(filename))  
    return text

def split_pdf_pages(root_directory, extract_to_folder):
 
 for root, dirs, files in os.walk(root_directory):
  for filename in files:
   basename, extension = os.path.splitext(filename)
   
   if extension == ".pdf":
    
    fullpath = root + "\\" + basename + extension

    
    opened_pdf = PyPDF2.PdfFileReader(open(fullpath,"rb"),False, None, True)

    
    for i in range(opened_pdf.numPages):
    
     output = PyPDF2.PdfFileWriter()
     output.addPage(opened_pdf.getPage(i))
     with open(extract_to_folder + "\\" + basename + "-%s.pdf" % i, "wb") as output_pdf:
      output.write(output_pdf)


# parameter variables
root_dir = r"______SAMPLES______(PDF or JPG)"
extract_to = r"spdf"


def convert_pdf(filename, output_path, resolution):

        all_pages = Image(filename=filename, resolution=resolution)
        for i, page in enumerate(all_pages.sequence):
            with Image(page) as img:
                img.format = 'jpg'
                img.background_color = Color('white')
                img.alpha_channel = 'remove'

                image_filename = os.path.splitext(os.path.basename(filename))[0]
                image_filename = '{}-{}.jpg'.format(image_filename, i)
                image_filename = os.path.join(output_path, image_filename)

                img.save(filename=image_filename)

def img_rot(x, name):
    for i in range(x):
        img_name = name[i]
        img_path = 'img\\' + img_name
        im = skimage.io.imread(img_path)
        newdata = pytesseract.image_to_osd(im, nice=1)
        rotation=(re.search('(?<=Rotate: )\d+', newdata).group(0))
        rotation = int(rotation)
        if(rotation !=0):
            unrotated_img = img_path
            imgx = cv2.imread(unrotated_img) 
            (rows, cols) = imgx.shape[:2]  
            M = cv2.getRotationMatrix2D((cols / 2, rows / 2), rotation, 1) 
            res = cv2.warpAffine(imgx, M, (cols, rows))
            rot_img_path = '_____EDIT THIS______(JPG)\\' + img_name
            cv2.imwrite(rot_img_path, res)
        else:
            rot_img_path = '_____EDIT THIS______(JPG)\\' + img_name
            shutil.copy(img_path,rot_img_path)

def max_v(sheetx):
    a=0
    for row in sheetx.iter_rows(min_row=6,min_col=2, max_col=2, max_row=1000006):
        for cell in row:
            if(cell.value !=None):
                a+=1
            else:
                return a

def encoding(rx,ex):
    path_of_image = '____EDITED____(JPG)\\' + ex
    with open(path_of_image, "rb") as k:
        #kk = base64.b64encode(k.read())
        #d = len(kk)
        #kkk = str(kk)
               
        wb = openpyxl.load_workbook('Templates.xlsx')
        sheet = wb.get_sheet_by_name('sheet1')
        #code = max_v(sheet)+1
        #file.write(kkk)
        #wb.save('Templates.xlsx')
        rimg_path = '_____EDIT THIS______(JPG)\\' + rx
        edited_path = '____EDITED____(JPG)\\' + ex
        template = cv2.imread(edited_path,1)
        inv = cv2.imread(rimg_path,1)
        template = cv2.bitwise_not(template)
        print("\nEnter the vendor name below:")
        file_c = input()
        file_code = 'Template storage\\' + file_c + ' ts.jpg'
        cv2.imwrite(file_code,template)
        ht = inv.shape[0]
        wt = inv.shape[1]
        d = (wt,ht)
        
        template_resize = cv2.resize(template, d, interpolation = cv2.INTER_AREA)
        
        out = cv2.add(template_resize,inv)
        cv2.imwrite('cache\\output.jpg', out)
        ocr_out = ocr('cache\output.jpg')
        print("\n")
        un_list = ["'",'"', ':', ',', ';', '?', '/','<','.','>','|','[',']','{','}','(',')','*','^','!','&','%','$','#','@']
        ocr_lines = ocr_out.split('\n')
        for i in range(len(ocr_lines)):
            line = ocr_lines[i].strip()
            if(len(line)>0):
                if(line[0] in un_list):
                    line = line.replace(line[0],"",1).strip()
            ocr_lines[i] = line
            
        for d in range(len(ocr_lines)):
            print("[", d+1, "]", ".......", ocr_lines[d])
        print("\n\t\t\tPress ENTER to exit...")
        uv = input()


                          



pictures = glob.glob('____EDITED____(JPG)\\*.jpg')
invoices = glob.glob('______SAMPLES______(PDF or JPG)\\*.jpg')
pdfs = glob.glob('______SAMPLES______(PDF or JPG)\\*.pdf')

#file = open('b64 encoded templates.txt', 'a')

if(len(pictures) == 0):
    if((len(pdfs) == 1) and (len(invoices) == 0)):
        split_pdf_pages(root_dir, extract_to)
        pdf_list = [f for f in listdir("spdf") if isfile(join("spdf", f))]
        k=len(pdf_list)        

        for i in range(k):
            F=pdf_list[i]
            FNAME='spdf\\' + F
            convert_pdf(FNAME, 'img',300)

        img_list = [f for f in listdir("img") if isfile(join("img", f))]
        img_rot(k,img_list)
        rot_img_list =  [f for f in listdir("_____EDIT THIS______(JPG)") if isfile(join("_____EDIT THIS______(JPG)", f))]
        print("\nEDIT THE IMAGE IN THE FOLDER '_____EDIT THIS______' AND SAVE IT IN '____EDITED____'\nFor more instructions refer readme.txt\n\nIF YOU HAVE COMPLETED EDITING TYPE '1' AND PRESS ENTER: ")
        u1 = 0
        while(u1!=1):
            u1 = int(input())
        edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
        if(len(edited_list)>1):
            print("\nPROGRAM DOES NOT SUPPORT MULTIPLE PAGE INVOICES\nRemove extra files in '____EDITED____' folder and run the program again\nPress enter to exit")
            u2 = input()
        elif(len(edited_list)==0):
            print("\nNo image found in ____EDITED____ folder\nPress enter to exit")
            u3 = input()
        else:
            edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
            ee  = edited_list[0]
            encoding(rot_img_list[0],ee)
    elif((len(pdfs)==0) and (len(invoices) == 1)):
        sample_list = [f for f in listdir("______SAMPLES______(PDF or JPG)") if isfile(join("______SAMPLES______(PDF or JPG)", f))]
        invoice_path = '______SAMPLES______(PDF or JPG)\\' + sample_list[0]      
        raw_img_path = 'img\\' + sample_list[0]
        shutil.copy(invoice_path,raw_img_path)
        img_list = [f for f in listdir("img") if isfile(join("img", f))]
        img_rot(len(img_list),img_list)
        rot_img_list =  [f for f in listdir("_____EDIT THIS______(JPG)") if isfile(join("_____EDIT THIS______(JPG)", f))]
        print("\nEDIT THE IMAGE IN THE FOLDER '_____EDIT THIS______' AND SAVE IT IN '____EDITED____'\nFor more instructions refer readme.txt\n\nIF YOU HAVE COMPLETED EDITING TYPE '1' AND PRESS ENTER: ")
        u1 = 0
        while(u1!=1):
            u1 = int(input())
        edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
        if(len(edited_list)>1):
            print("\nPROGRAM DOES NOT SUPPORT MULTIPLE PAGE INVOICES\nRemove extra files in '____EDITED____' folder and run the program again\nPress enter to exit")
            u2 = input()
        elif(len(edited_list)==0):
            print("\nNo image found in ____EDITED____ folder\nPress enter to exit")
            u3 = input()
        else:
            edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
            ee  = edited_list[0]
            encoding(rot_img_list[0],ee)
    elif((len(pdfs)==1) and (len(invoices)==1)):
        print("Would you like to use the PDF or the IMAGE?\n[1 for pdf, 0 for img]")
        u5 = int(input())
        if(u5==1):
            split_pdf_pages(root_dir, extract_to)
            pdf_list = [f for f in listdir("spdf") if isfile(join("spdf", f))]
            k=len(pdf_list)        

            for i in range(k):
                F=pdf_list[i]
                FNAME='spdf\\' + F
                convert_pdf(FNAME, 'img',300)

            img_list = [f for f in listdir("img") if isfile(join("img", f))]
            img_rot(k,img_list)
            rot_img_list =  [f for f in listdir("_____EDIT THIS______(JPG)") if isfile(join("_____EDIT THIS______(JPG)", f))]
            print("\nEDIT THE IMAGE IN THE FOLDER '_____EDIT THIS______' AND SAVE IT IN '____EDITED____'\nFor more instructions refer readme.txt\n\nIF YOU HAVE COMPLETED EDITING TYPE '1' AND PRESS ENTER: ")
            u1 = 0
            while(u1!=1):
                u1 = int(input())
            edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
            if(len(edited_list)>1):
                print("\nPROGRAM DOES NOT SUPPORT MULTIPLE PAGE INVOICES\nRemove extra files in '____EDITED____' folder and run the program again\nPress enter to exit")
                u2 = input()
            elif(len(edited_list)==0):
                print("\nNo image found in ____EDITED____ folder\nPress enter to exit")
                u3 = input()
            else:
                edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
                ee  = edited_list[0]
                encoding(rot_img_list[0],ee)
        elif(u5==0):
            sample_list = [f for f in listdir("______SAMPLES______(PDF or JPG)") if isfile(join("______SAMPLES______(PDF or JPG)", f))]
            invoice_path = '______SAMPLES______(PDF or JPG)\\' + sample_list[0]      
            raw_img_path = 'img\\' + sample_list[0]
            shutil.copy(invoice_path,raw_img_path)
            img_list = [f for f in listdir("img") if isfile(join("img", f))]
            img_rot(len(img_list),img_list)
            rot_img_list =  [f for f in listdir("_____EDIT THIS______(JPG)") if isfile(join("_____EDIT THIS______(JPG)", f))]
            print("\nEDIT THE IMAGE IN THE FOLDER '_____EDIT THIS______' AND SAVE IT IN '____EDITED____'\nFor more instructions refer readme.txt\n\nIF YOU HAVE COMPLETED EDITING TYPE '1' AND PRESS ENTER: ")
            u1 = 0
            while(u1!=1):
                u1 = int(input())
            edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
            if(len(edited_list)>1):
                print("\nPROGRAM DOES NOT SUPPORT MULTIPLE PAGE INVOICES\nRemove extra files in '____EDITED____' folder and run the program again\nPress enter to exit")
                u2 = input()
            elif(len(edited_list)==0):
                print("\nNo image found in ____EDITED____ folder\nPress enter to exit")
                u3 = input()
            else:
                edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
                ee  = edited_list[0]
                encoding(rot_img_list[0],ee)

        else:
            print("Invalid input. Press Enter to exit")
            uv = input()
     
            
           
           
       
       
elif(len(pictures) == 1):
    print("\nTHERE APPEARS TO BE '1' IMAGE IN YOUR ____EDITED____ FOLDER. WOULD YOU LIKE TO PROCEED?\n['1' for yes. '0' for no]")
    u4=int(input())
    edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
    if(len(edited_list)>1):
        print("\nPROGRAM DOES NOT SUPPORT MULTIPLE PAGE INVOICES\nRemove extra files in '____EDITED____' folder and run the program again\nPress enter to exit")
        u2 = input()
    elif(len(edited_list)==0):
        print("\nNo image found in ____EDITED____ folder\nPress enter to exit\n")
        u3 = input()
    if(u4==1):
        if(len(edited_list)>1):
            print("\nPROGRAM DOES NOT SUPPORT MULTIPLE PAGE INVOICES\nRemove extra files in '____EDITED____' folder and run the program again\nPress enter to exit")
            u2 = input()
        elif(len(edited_list)==0):
            print("\nNo image found in ____EDITED____ folder\nPress enter to exit\n")
            u3 = input()
        sample_list = [f for f in listdir("______SAMPLES______(PDF or JPG)") if isfile(join("______SAMPLES______(PDF or JPG)", f))]
        invoice_path = '______SAMPLES______(PDF or JPG)\\' + sample_list[0]      
        raw_img_path = 'img\\' + sample_list[0]
        shutil.copy(invoice_path,raw_img_path)
        img_list = [f for f in listdir("img") if isfile(join("img", f))]
        img_rot(len(img_list),img_list)
        rot_img_list =  [f for f in listdir("_____EDIT THIS______(JPG)") if isfile(join("_____EDIT THIS______(JPG)", f))]
        edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
        ee  = edited_list[0]
        encoding(rot_img_list[0],ee)
    else:
        useless = glob.glob('____EDITED____(JPG)\*')
        for f in useless:
            os.remove(f)

        if((len(pdfs) == 1) and (len(invoices) == 0)):
            split_pdf_pages(root_dir, extract_to)
            pdf_list = [f for f in listdir("spdf") if isfile(join("spdf", f))]
            k=len(pdf_list)        

            for i in range(k):
                F=pdf_list[i]
                FNAME='spdf\\' + F
                convert_pdf(FNAME, 'img',300)

            img_list = [f for f in listdir("img") if isfile(join("img", f))]
            img_rot(k,img_list)
            rot_img_list =  [f for f in listdir("_____EDIT THIS______(JPG)") if isfile(join("_____EDIT THIS______(JPG)", f))]
            print("\nEDIT THE IMAGE IN THE FOLDER '_____EDIT THIS______' AND SAVE IT IN '____EDITED____'\nFor more instructions refer readme.txt\n\nIF YOU HAVE COMPLETED EDITING TYPE '1' AND PRESS ENTER: ")
            u1 = 0
            while(u1!=1):
                u1 = int(input())
            edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
            if(len(edited_list)>1):
                print("\nPROGRAM DOES NOT SUPPORT MULTIPLE PAGE INVOICES\nRemove extra files in '____EDITED____' folder and run the program again\nPress enter to exit")
                u2 = input()
            elif(len(edited_list)==0):
                print("\nNo image found in ____EDITED____ folder\nPress enter to exit")
                u3 = input()
            else:
                edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
                ee  = edited_list[0]
                encoding(rot_img_list[0],ee)
        elif((len(pdfs)==0) and (len(invoices) == 1)):
            sample_list = [f for f in listdir("______SAMPLES______(PDF or JPG)") if isfile(join("______SAMPLES______(PDF or JPG)", f))]
            invoice_path = '______SAMPLES______(PDF or JPG)\\' + sample_list[0]      
            raw_img_path = 'img\\' + sample_list[0]
            shutil.copy(invoice_path,raw_img_path)
            img_list = [f for f in listdir("img") if isfile(join("img", f))]
            img_rot(len(img_list),img_list)
            rot_img_list =  [f for f in listdir("_____EDIT THIS______(JPG)") if isfile(join("_____EDIT THIS______(JPG)", f))]
            print("\nEDIT THE IMAGE IN THE FOLDER '_____EDIT THIS______' AND SAVE IT IN '____EDITED____'\nFor more instructions refer readme.txt\n\nIF YOU HAVE COMPLETED EDITING TYPE '1' AND PRESS ENTER: ")
            u1 = 0
            while(u1!=1):
                u1 = int(input())
            edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
            if(len(edited_list)>1):
                print("\nPROGRAM DOES NOT SUPPORT MULTIPLE PAGE INVOICES\nRemove extra files in '____EDITED____' folder and run the program again\nPress enter to exit")
                u2 = input()
            elif(len(edited_list)==0):
                print("\nNo image found in ____EDITED____ folder\nPress enter to exit")
                u3 = input()
            else:
                edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
                ee  = edited_list[0]
                encoding(rot_img_list[0],ee)
        elif((len(pdfs)==1) and (len(invoices)==1)):
            print("Would you like to use the PDF or the IMAGE?\n[1 for pdf, 0 for img]")
            u5 = int(input())
            if(u5==1):
                split_pdf_pages(root_dir, extract_to)
                pdf_list = [f for f in listdir("spdf") if isfile(join("spdf", f))]
                k=len(pdf_list)        

                for i in range(k):
                    F=pdf_list[i]
                    FNAME='spdf\\' + F
                    convert_pdf(FNAME, 'img',300)

                img_list = [f for f in listdir("img") if isfile(join("img", f))]
                img_rot(k,img_list)
                rot_img_list =  [f for f in listdir("_____EDIT THIS______(JPG)") if isfile(join("_____EDIT THIS______(JPG)", f))]
                print("\nEDIT THE IMAGE IN THE FOLDER '_____EDIT THIS______' AND SAVE IT IN '____EDITED____'\nFor more instructions refer readme.txt\n\nIF YOU HAVE COMPLETED EDITING TYPE '1' AND PRESS ENTER: ")
                u1 = 0
                while(u1!=1):
                    u1 = int(input())
                edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
                if(len(edited_list)>1):
                    print("\nPROGRAM DOES NOT SUPPORT MULTIPLE PAGE INVOICES\nRemove extra files in '____EDITED____' folder and run the program again\nPress enter to exit")
                    u2 = input()
                elif(len(edited_list)==0):
                    print("\nNo image found in ____EDITED____ folder\nPress enter to exit")
                    u3 = input()
                else:
                    edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
                    ee  = edited_list[0]
                    encoding(rot_img_list[0],ee)
            elif(u5==0):
                sample_list = [f for f in listdir("______SAMPLES______(PDF or JPG)") if isfile(join("______SAMPLES______(PDF or JPG)", f))]
                invoice_path = '______SAMPLES______(PDF or JPG)\\' + sample_list[0]      
                raw_img_path = 'img\\' + sample_list[0]
                shutil.copy(invoice_path,raw_img_path)
                img_list = [f for f in listdir("img") if isfile(join("img", f))]
                img_rot(len(img_list),img_list)
                rot_img_list =  [f for f in listdir("_____EDIT THIS______(JPG)") if isfile(join("_____EDIT THIS______(JPG)", f))]
                print("\nEDIT THE IMAGE IN THE FOLDER '_____EDIT THIS______' AND SAVE IT IN '____EDITED____'\nFor more instructions refer readme.txt\n\nIF YOU HAVE COMPLETED EDITING TYPE '1' AND PRESS ENTER: ")
                u1 = 0
                while(u1!=1):
                    u1 = int(input())
                edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
                if(len(edited_list)>1):
                    print("\nPROGRAM DOES NOT SUPPORT MULTIPLE PAGE INVOICES\nRemove extra files in '____EDITED____' folder and run the program again\nPress enter to exit")
                    u2 = input()
                elif(len(edited_list)==0):
                    print("\nNo image found in ____EDITED____ folder\nPress enter to exit")
                    u3 = input()
                else:
                    edited_list = [f for f in listdir("____EDITED____(JPG)") if isfile(join("____EDITED____(JPG)", f))]
                    ee  = edited_list[0]
                    encoding(rot_img_list[0],ee)

            else:
                print("Invalid input. Press Enter to exit")
                uv = input()
        elif((len(pdfs)>1) or (len(invoices)>1)):
            print("Too many files in samples folder. Press enter to exit")
            uv = input()
        elif((len(pdfs)==0) and (len(invoices) == 0)):
            print("No files with 'pdf' or 'jpg' extension in samples folder. Press enter to exit")
            uv = input()
           
elif(len(pictures)>1):
    print("\nERROR TOO MANY FILES IN '____EDITED____' FOLDER. Press enter to exit\n")
    uu = input()
    

#useless_1 = glob.glob('_____EDIT THIS______(JPG)/*')
#for f in useless_1:
#    os.remove(f)

useless_2 = glob.glob('cache/*')
for f in useless_2:
    os.remove(f)
    
useless_3 = glob.glob('img/*')
for f in useless_3:
    os.remove(f)
useless_4 = glob.glob('spdf/*')
for f in useless_4:
    os.remove(f)

#file.close()
