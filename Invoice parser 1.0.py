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
from tqdm import tqdm
import time



pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\\tesseract.exe"

#Prerequesite: Ghostscript -    'pip install ghostscript' and https://www.ghostscript.com/download/gsdnld.html
#               Imagemagick - https://imagemagick.org/script/download.php
#               Tesseract-OCR - https://github.com/UB-Mannheim/tesseract/wiki

#Requires a Stable internet connection for Space OCR to work but not for pytesseract


#TO DISABLE WARNINGS
warnings.filterwarnings("ignore")


#REMOVE RESIDUAL FILES
useless_files_1 = glob.glob('D:\Documents\CSE\Python\Invoice image processing\Split PDFs\*')
for f in useless_files_1:
    os.remove(f)

useless_files_2 = glob.glob('D:\Documents\CSE\Python\Invoice image processing\Images\*')
for f in useless_files_2:
    os.remove(f)

useless_files_3 = glob.glob('D:\Documents\CSE\Python\Invoice image processing\Rotated Images\*')
for f in useless_files_3:
    os.remove(f)

useless_files_4 = glob.glob('D:\Documents\CSE\Python\Invoice image processing\Clean images\*')
for f in useless_files_4:
    os.remove(f)

    
sample_list = [f for f in listdir("D:\Documents\CSE\Python\Invoice image processing\Samples") if isfile(join("D:\Documents\CSE\Python\Invoice image processing\Samples", f))]



#
#
#PDF SPLITTING SECTION
#
#

def split_pdf_pages(root_directory, extract_to_folder):
 
 for root, dirs, files in os.walk(root_directory):
  for filename in files:
   basename, extension = os.path.splitext(filename)
   
   if extension == ".pdf":
    
    fullpath = root + "\\" + basename + extension

    
    opened_pdf = PyPDF2.PdfFileReader(open(fullpath,"rb"),False, None, True)

    
    for i in range(opened_pdf.numPages):
    # write the page to a new pdf
     output = PyPDF2.PdfFileWriter()
     output.addPage(opened_pdf.getPage(i))
     with open(extract_to_folder + "\\" + basename + "-%s.pdf" % i, "wb") as output_pdf:
      output.write(output_pdf)


# parameter variables
root_dir = r"D:\\Documents\\CSE\Python\\Invoice image processing\\Samples"
extract_to = r"D:\\Documents\\CSE\Python\\Invoice image processing\\Split PDFs"



split_pdf_pages(root_dir, extract_to)

#Creating list of files
file_list = [f for f in listdir("D:\Documents\CSE\Python\Invoice image processing\Split PDFs") if isfile(join("D:\Documents\CSE\Python\Invoice image processing\Split PDFs", f))]


# ONLY FOR TESTING
#
#
#
#print (file_list)
#
#
#




#
#PDF TO IMAGE SECTION
#
k=len(file_list)

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


#Change DPI below [DO NOT GO BELOW '300']

for i in range(k):
    F=file_list[i]
    FNAME='D:\Documents\CSE\Python\Invoice image processing\Split PDFs\\' + F
    convert_pdf(FNAME, 'D:\Documents\CSE\Python\Invoice image processing\Images',300)

#List of images
    
img_list= [f for f in listdir("D:\Documents\CSE\Python\Invoice image processing\Images") if isfile(join("D:\Documents\CSE\Python\Invoice image processing\Images", f))]


#FOR TESTING
#
#
#print(img_list)
#
#


#Checking if image is rotated or not and then rotating the image

for i in range(k):
    img_name = img_list[i]
    img_path = 'D:\Documents\CSE\Python\Invoice image processing\Images\\' + img_name
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
        rot_img_path = 'D:\Documents\CSE\Python\Invoice image processing\Rotated Images\\' + img_name
        cv2.imwrite(rot_img_path, res)
    else:
        rot_img_path = 'D:\Documents\CSE\Python\Invoice image processing\Rotated Images\\' + img_name
        shutil.copy(img_path,rot_img_path)      



#List of rotated images
rot_img_list=[f for f in listdir("D:\Documents\CSE\Python\Invoice image processing\Rotated Images") if isfile(join("D:\Documents\CSE\Python\Invoice image processing\Rotated Images", f))]


#Check for duplicate images
img_sig = []
for f in rot_img_list:
    file_dir = "D:\Documents\CSE\Python\Invoice image processing\Rotated Images\\" + f
    imgg = PythonMagick.Image(file_dir)
    sig = imgg.signature()
    if(sig not in img_sig[:]):
        img_sig.append(sig)
    else:
        os.remove(file_dir)
    
#Refreshing rot_img_list
rot_img_list=[f for f in listdir("D:\Documents\CSE\Python\Invoice image processing\Rotated Images") if isfile(join("D:\Documents\CSE\Python\Invoice image processing\Rotated Images", f))]
new_k = len(rot_img_list)

#########               OCR             #########

#Space_OCR requires internet connection.
'''

def ocr_file(filename, overlay=False, api_key='3c17757ddd88957', language='eng'):
    
    payload = {'isOverlayRequired': overlay,
               'apikey': api_key,
               'language': language,
               }
    with open(filename, 'rb') as f:
        r = requests.post('https://api.ocr.space/parse/image',
                          files={filename: f},
                          data=payload,
                          )
    return r.content.decode()


def ocr_url(url, overlay=False, api_key='3c17757ddd88957', language='eng'):
    
    payload = {'url': url,
               'isOverlayRequired': overlay,
               'apikey': api_key,
               'language': language,
               }
    r = requests.post('https://api.ocr.space/parse/image',
                      data=payload,
                      )
    return r.content.decode()

ocr_out=[]

for i in range(k):
    F_OCR = rot_img_list[i]
    FILE_OCR = 'D:\Documents\CSE\Python\Invoice image processing\Rotated Images\\' + F_OCR
    test_file = ocr_file(filename=FILE_OCR, language='eng')
    ocr_out.append(test_file)
    print(F_OCR)
    print("\n")
    print(ocr_out[i])
    print("\n\n\n")



'''
ocr_out=[]

def ocr(filename):  
    
    text = pytesseract.image_to_string(PIL.Image.open(filename))  
    return text

for i in range(new_k):
    F_OCR = rot_img_list[i]
    FILE_OCR = 'D:\Documents\CSE\Python\Invoice image processing\Rotated Images\\' + F_OCR
    output_of_ocr = ocr(FILE_OCR)
    ocr_out.append(output_of_ocr)
'''
    print("\n")
    print(F_OCR)
    print("\n")
    print(ocr_out[i])
    print("\n\n\n")
 '''


currencies = {'$' : 'USD', '₹' : 'INR', '£': 'GBP', '€': 'EUR', '¥' : 'JPY'}

#PO
po_vcode = []
po_invparty = []
po_off_vess = []
po_place_port = []
po_eta = []
po_number = []
po_date = []
po_country = []
po_etd = []
po_currency = []
po_file = []
po_shipno = []



#Invoices
#WORLD FUEL SERVICES
wfs_count = 0
wfs_cust = []
wfs_inv = []
wfs_invdate = []
wfs_file = []
wfs_amount = []
wfs_currency = []

#SHM SHIPCARE
shm_count = 0
shm_inv = []
shm_invdate = []
shm_file = []
shm_amount = []
shm_currency = []
shm_buyno = []
shm_buydate = []
shm_supref = []

#INDIAN OIL
io_count = 0
io_file = []
io_vsl = []
io_place = []
io_bill = []
io_bdate = []
io_amount = []
io_currency = []

#INMARSAT
inm_count = 0
inm_file = []
inm_inv = []
inm_date = []
inm_amount = []
inm_currency = []
inm_bperiod = []

#ENERGY SHIPPING
esh_count = 0
esh_file = []
esh_inv = []
esh_date = []
esh_amount = []
esh_currency = []

#ADANI GLOBAL
adg_count = 0
adg_file = []
adg_inv = []
adg_date = []
adg_amount = []
adg_currency = []

#MARKS MARINE
mm_count = 0
mm_file = []
mm_inv = []
mm_date = []
mm_amount = []
mm_currency = []

#SIRM UK
srm_count = 0
srm_file = []
srm_inv = []
srm_date = []
srm_amount = []
srm_currency = []
srm_cus = []

#ISS MACHINERY
iss_count = 0
iss_file = []
iss_inv = []
iss_date = []
iss_amount = []
iss_currency = []

#CATHAY PACIFIC CARGO
cpc_count = 0
cpc_file = []
cpc_inv = []
cpc_date = []
cpc_amount = []
cpc_currency = []

n = "None"

for i in range(new_k):
    info = ocr_out[i]
    length = len(info)
    if((('sci contact person' in info.lower()) == True)):
        po_file.append(rot_img_list[i])
        po_line = []
        for line in info.split('\n'):
            po_line.append(line)
        po_counter = 0
        data = []
        no_of_lines = len(po_line)
        po = re.search(r"(\d{9,10}/\d{2}/\d{2}/\d{4})", info)
        if(po == None):
            po_name = 'D:\Documents\CSE\Python\Invoice image processing\Rotated Images\\' + rot_img_list[i]
            img = cv2.imread(po_name,0)
            
            img = cv2.bitwise_not(img)  

            kernel_clean = numpy.ones((2,2),numpy.uint8)
            cleaned = cv2.erode(img, kernel_clean, iterations=1)

            kernel_line = numpy.ones((1, 5), numpy.uint8)  
            clean_lines = cv2.erode(cleaned, kernel_line, iterations=6)
            clean_lines = cv2.dilate(clean_lines, kernel_line, iterations=6)

            cleaned_img_without_lines = cleaned - clean_lines
            cleaned_img_without_lines = cv2.bitwise_not(cleaned_img_without_lines)

            clean_po = 'D:\Documents\CSE\Python\Invoice image processing\Clean images\\' +rot_img_list[i]
            cv2.imwrite(clean_po,cleaned_img_without_lines)
            po_ocr = ocr(clean_po)
            po_x = re.search(r"(\d{9,10}/\d{2}/\d{2}/\d{4})", po_ocr)
            data = po_x.group(1).split('/')
            po_number.append(data[0])
            name = ''
            name=name + data[1]+'/'+ data[2] + '/' + data [3]
            po_date.append(name)
            useless_files_4 = glob.glob('D:\Documents\CSE\Python\Invoice image processing\Clean images\*')
            for f in useless_files_4:
                os.remove(f)
        else:
            data = po.group(1).split('/')
            po_number.append(data[0])
            name = ''
            name=name + data[1]+'/'+ data[2] + '/' + data [3]
            po_date.append(name)
        

        if(('vendor code' in info.lower()) != True):
            po_vcode.append(n)
        if(('invoicing party' in info.lower()) != True):
            po_invparty.append(n)
        if(('sci office/vessel' in info.lower()) != True):
            po_off_vess.append(n)
        if(('place/port' in info.lower()) != True):
            po_place_port.append(n)
        if(('eta :' in info.lower()) != True):
            po_eta.append(n)
        if(('country' in info.lower()) != True):
            po_country.append(n)
        if((('etd :' in info.lower()) != True) and (('etd:' in info.lower()) != True)):
            po_etd.append(n)
        if(('pr no ' in info.lower()) != True):
            po_currency.append(n)
        if('danaos pr no:-' not in info.lower()):
            po_shipno.append(n)
            
        for d in range(no_of_lines):
            item = po_line[d]
            if(('vendor code' in item.lower()) == True):
                data=item.split()
                po_vcode.append(data[3])
            if(('invoicing party' in item.lower()) == True):
                data = item.split()
                po_invparty.append(data[-1])
            if(('sci office/vessel' in item.lower()) == True):
                if('country' not in item.lower()):
                    data=item[item.index(':')+2:]
                    
                else:
                    data = item[item.index(':')+2:item.lower().index('country')-1]
                po_off_vess.append(data)
            if(('place/port' in item.lower()) == True):
                data = item.split()
                flag1 = []
                portname = ''
                countryname = ''
                for kk in range(len(data)):
                    if(data[kk] == ":"):
                        flag1.append(kk)
                    elif(data[kk].lower() == "country"):
                        flag2 = kk
                    elif(data[kk].lower() == "business"):
                        flag2=kk
                if((("country" in item.lower()) !=True) and ('business' not in item.lower())):
                    flag2 = len(data)
                
                    
                for kk in range((flag1[0]+1),flag2):
                    portname=portname + data[kk]+" "
                po_place_port.append(portname)
                
                
            if(('eta :' in item.lower()) == True):
                data=item.split()
                etaflag = []
                for kk in range(len(data)):
                    if(data[kk] == ":"):
                        etaflag.append(kk)
                po_eta.append(data[etaflag[0]+1])
                
            if(('country :' in item.lower()) == True):
                if(item.count(":") == 1):
                    data=item[item.index(':')+2:]
                elif(item.count(":") == 2):
                    data = item[item.lower().index('country')+10:]
                    
                po_country.append(data)
            if((('etd :' in item.lower()) == True) or (('etd:' in item.lower()) == True)):
                data = item.split()
                po_etd.append(data[-1])
            if(('pr no ' in item.lower()) == True):
                data = item.split()
                po_currency.append(data[-1])
            if('danaos pr no:-' in item.lower()):
                raw = item.split()
                data = raw[2]
                po_shipno.append(data[4:8])
         
        po_counter+=1
    
   
    elif((('world fuel services' in info.lower()) == True) and (('bank of america' in info.lower()) == True)):
        each_line = []
        for line in info.split('\n'):
            each_line.append(line)
        line_counter = 0
        data = []
        no_of_lines = len(each_line)
        datawfs = info[info.lower().index('remit this amount')+18:]
        amount = re.search(r"\d{0,3},?\d{0,3},?\d{0,3},?\d{1,3}\.\d{2}", datawfs, re.M)
        if((amount == None) or (amount == "None")):
            wfs_amount.append("None")
        else:
            am = amount.group(0)
            wfs_amount.append(am)
        cur = re.search(r"\b[A-Z]{3}\b", datawfs, re.M)
        if((cur == None) or (cur == "None")):
            wfs_currency.append("None")
        else:
            currency = cur.group(0)
            wfs_currency.append(currency)

        if(('customer no' in info.lower()) == False):
            wfs_cust.append("None")
            wfs_inv.append("None")
            wfs_invdate.append("None")
        
        for d in range(no_of_lines):
            item = each_line[d]
            if(('customer no' in item.lower()) == True):
                raw = each_line[line_counter + 1]
                data = raw.split()
                wfs_cust.append(data[0])
                wfs_inv.append(data[1])
                wfs_invdate.append(data[2])
             
            line_counter+=1
        wfs_file.append(rot_img_list[i])
        wfs_count+=1        
        
          
    elif((('shm shipcare' in info.lower()) == True) and (('we declare that this invoice' in info.lower()) == True)):
        each_line = []
        
        for line in info.split('\n'):
            each_line.append(line)
        data = []

        if('invoice no.' not in info.lower()):
            shm_inv.append("None")
            shm_invdate.append("None")
        
        if("amount chargeable" not in info.lower()):
            shm_currency.append("None")
            shm_amount.append("None")
            
        for d in range(len(each_line)):
            item = each_line[d]
            if('invoice no.' in item.lower()):
                raw  = each_line[d+1]
                data = raw.split()
                shm_inv.append(data[-2])
                shm_invdate.append(data[-1])

            if("amount chargeable" in item.lower()):
                raw = each_line[d-1]
                data = raw.split()
                shm_currency.append(data[-2])
                shm_amount.append(data[-1])
        shm_file.append(rot_img_list[i])
        shm_count+=1

    elif(('indian oil corp' in info.lower()) and ('yours sincerely' in info.lower())):
        each_line = []
        raw = [s.start() for s in re.finditer('no./date',info.lower())]
        data = info[raw[1]+8:info.lower().index('thanking')-1]
        data = data.strip()
        for line in data.split('\n'):
            each_line.append(line)
        raw = each_line[0]
        flag_io1 = raw.find('|')
        if(flag_io1 == -1):
            flag_io1 = 999999999
        flag_io2 = re.search(r"\d", raw).start()
        flag_io=min(flag_io1,flag_io2)
        vsl_data = data[:flag_io-1]
        vsl_inf = vsl_data[:vsl_data.lower().index('at')-1]
        io_vsl.append(vsl_inf.strip())
        pl_data = vsl_data[vsl_data.lower().index('at')+3:]
        io_place.append(pl_data.strip())
        io_invinf = re.search(r"\d+/", data)
        if(io_invinf == None):
            io_bill.append("None")
        else:
            io_invinf = io_invinf.group(0)
            io_invinf = io_invinf[:-1]
            io_bill.append(io_invinf)
        inv_am = each_line[0].split()[-1]
        io_amount.append(inv_am)
        inv_dat = re.search(r"\d{2}\.\d{2}\.\d{2}" , each_line[1])
        if(inv_dat == None):
            io_bdate.append("None")
        else:
            inv_dat = inv_dat.group(0)
            io_bdate.append(inv_dat)
        cur = info[info.lower().index('amount')+6:info.lower().index('no./')]
        cur = cur.strip()
        cur_inf = re.search(r"\(\D+\)", cur)
        if(cur_inf == None):
            io_currency.append("None")
        else:
            cur_inf = cur_inf.group(0)
            cur_inf = cur_inf[1:-1]
            io_currency.append(cur_inf)
        io_count+=1
        io_file.append(rot_img_list[i])

    elif(('inmarsat' in info.lower()) and ('please remit payment including' in info.lower())):
        each_line = []
        if('detach' in info.lower()):
            info = info[:info.lower().index('detach')-1]
        for line in info.split('\n'):
            each_line.append(line)
            
        if('invoice number' not in info.lower()):
            inm_inv.append("None")
        if('invoice date' not in info.lower()):
            inm_date.append("None")
        if('billing period' not in info.lower()):
            inm_bperiod.append("None")
        if('invoice currency' not in info.lower()):
            inm_currency.append("None")
        if('charges for invoice' not in info.lower()):
            inm_amount.append("None")
        
        for d in range(len(each_line)):
            item = each_line[d]
            if('invoice number' in item.lower()):
                data = item[item.index(':')+1:].strip()
                inm_inv.append(data)
            if('invoice date' in item.lower()):
                data = item[item.index(':')+1:].strip()
                inm_date.append(data)
            if('billing period' in item.lower()):
                data = item[item.index(':')+1:].strip()
                inm_bperiod.append(data)
            if('invoice currency' in item.lower()):
                data = item[item.index(':')+1:].strip()
                inm_currency.append(data)
            if('charges for invoice' in item.lower()):
                data=item.split()
                inm_amount.append(data[4])
        inm_count+=1
        inm_file.append(rot_img_list[i])
        
    elif(('energy shipping' in info.lower()) and ('account name' in info.lower())):
        each_line = []
        for line in info.split('\n'):
            each_line.append(line)
        if('invoice no.' not in info.lower()):
            esh_inv.append("None")
        if('date' not in info.lower()):
            esh_date.append("None")
        if(' total' not in info.lower()):
            esh_amount.append("None")
            esh_currency.append("None")
            
        for d in range(len(each_line)):
            item = each_line[d]
            if(('invoice no.' in item.lower()) and ('please insert' not in item.lower())):
                data = item.split()
                esh_inv.append(data[-1])
            if('date' in item.lower()):
                data = item.split()
                esh_date.append(data[-1])
            if(' total' in item.lower()):
                data = item.split()
                esh_amount.append(data[-1])
                raw = each_line[d+1].split()
                esh_currency.append(raw[-2])
        esh_count+=1
        esh_file.append(rot_img_list[i])

    elif(('adani ' in info.lower()) and ('retail-invoice' in info.lower())):
        each_line = []
        for line in info.split('\n'):
            each_line.append(line)

        if('invoice no' not in info.lower()):
            adg_inv.append("None")
            adg_date.append("None")
        
        for d in range(len(each_line)):
            item = each_line[d]
            if('invoice no.' in item.lower()):
                data = item.split()
                adg_inv.append(data[-3])
                adg_date.append(data[-1])
            if(('amount chargeable' in item.lower())):
                data = item[item.index(':')+1:item.lower().index('only')-1]
                cur = re.search(r"\b[A-Z]{3}\b", data)
                if(cur == None):
                    adg_currency.append("None")
                else:
                    cur = cur.group(0)
                    adg_currency.append(cur)
                datax = data[data.index(cur)+4:].strip()
                amount = w2n.word_to_num(datax)
                if(amount == None):
                    adg_amount.append("None")
                    k = len(adg_amount)
                else:
                    adg_amount.append(amount)
        if(adg_amount[k-1]=="None"):
            for d in range(len(each_line)):
                itemx = each_line[d]
                if(('only' in itemx.lower()) and ('juri' not in itemx.lower())):
                    datay = itemx[itemx.index(':')+1:item.lower().index('only')-1]
                    curx = re.search(r"\b[A-Z]{3}\b", datay)
                    if(cur == None):
                        adg_currency[k-1] = n
                    else:
                        curx = curx.group(0)
                        adg_currency[k-1] = curx
                    dataz = data[data.index(curx)+4:].strip()
                    amount = w2n.word_to_num(dataz)
                    if(amount == None):
                        adg_amount[k-1] = n
                        
                    else:
                        adg_amount[k-1] = amount
                    
                    
                
        adg_count+=1
        adg_file.append(rot_img_list[i])

    elif((('marks marine' in info.lower()) or ('marksmarine' in info.lower())) and ('for the remittence' in info.lower())):
        each_line = []
        data = []
        each_line = info.split('\n')

        if('invoice #' not in info.lower()):
            mm_inv.append(n)
        if('invoice date' not in info.lower()):
            mm_date.append(n)
        if('currency' not in info.lower()):
            mm_currency.append(n)
        if('us dollar' not in info.lower()):
            mm_amount.append(n)
            
        
        for d in range(len(each_line)):
            item = each_line[d]
            if('invoice #' in item.lower()):
                data = item.split()
                mm_inv.append(data[-1])
            if('invoice date' in item.lower()):
                data = item.split()
                mm_date.append(data[-1])
            if('currency' in item.lower()):
                data = item.split()
                mm_currency.append(data[-1][-3:].upper())
            if('us dollar' in item.lower()):
                datax = item
                raw = datax[datax.lower().index('dollar')+7:datax.lower().index('only')-1].strip()
                amount = w2n.word_to_num(raw)
                mm_amount.append(amount)
        
                
        mm_count+=1
        mm_file.append(rot_img_list[i])
                        
    elif(('sirm uk' in info.lower()) and ('please quote invoice' in info.lower())):        
        if('please' in  info.lower()):
            data = info[:info.lower().index('please')]
            if('invoice no' in data.lower()):
                data1 = data[data.lower().index('invoice no'):]
                k1 = re.search(r"\b\d+\b", data1)
                if(k1 == None):
                    srm_inv.append(n)
                else:
                    srm_inv.append(k1.group(0))
            else:
                srm_inv.append(n)
            if('customer no' in data.lower()):
                data2 = data[data.lower().index('customer no'):]
                k2 = re.search(r"\b\d+\b", data2)
                if(k2 == None):
                    srm_cus.append(n)
                else:
                    srm_cus.append(k2.group(0))
            else:
                srm_cus.append(n)
            if('invoice date' in data.lower()):
                data3 = data[data.lower().index('invoice date'):]
                k3 = re.search(r"\b\d{2}/\d{2}/\d{2}\b", data3)
                if(k3 == None):
                    srm_date.append(n)
                else:
                    srm_date.append(k3.group(0))
            else:
                srm_date.append(n)
        else:
            srm_inv.append(n)
            srm_cus.append(n)
            srm_date.append(n)
        if(('due on receipt' in info.lower()) and ('amount due' in info.lower())):
            datax = info[info.lower().index('due on receipt'):info.lower().index('amount due')]
            k4 = re.search(r"[₹£$€¥]\s\d*,*\.*\d+\.\d{2}\b", datax)
            if(k4 == None):
                srm_amount.append(n)
                srm_currency.append(n)
            else:
                raw = k4.group(0)
                rawx = raw.split()
                srm_amount.append(rawx[1])
                cur = currencies[rawx[0]]
                srm_currency.append(cur)
        else:
            srm_amount.append(n)
            srm_currency.append(n)

        srm_file.append(rot_img_list[i])
        srm_count+=1

    elif(('iss machinery' in info.lower()) and ('billing invoice' in info.lower())):
        each_line = info.split('\n')

        if('invoice no' not in info.lower()):
            iss_inv.append(n)
        if('invoice date' not in info.lower()):
            iss_date.append(n)
        if('grand total amount' not in info.lower()):
            iss_amount.append(n)
            iss_currency.append(n)
            
        for d in range(len(each_line)):
            item = each_line[d]
            if('invoice no' in item.lower()):
                data = item.split()
                iss_inv.append(data[-1])
            if('invoice date' in item.lower()):
                raw = re.search(r"[A-Za-z]+\s\d{1,2},\s*\d{4}", item)
                if(raw != None):
                    iss_date.append(raw.group(0))
                else:
                    iss_date.append(n)
            if('grand total amount' in item.lower()):
                raw = re.search(r"[A-Z]{3}\s?\d{0,3},?\s?\d{0,3},?\s?\d{0,3},?\s?\d{1,3}\.\s?\d{2}", item)
                if(raw != None):
                    cur = re.search(r"[A-Z]{3}", raw.group(0))
                    if(cur != None):
                        iss_currency.append(cur.group(0))
                        amount  = raw.group(0)[raw.group(0).index(cur.group(0))+3:]
                        data = amount.replace(" ","")
                        iss_amount.append(data)
                        
                    else:
                        iss_currency.append(n)
                        iss_amount.append(n)
                else:
                    iss_currency.append(n)
                    iss_amount.append(n)

        iss_file.append(rot_img_list[i])
        iss_count+=1

    
    elif(('cathay pacific cargo' in info.lower()) and ('delivery order ' in info.lower())):
        each_line = info.split('\n')

        if('invoice number' not in info.lower()):
            cpc_inv.append(n)
        if('invoice date' not in info.lower()):
            cpc_date.append(n)
        if('d/o total' not in info.lower()):
            cpc_amount.append(n)
        if('only' in info.lower()):
            cpc_currency.append("INR")
        else:
            cpc_currency.append(n)

        for d in range(len(each_line)):
            item = each_line[d]
            if('invoice number' in item.lower()):
                data = item.split()
                cpc_inv.append(data[-1])
            if('invoice date' in item.lower()):
                date = re.search(r"\b\d{1,2}[A-Z]{3}\d{4}\b", item)
                if(date != None):
                    cpc_date.append(date.group(0))
                else:
                    cpc_date.append(n)
            if('d/o total' in item.lower()):
                amount = re.findall(r"\d+", item)[-1]
                cpc_amount.append(amount)
        cpc_file.append(rot_img_list[i])
        cpc_count+=1
        

            
        
                          
            
    
    
    
if(wfs_count>0):
    print("\nWORLD FUEL SERVICES")
    for i in range(wfs_count):
        print("\n\nWFS file:\t" , wfs_file[i] ,"\nInvoice number:\t" , wfs_inv[i] , "\nCustomer no.\t", wfs_cust[i],"\nInvoice date:\t" , wfs_invdate[i], "\nTotal amount:\t" , wfs_amount[i],"\nCurrency:\t" , wfs_currency[i])

if(shm_count>0):
    print("\nSHM SHIPCARE\n")
    for i in range(shm_count):
        print("\n\nSHM file:\t" , shm_file[i],"\nInvoice number:\t", shm_inv[i], "\nInvoice date:\t", shm_invdate[i],"\nTotal amount:\t", shm_amount[i], "\nCurrency:\t", shm_currency[i])

if(io_count>0):
    print("\nINDIAN OIL\n")
    for i in range(io_count):
        print("\n\nIO file:  \t", io_file[i], "\nInvoice number:\t", io_bill[i], "\nInvoice date:\t", io_bdate[i], "\nVessel name:\t", io_vsl[i], "\nPlace:   \t", io_place[i], "\nTotal amount:\t", io_amount[i], "\nCurrency:\t", io_currency[i])

if(inm_count>0):
    print("\nINMARSAT\n")
    for i in range(inm_count):
        print("\n\nINM FILE:\t", inm_file[i], "\nInvoice number:\t", inm_inv[i], "\nInvoice date:\t",  inm_date[i], "\nBilling Period:\t", inm_bperiod[i], "\nTotal amount:\t", inm_amount[i], "\nCurrency:\t", inm_currency[i])       

if(esh_count>0):
    print("\nENERGY SHIPPING\n")
    for i in range(esh_count):
        print("\n\nESH FILE:\t", esh_file[i], "\nInvoice number:\t" , esh_inv[i], "\nInvoice date:\t" , esh_date[i], "\nTotal amount:\t",esh_amount[i], "\nCurrency:\t", esh_currency[i])

if(adg_count>0):
    print("\nADANI GLOBAL\n")
    for i in range(adg_count):
        print("\n\nADG FILE:\t", adg_file[i], "\nInvoice number:\t", adg_inv[i], "\nInvoice date:\t", adg_date[i], "\nTotal amount:\t", adg_amount[i], "\nCurrency:\t", adg_currency[i])
        
if(mm_count>0):
    print("\nMARKS MARINE\n")
    for i in range(mm_count):
        print("\n\nMM FILE:  \t", mm_file[i], "\nInvoice number:\t", mm_inv[i] , "\nInvoice date:\t" , mm_date[i], "\nTotal amount:\t", mm_amount[i], "\nCurrency:\t", mm_currency[i])

if(srm_count>0):
    print("\nSIRM UK\n")
    for i in range(srm_count):
        print("\n\nSIRM FILE:\t", srm_file[i], "\nInvoice number:\t", srm_inv[i],"\nCustomer number:", srm_cus[i], "\nInvoice date:\t", srm_date[i], "\nTotal amount:\t", srm_amount[i], "\nCurrency:\t", srm_currency[i])

if(iss_count>0):
    print("\nISS MACHINERY\n")
    for i in range(iss_count):
        print("\nISS FILE:  \t", iss_file[i], "\nInvoice number:\t", iss_inv[i], "\nInvoice date:\t", iss_date[i], "\nTotal amount:\t", iss_amount[i], "\nCurrency:\t", iss_currency[i])

if(cpc_count>0):
    print("\nCATHAY PACIFIC CARGO\n")
    for i in range(cpc_count):
        print("\nCPC FILE:   \t", cpc_file[i], "\nInvoice number:\t", cpc_inv[i], "\nInvoice date:\t", cpc_date[i], "\nTotal amount:\t", cpc_amount[i], "\nCurrency:\t", cpc_currency[i])

for i in range(len(po_file)):
    if(po_currency[i]== "None"):
        file_name = po_file[i]
        file_n = "D:\Documents\CSE\Python\Invoice image processing\Samples\\" + file_name[:-8] + ".pdf"
        pgnum = PdfFileReader(open(file_n, "rb")).getNumPages()
        image_name = po_file[i]
        for j in range(pgnum):
            image_name = image_name[:-7] + str(j) + "-0.jpg"
            index = rot_img_list.index(image_name)
            data = ocr_out[index]
            if(('including tax in' in data.lower()) == True):
                datax = data[data.lower().index("including tax in")+17: data.lower().index("including tax in")+20]
            else:
                continue
        po_currency[i] = datax
    if(po_shipno[i] == "None"):
        for d in range(len(po_line)):
            item = po_line[d]
            if('no:-' in item.lower()):
                raw = item[item.lower().index('no:-')+4:8]
                po_shipno[i] = raw

if(len(po_file)>0):
    print("\n\n\nPO\n")
    for i in range(len(po_file)):
        print("\n\n\nPO FILE:",po_file[i],"\n\nPO number:\t",po_number[i],"\nPO date:\t",po_date[i],"\nVendor Code:\t", po_vcode[i],"\nInvoice Party:\t", po_invparty[i],"\nVessel Code:\t", po_shipno[i], "\nVessel Name:\t",po_off_vess[i],"\nPort:\t\t",po_place_port[i],"\nETA:\t\t",po_eta[i],"\nPO country:\t",po_country[i],"\nPO etd:\t\t",po_etd[i],"\nPO currency:\t",po_currency[i])

useless_files_1 = glob.glob('D:\Documents\CSE\Python\Invoice image processing\Split PDFs\*')
for f in useless_files_1:
    os.remove(f)

useless_files_2 = glob.glob('D:\Documents\CSE\Python\Invoice image processing\Images\*')
for f in useless_files_2:
    os.remove(f)

useless_files_3 = glob.glob('D:\Documents\CSE\Python\Invoice image processing\Rotated Images\*')
for f in useless_files_3:
    os.remove(f)




    
    
       
           










    

    
    
    
    
