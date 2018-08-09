from wand.display import display
from wand.image import Image
from PIL import Image as PI
import pyocr
import pyocr.builders
import io
import os
import PyPDF2
import time

def roundZeros(num):
    numZeros = 0
    while(num > 0):
        num //= 10
        numZeros += 1
    return 10**(numZeros - 1)
    
def main():
    tool = pyocr.get_available_tools()[0]
    lang = tool.get_available_languages()[0]
    counter = 0
    s = time.ctime()
    dis = s.split()
    name = dis[1] + "_" + dis[2] + "_" + dis[4] 
    log = open('error_log_' + name + '.txt', 'w')
                       
    pdf_collection = open('concat_test.pdf', 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdf_collection, strict=False)
    formalPath = 'test_invoices/prepros'
    files = os.listdir('test_invoices')
    files.sort()
    
    if not 'prepros' in files:
        os.mkdir(formalPath)
    count = roundZeros(pdfReader.numPages)
    
    for page in range(pdfReader.numPages):
        print('read')
        pageWriter = PyPDF2.PdfFileWriter()
        with open(formalPath + '/' + 'prepros_' + str(count) + '.pdf', 'wb') as fi:
            pageWriter.addPage(pdfReader.getPage(page))
            pageWriter.write(fi)
        count += 1
    path = None
    
    preprosList = os.listdir(formalPath)
    preprosList.sort()
    for file in os.listdir(formalPath):
        req_image = []
        final_text = []
        final_digits = []
        if os.path.isdir(formalPath + '/' + file):
            continue
            
        print(file)
        image_pdf = Image(filename= formalPath + '/' + file, resolution=300)
        image_jpeg = image_pdf.convert('jpeg')
        
        w, h = image_jpeg.size
        image_jpeg.crop(w//2, 0, width=w, height=h//8)
        # display(image_jpeg)
        
        for img in image_jpeg.sequence:
            img_page = Image(image=img)
            req_image.append(img_page.make_blob('jpeg'))
            
        for img in req_image:
            txt = tool.image_to_string(
            PI.open(io.BytesIO(img)),
            lang=lang,
            builder=pyocr.builders.TextBuilder()
            )
            final_text.append(txt)
            
        for item in final_text:            
            temp = item.strip().split()
            if "Invoice" in temp:
                print("invoice")
                i = temp.index("Invoice")
                assert(i != len(temp))
                possText = temp[i+1]
                dirName = possText.strip().split()[0]
                fiName = dirName.strip() + '.pdf'
                print(fiName, dirName)
                if dirName not in os.listdir(formalPath):
                    print('maker')
                    os.mkdir(formalPath + '/' + dirName)
                os.rename(formalPath + '/'+ file, formalPath + '/' + dirName + "/" + fiName)
                path = formalPath + '/' + dirName 
                counter = 0
                break
            
            else:
                if path is not None:
                    counter += 1
                    if dirName not in os.listdir(formalPath):
                        print('made')
                        os.mkdir(formalPath + '/' + dirName)
                    os.rename(formalPath + '/' + file, path + "/" + 'support_doc_' + str(counter) + '.pdf')
                else:
                    log.write('Error for document ' + file + ';, not successfully scanned.\n')
                    counter = 0
    log.close()
if __name__ == '__main__':
    main()
                
                        
                
            