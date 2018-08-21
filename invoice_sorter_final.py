from wand.display import display
from wand.image import Image
from PIL import Image as PI
import pyocr
import pyocr.builders
import io
import os
import PyPDF2
import time
import pickle
import shutil

def roundHalfUp(num):
    if num*10%10 >= 5:
        return int(num) + 1
    return int(num)

def combiner(direc, finPath=None):
    nameless = True
    count = 0
    pdfWriter = PyPDF2.PdfFileWriter()
    itemName = direc.split('/')[-1]
    print(itemName)
    name = itemName + '.pdf'
    
    if finPath is not None:
        if name in os.listdir(finPath):
            os.remove(os.path.join(finPath, name))
        final_file = open(os.path.join(finPath, name), 'wb')
        
    else:
        if name in os.listdir(os.pardir):
            os.remove(os.path.join(os.pardir, name))
        final_file = open(os.path.join(os.pardir, name), 'wb')
    
    for item in sorted(os.listdir(direc)):
        if item.endswith('pdf'):
            tempFile = open(direc + '/' + item, 'rb')
            tempReader = PyPDF2.PdfFileReader(tempFile, strict=False)
            for pageNum in range(tempReader.numPages):
                pdfWriter.addPage(tempReader.getPage(pageNum))
            pdfWriter.write(final_file)
            tempFile.close()
            
    final_file.close()
    print('merged')
        
def makePickle():
    pickle_out = open('count.pickle', 'wb')
    pickle.dump({'msc' : 0, 'msf' : 0, 'last' : roundHalfUp(time.time())}, pickle_out)
    pickle_out.close()

def cleanScanned(path):
    for remaining in os.listdir(path):
        if remaining.endswith('pdf'):
            os.remove(path + '/' + remaining)
            
def cleanPre(path):
    for direcs in os.listdir(path):
        # print(direcs)
        if os.path.isdir(path + '/' + direcs):
            # print('made')
            combiner('prepros/' + direcs)
            shutil.rmtree('prepros/' + direcs)
        if direcs.endswith('pdf'):
            os.remove('prepros/' + direcs)

def testFirst(path):
    for direcs in os.listdir(path):
        # print(direcs)
        if os.path.isdir(path + '/' + direcs):
            # print('made')
            combiner('prepros/' + direcs, 'test')
            shutil.rmtree('prepros/' + direcs)
        if direcs.endswith('pdf'):
            os.remove('prepros/' + direcs)
    
def main():
    scanPath = 'scanned'
    formalPath = 'prepros'
    test = False # turn False to properly sort
    files = os.listdir()
    files.sort()

    if not 'count.pickle' in files:
        makePickle()
        
    if not 'prepros' in files:
        os.mkdir(formalPath)
        
    pickle_in = open('count.pickle', 'rb')
    in_dict = pickle.load(pickle_in)
    pickle_in.close()
    path = None
    
    if roundHalfUp(time.time()) - in_dict['last'] >= 36000:
        in_dict['msf'] = 0
        in_dict['msc'] = 0
    in_dict['last'] = roundHalfUp(time.time())
    
    # hardcoded number to match the number of invoices expected
    count = 100
    tool = pyocr.get_available_tools()[0]
    lang = tool.get_available_languages()[0]
    
    # used to count the number of supporting documents following an invoice
    counter = 0
    s = time.ctime()
    dis = s.split()
    name = dis[1] + "_" + dis[2] + "_" + dis[4]
    
    # error log to record any misfiled documents
    log = open('error_log_' + name + '.txt', 'a') 

    for scanFileName in os.listdir(scanPath):
        
        if scanFileName.endswith('pdf'):
            try:
                pdf_collection = open(scanPath + '/' + scanFileName, 'rb')
                pdfReader = PyPDF2.PdfFileReader(pdf_collection, strict=False)
                for page in range(pdfReader.numPages):
                    print('read')
                    pageWriter = PyPDF2.PdfFileWriter()
                    with open(formalPath + '/' + 'prepros_' + str(count) + '.pdf', 'wb') as fi:
                        pageWriter.addPage(pdfReader.getPage(page))
                        pageWriter.write(fi)
                    count += 1
                path = None
                pdf_collection.close()
                
            except Exception as e:
                log.write('Error for scanning document ' + scanFileName + '\n')
                os.rename('scanned/' + scanFileName, 'misfiled/misfiled_collection_' + str(in_dict['msc']) + '.pdf')
                in_dict['msc'] += 1
    
    preprosList = os.listdir(formalPath)
    preprosList.sort()
    
    for file in preprosList:
        req_image = []
        final_text = []
        final_digits = []
        if not file.endswith('pdf'):
            continue
            
        print(file)
        image_pdf = Image(filename = formalPath + '/' + file, resolution=300)
        image_jpeg = image_pdf.convert('jpeg')
        
        w, h = image_jpeg.size
        image_jpeg.crop(w//2, 0, width=(7 * w)//8, height=h//8)
        image_jpeg.gaussian_blur(3, 2) # check the results for this; if not working, change the sigma, It appears a little big
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
            for i in range(len(temp)):
                temp[i] = temp[i].strip()
            if "Invoice" in temp:
                print("invoice")
                i = temp.index("Invoice")
                assert(i != len(temp))
                possText = temp[i+1]
                dirName = possText.strip().split()[0]
                fiName = dirName.strip() + '.pdf'
                try:
                    # print(os.listdir(formalPath))
                    if dirName not in os.listdir(formalPath):
                        os.mkdir(formalPath + '/' + dirName)
                        os.rename(formalPath + '/'+ file, formalPath + '/' + dirName + "/" + fiName)
                    else:
                        os.rename(formalPath + '/'+ file, formalPath + '/' + dirName + "/" + "_copy_" + str(counter) + fiName)
                    path = formalPath + '/' + dirName
                    print(os.listdir(path))
                    counter = 0
                    break
                
                except Exception as e:
                    log.write('Error for document ' + file + ';, likely a duplicate file. Computer says ' + e + ' \n')
                    
            else:
                if path is not None:
                    counter += 1
                    if dirName not in os.listdir(formalPath):
                        print('made')
                        os.mkdir(formalPath + '/' + dirName)
                    if counter >= 6:
                        log.write('Error for document ' + file + ';, may be incorrectly filed. \n')
                        os.rename(formalPath + '/' + file, 'misfiled/misfiled_' + str(in_dict['msf']) + '.pdf')
                    else:
                        os.rename(formalPath + '/' + file, path + "/" + 'support_doc_' + str(counter) + '.pdf')
                else:
                    log.write('Error for document misfiled/misfiled_' + str(in_dict['msf']) + '.pdf;, not successfully scanned.\n')
                    os.rename(formalPath + '/' + file, 'misfiled/misfiled_' + str(in_dict['msf']) + '.pdf')
                    counter = 0
                    in_dict['msf'] += 1
                    
    pickle.dump(in_dict, open('count.pickle', 'wb'))
    
    if not test:
        cleanPre(formalPath)
        cleanScanned(scanPath)
        
    else:
        testFirst(formalPath)
        cleanScanned(scanPath)
    
    log.close()
    
if __name__ == '__main__':
    main()
                
                        
                
            
