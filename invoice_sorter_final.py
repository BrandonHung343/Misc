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

# Combines items with the same invoice number name into a single pdf file 
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
    # Checks if it already exists, erases if it does
    else:
        if name in os.listdir(os.pardir):
            os.remove(os.path.join(os.pardir, name))
        final_file = open(os.path.join(os.pardir, name), 'wb')
    # Loops over sorted items, zips them up into a single file if they are pdf
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

# Stores useful data in a pickle file
def makePickle():
    pickle_out = open('count.pickle', 'wb')
    pickle.dump({'msc' : 0, 'msf' : 0, 'last' : roundHalfUp(time.time())}, pickle_out)
    pickle_out.close()

# Deletes all items in the named folder
def cleanScanned(path):
    for remaining in os.listdir(path):
        if remaining.endswith('pdf'):
            os.remove(path + '/' + remaining)

# Removes all the stuff in the folder we don't want, and removes directories
def cleanPre(path):
    for direcs in os.listdir(path):
        # print(direcs)
        if os.path.isdir(path + '/' + direcs):
            # print('made')
            combiner('prepros/' + direcs)
            shutil.rmtree('prepros/' + direcs)
        if direcs.endswith('pdf'):
            os.remove('prepros/' + direcs)

def clearPrepros(path):
    for direcs in os.listdir(path):
        # print(direcs)
        if os.path.isdir(path + '/' + direcs):
            # print('made')

            shutil.rmtree('prepros/' + direcs)
        if direcs.endswith('pdf'):
            os.remove('prepros/' + direcs)

# Runs the test code to ensure this works; test mode doesn't delete files
def testFirst(path):
    for direcs in os.listdir(path):
        # print(direcs)
        if os.path.isdir(path + '/' + direcs):
            # print('made')
            shutil.rmtree('prepros/' + direcs)
        if direcs.endswith('pdf'):
            os.remove('prepros/' + direcs)
    
def main():
    scanPath = 'scanned'
    formalPath = 'prepros'
    clearPrepros(formalPath)
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
    count = 1000
    tool = pyocr.get_available_tools()[0]
    lang = tool.get_available_languages()[0]
    
    # used to count the number of supporting documents following an invoice
    counter = 0
    s = time.ctime()
    dis = s.split()
    name = dis[1] + "_" + dis[2] + "_" + dis[4]

    numbers_dict = {}
    
    # error log to record any misfiled documents
    log = open('error_log_' + name + '.txt', 'a') 

    for scanFileName in os.listdir(scanPath):
        # Loops over items in file path, reads and adds
        # to prepros folder if pdf
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
                    # naming scheme necessary for python lexigraphical sort (101, 102, etc)
                path = None
                pdf_collection.close()
                
            except Exception as e:
                log.write('Error while scanning document ' + scanFileName + '\n')
                os.rename('scanned/' + scanFileName, 'misfiled/' + scanFileName)
                in_dict['msc'] += 1
    
    preprosList = os.listdir(formalPath)
    preprosList.sort()
    
    last_inv = ""
    first = True
    for file in preprosList:
        req_image = []
        final_text = []
        final_digits = []
        if not file.endswith('pdf'):
            continue
            
        print(file)
        # loops over list and performs OCR
        image_pdf = Image(filename = formalPath + '/' + file, resolution=300)
        image_jpeg = image_pdf.convert('jpeg')
        
        w, h = image_jpeg.size
        image_jpeg.crop(w//2, 0, width=(7 * w)//8, height=h//8)
        image_jpeg.gaussian_blur(3, 2) # check the results for this; if not working, change the sigma, It appears a little big
        # display(image_jpeg)
        
        # handle exceptions from the image processing
        try:
            for img in image_jpeg.sequence:
                img_page = Image(image=img)
                req_image.append(img_page.make_blob('jpeg'))
                # turns them into blobs and does pyocr
            
            for img in req_image:
                txt = tool.image_to_string(
                PI.open(io.BytesIO(img)),
                lang=lang,
                builder=pyocr.builders.TextBuilder()
                )
                final_text.append(txt)

        except Exception as e: 
            errMsg = 'Error while reading ' + file + ';, need to rescan. Moved to misfiled. Computer says ' + e + ' \n'
            print(errMsg)
            log.write(errMsg)
            os.rename(formalPath + '/' + file, 'misfiled/misfiled_' + file)
            counter = 0
            in_dict['msf'] += 1
            continue
        
        for item in final_text:            
            temp = item.strip().split()
            for i in range(len(temp)):
                temp[i] = temp[i].strip()
            if "Invoice" in temp:
                print("invoice")
                # if detected, saves in the final path with number as name
                i = temp.index("Invoice")
                # catches the error if the length is not correct
                try: 
                    assert(i != len(temp))
                except Exception as e:
                    errMsg = 'Error for document ' + file + ';, invoice number not read. Moved to misfiled. Computer says ' + e + ' \n'
                    print(errMsg)
                    log.write(errMsg)
                    os.rename(formalPath + '/' + file, 'misfiled/' + file)
                    counter = 0
                    in_dict['msf'] += 1
                    break

                possText = temp[i+1]
                dirName = possText.strip().split()[0] 
                fiName = dirName + '.pdf'
                if first:
                    last_inv = dirName
                    first = False 
                if last_inv is not dirName:
                    print("new invoice")
                    numbers_dict[last_inv] = numbers_dict[last_inv] + counter + 1;
                    print(numbers_dict)
                    counter = 0
                try:
                    # print(os.listdir(formalPath))
                    if dirName not in numbers_dict:
                        os.mkdir(formalPath + '/' + dirName)
                        os.rename(formalPath + '/'+ file, formalPath + '/' + dirName + "/" + fiName)
                        numbers_dict[dirName] = 0
                    else:
                        print("Number in dict", dirName)
                        os.rename(formalPath + '/'+ file, formalPath + '/' + dirName + "/" + str(numbers_dict[dirName]) + "_" + fiName)
                    path = formalPath + '/' + dirName
                    print(os.listdir(path))
                    last_inv = dirName
                    break
                
                except Exception as e:
                    errMsg = 'Error for document ' + file + ';, likely a duplicate file. Moved to misfiled. Computer says ' + e + ' \n'
                    print(errMsg)
                    log.write(errMsg)
                    os.rename(formalPath + '/' + file, 'misfiled/' + file)
                    counter = 0
                    in_dict['msf'] += 1
                    break

            else:
                if path is not None:
                    if dirName not in os.listdir(formalPath):
                        try: 
                            print('made')
                            os.mkdir(formalPath + '/' + dirName)
                        except Exception as e:
                            errMsg = 'Error on ' + file + '; could not make the the folder\n'
                            print(errMsg)
                            log.write(errMsg)
                            os.rename(formalPath + '/' + file, 'misfiled/' + file)
                            break
                    # 
                    if counter >= 10:
                        errMsg = 'Error for document ' + file + ';, may be with the wrong invoice. Moved to misfiled \n'
                        print(errMsg)
                        log.write(errMsg)
                        os.rename(formalPath + '/' + file, 'misfiled/' + file)
                    else:
                        os.rename(formalPath + '/' + file, path + "/" + 'support_doc_' + str(numbers_dict[dirName] + counter + 1) + '.pdf')
                else:
                    errMsg = 'Error on ' + file + '; Moved to misfiled. Invoice + PO documents probably need to be rescanned \n'
                    print(errMsg)
                    log.write(errMsg)
                    os.rename(formalPath + '/' + file, 'misfiled/' + file)
                    counter = 0
                    in_dict['msf'] += 1
                    break

            counter += 1
            print("Counter", counter)
            print(numbers_dict[dirName])
            
                    
    pickle.dump(in_dict, open('count.pickle', 'wb'))
    # Erases all the leftover files in the paths
    if not test:
        cleanPre(formalPath)
        cleanScanned(scanPath)
        
    else:
        testFirst(formalPath)
        cleanScanned(scanPath)
    
    log.close()
    
if __name__ == '__main__':
    main()
    print("Finished!")
                
                        
                
            
