from wand.display import display
from wand.image import Image
from PIL import Image as PI
import pyocr
import pyocr.builders
import pytesseract
import io
import os
import errno
import PyPDF2
import time
import pickle
import shutil
import stat
import ctypes
import signal

s = time.ctime()
dis = s.split()
name = dis[1] + "_" + dis[2] + "_" + dis[4]
log = open('error_log_' + name + '.txt', 'a')

# Does some magic to remove the read only functions, from stack overflow
def handleRemoveReadonly(func, path, exc):
  excvalue = exc[1]
  if func in (os.rmdir, os.remove) and excvalue.errno == errno.EACCES:
      os.chmod(path, stat.S_IRWXU| stat.S_IRWXG| stat.S_IRWXO) # 0777
      func(path)
  else:
      raise

def signal_handler(sig, frame):
    global log
    log.flush()
    log.close()
    sys.exit(0)

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

    # Checks if it already exists in the regular paths and erases if true    
    if finPath is not None:
        save_path = os.path.join(finPath, name)
        print("Saving to ", save_path)
        if name in os.listdir(finPath):
            os.remove(save_path)
        final_file = open(save_path, 'wb')
    else:
        save_path = os.path.join(os.pardir, name)
        print("Saving to ", save_path)
        if name in os.listdir(os.pardir):
            os.remove(save_path)
        final_file = open(save_path, 'wb')

    # Loops over sorted items, zips them up into a single file if they are pdf
    try:
        for item in sorted(os.listdir(direc)):
            if item.endswith('pdf'):
                tempFile = open(direc + '/' + item, 'rb')
                tempReader = PyPDF2.PdfFileReader(tempFile, strict=False)
                for pageNum in range(tempReader.numPages):
                    pdfWriter.addPage(tempReader.getPage(pageNum))
                pdfWriter.write(final_file)
                tempFile.close()
    except Exception as e:
        print("Error: ", e)        
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
            shutil.rmtree('prepros/' + direcs, ignore_errors=False, onerror=handleRemoveReadonly)
        if direcs.endswith('pdf'):
            os.remove('prepros/' + direcs)

def clearPrepros(path):
    for direcs in os.listdir(path):
        # print(direcs)
        if os.path.isdir(path + '/' + direcs):
            # print('made')

            shutil.rmtree('prepros/' + direcs, ignore_errors=False, onerror=handleRemoveReadonly)
        if direcs.endswith('pdf'):
            os.remove('prepros/' + direcs)

# Runs the test code to ensure this works; test mode doesn't delete files
def testFirst(path):
    for direcs in os.listdir(path):
        # print(direcs)
        if os.path.isdir(path + '/' + direcs):
            # print('made')
            shutil.rmtree('prepros/' + direcs, ignore_errors=False, onerror=handleRemoveReadonly)
        if direcs.endswith('pdf'):
            os.remove('prepros/' + direcs)


def checkForText(text: str, text_list: list) -> int:
    """Finds the text index if it exists in the text_list"""
    try:
        return text_list.index(text)
    except IndexError:
        return -1
    
def main():
    global log
    scanPath = 'scanned'
    formalPath = 'prepros'
    clearPrepros(formalPath)
    test = False # turn False to properly sort
    files = os.listdir()
    files.sort()
    hadErrors = False
    signal.signal(signal.SIGINT, signal_handler)

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
    pyocr.tesseract.TESSERACT_CMD = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
    tool = pyocr.get_available_tools()[0]
    lang = tool.get_available_languages()[0]
    
    # used to count the number of supporting documents following an invoice
    counter = 0
    
    numbers_dict = {}
    
    # error log to record any misfiled documents
    print("Reading scans...")

    for scanFileName in os.listdir(scanPath):
        # Loops over items in file path, reads and adds
        # to prepros folder if pdf
        if scanFileName.endswith('pdf'):
            try:
                pdf_collection = open(scanPath + '/' + scanFileName, 'rb')
                pdfReader = PyPDF2.PdfFileReader(pdf_collection, strict=False)
                for page in range(pdfReader.numPages):
                    # print('read')
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
                hadErrors = True
    
    preprosList = os.listdir(formalPath)
    preprosList.sort()

    print("Done scanning")
    
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
        image_jpeg.gaussian_blur(sigma=2.0, radius=5) # check the results for this; if not working, change the sigma, It appears a little big
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
            hadErrors = True
            continue
        
        for item in final_text:            
            temp = item.strip().split()
            for i in range(len(temp)):
                temp[i] = temp[i].strip()
                temp[i] = temp[i].lower()
            invoice_index = checkForText("invoice", temp)
            credit_index = checkForText("memo", temp)
            if invoice_index > 0 or credit_index > 0:
                # print("invoice")
                # if detected, saves in the final path with number as name
                # catches the error if the length is not correct
                if credit_index > 0:
                    i = credit_index
                    doc_type = "invoice"
                    doc_prefix = "cm"
                else:
                    i = invoice_index
                    doc_type = "credit memo"
                    doc_prefix = ""

                try: 
                    assert(i != len(temp))
                except Exception as e:
                    errMsg = 'Error for document ' + file + ';, ' + doc_type + ' number not read. Moved to misfiled. Computer says ' + e + ' \n'
                    print(errMsg)
                    log.write(errMsg)
                    os.rename(formalPath + '/' + file, 'misfiled/' + file)
                    counter = 0
                    in_dict['msf'] += 1
                    hadErrors = True
                    break

                possText = temp[i+1]
                dirName = possText.strip().split()[0] 
                fiName = doc_prefix + dirName + '.pdf'
                if first:
                    last_inv = dirName
                    first = False 
                if last_inv is not dirName:
                    print("New " + doc_type)
                    numbers_dict[last_inv] = numbers_dict[last_inv] + counter + 1
                    # print(numbers_dict)
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
                    hadErrors = True
                    break

            else:
                if path is not None:
                    if dirName not in os.listdir(formalPath):
                        try: 
                            # print('made')
                            os.mkdir(formalPath + '/' + dirName)
                        except Exception as e:
                            errMsg = 'Error on ' + file + '; could not make the the folder\n'
                            print(errMsg)
                            log.write(errMsg)
                            os.rename(formalPath + '/' + file, 'misfiled/' + file)
                            hadErrors = True
                            break
                    # Checks to see if we have too many pages attached to this invoice
                    if counter > 10:
                        errMsg = 'Error for document ' + file + ';, may be with the wrong invoice. Moved to misfiled \n'
                        print(errMsg)
                        log.write(errMsg)
                        try:
                            os.rename(formalPath + '/' + file, 'misfiled/' + file)
                        except Exception as e:
                            log.write("While sorting into misfiled, computer says " + e)
                            hadErrors = True
                    else:
                        os.rename(formalPath + '/' + file, path + "/" + 'support_doc_' + str(numbers_dict[dirName] + counter + 1) + '.pdf')
                else:
                    errMsg = 'Error on ' + file + '; Moved to misfiled. Invoice + PO documents probably need to be rescanned \n'
                    print(errMsg)
                    log.write(errMsg)
                    os.rename(formalPath + '/' + file, 'misfiled/' + file)
                    counter = 0
                    in_dict['msf'] += 1
                    hadErrors = True
                    break

            counter += 1
            print("Counter", counter)
            # print(numbers_dict[dirName])
            
                    
    pickle.dump(in_dict, open('count.pickle', 'wb'))
    # Erases all the leftover files in the paths
    if not test:
        cleanPre(formalPath)
        cleanScanned(scanPath)
        
    else:
        testFirst(formalPath)
        cleanScanned(scanPath)
    
    log.close()
    
    windowText = "No errors, yay!"
    if (hadErrors):
        windowText = "Errors while scanning, please see today's error log error_log_" + name + ".txt for more information"
    ctypes.windll.user32.MessageBoxW(0, windowText, "Invoice Sorter Message Box", 0)

    
if __name__ == '__main__':
    main()
    print("Finished!")
                
                        
                
            
