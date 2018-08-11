import PyPDF2
import os

count = 0

pdfWriter = PyPDF2.PdfFileWriter()
final_file = open('concat_test.pdf', 'wb')

for item in os.listdir('test_invoices'):
    
    if item.endswith('pdf'):
        tempFile = open('test_invoices/' + item, 'rb')
        tempReader = PyPDF2.PdfFileReader(tempFile, strict=False)
        for pageNum in range(tempReader.numPages):
            pdfWriter.addPage(tempReader.getPage(pageNum))
        pdfWriter.write(final_file)
        tempFile.close()
        
final_file.close()
        