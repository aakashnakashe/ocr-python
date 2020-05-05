import io, os
from PIL import Image
import pytesseract as tess
from wand.image import Image as wi
tess.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
from docx import Document
import time

userfilelocation = str(input('Your Folder Location: '))
useroutputlocation = str(input('Where do you want to store your Output: '))
os.chdir(userfilelocation)
print("\nYour PDF has been converting into Images..")
time.sleep(3)
print("\nYou data from the PDF had been extracted...")
time.sleep(2)
print("\nDo you want to save this file in docx format")
answer = input('\nPress y for Yes and n for No: ').lower()

try:
    for fileName in os.listdir(userfilelocation):
        if fileName.endswith('.pdf'):
            inputPath = os.path.join(userfilelocation, fileName)
            pdf = wi(filename = fileName+'', resolution = 300)
            pdfImage = pdf.convert('jpeg')

            imageBlobs = []
            pagecount = 1

            for img in pdfImage.sequence:
                imgPage = wi(image = img)
                imageBlobs.append(imgPage.make_blob('jpeg'))

                for imgBlob in imageBlobs:
                    im = Image.open(io.BytesIO(imgBlob))
                    text = tess.image_to_string(im, lang = 'eng')
                    
                    fullTempPath = os.path.join(useroutputlocation, 'pdf_' + fileName + "_" + str(pagecount) + ".txt") 
                    # print(result)
                    # saving the text for every image in a separate .txt file 
                    file1 = open(fullTempPath, "w") 
                    file1.write(text) 
                    file1.close()
                    pagecount = pagecount + 1
                    print("\nYour data in text format has been saved succesfully")

                    # for saving the file in docx
                if answer == "y":            
                    document = Document()
                    myfile = open(fullTempPath+'').read()
                    # myfile = open('pdf_'+ fileName + "_" + str(pagecount) + ".txt").read()
                    p = document.add_paragraph(myfile)
                    document.save('pdf_'+ fileName + "_" + str(pagecount) + ".docx")
                    print("\nYour word file has been saved successfully")
                else:        
                    print("\nYour data is extracted from PDF")
except Exception as error:
    print('\nCaught this error: ' + repr(error))

time.sleep(2)
input('\nPress Enter to close...')