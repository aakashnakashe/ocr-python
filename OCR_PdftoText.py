import io, os
from PIL import Image
import pytesseract as tess
tess.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
from wand.image import Image as wi
from docx import Document


userfilelocation = str(input('Your Folder Location: '))
pdffilename = str(input('Your File Name: '))
os.chdir(userfilelocation)

userfilename = input('What should i call the file?: ')

try:
    print("\nStarted extracting the data from the pdf...")
    pdf = wi(filename = pdffilename+'.pdf', resolution = 300)
    # pdf = wi(filename = "Portion.pdf", resolution = 300)
    pdfImage = pdf.convert('jpeg')

    imageBlobs = []

    for img in pdfImage.sequence:
        imgPage = wi(image = img)
        imageBlobs.append(imgPage.make_blob('jpeg'))

    # recognized_text = []

    for imgBlob in imageBlobs:
        im = Image.open(io.BytesIO(imgBlob))
        text = tess.image_to_string(im, lang = 'eng')
    #     recognized_text.append(text)

    # print(recognized_text)
    print("\nSaving the extracted data in text format...")
    # for saving the file in text format
    with open(userfilename+'.txt', mode='w') as file:
        file.write(text)
        print("\nYour pdf to text conversion has been completed")

    print("\nDo you want to save your file in .docx format")
    answer = input("Enter 'y' for Yes and 'n' for No and then press Enter")   
    if answer == "y":
        # for saving the file in docx

        document = Document()
        myfile = open(userfilename+'.txt').read()
        p = document.add_paragraph(myfile)
        document.save(userfilename+'.docx')
        print('Your document in docx format has successfully saved')

except Exception as error:
    print('\nCaught this error: ' + repr(error))

input('Press any key to close...')