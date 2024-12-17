import re
import docx
import os



def iter_headings(paragraphs):
    for paragraph in paragraphs:
        if paragraph.style.name.startswith('Heading'):
            yield paragraph

def getText(filename):
    doc = docx.Document(filename)
    fullText = [] #all text will be joined at the end
    mainHeader = "" #header
    body = False #we have reached the body paragraph, collect everything
    length = 0 # length of document
    section = 0 # section stated on the document
    Data = [] # --> [header, length, section, text]


    count = 0
    for para in doc.paragraphs:
        if mainHeader == "": #this section extracts header
            # print(f"this is the header:")
            for header in iter_headings(doc.paragraphs):
                mainHeader = header.text
                # print(mainHeader)
            # print("end of header\n")

        if para.text.startswith("Length"):
            length = para.text.replace(u'\xa0', u' ') #gets rid of \xa0
        elif para.text.startswith("Section"):
            section = para.text.replace(u'\xa0', u' ') #gets rid of \xa0
        elif para.text.startswith("Body"):
            body = True
        elif body == True: #we are in the body paragraph
                if (para.text != mainHeader) and (para.text != ""):
                    fullText.append(para.text.replace("\n", " "))

        
    Data = [mainHeader, length, section, " ".join(fullText)]
    return Data



onlyfiles = os.listdir()
FileData = []

for file in onlyfiles:
    if file != "extractText.py":
        FileData.append(getText(file))

for data in FileData:
    for index in data:
        print(index)
        print()
    break