import csv
import re
import docx
import os

# from translate import Translator
# translator= Translator(from_lang="german",to_lang="english")
# # translation = translator.translate("Guten Morgen")


def write_into_file(data): # write into file in csv format
    pass


def iter_headings(paragraphs):
    for paragraph in paragraphs:
        if paragraph.style.name.startswith('Heading'):
            yield paragraph

def getText(filename):
    doc = docx.Document(filename)
    fullText = [] #all text will be joined at the end
    mainHeader = "" #header
    body = False #we have reached the body paragraph, collect everything
    headerData = []
    headerDataDone = False
    Data = {} # --> [header, length, section, text]


    for para in doc.paragraphs:

        if mainHeader == "": #this section extracts header
            for header in iter_headings(doc.paragraphs):
                Data["title"] = header.text

        for run in para.runs:
            if run.bold:
                if "Section" in run.text:
                    Data["section"] = para.text.replace(u'\xa0', u' ').replace("\n"," ").replace("Section:", "")
                elif "Length" in run.text:
                    Data["length"] = para.text.replace(u'\xa0', u' ').replace("\n"," ").replace("Length:", "")
                elif "Byline" in run.text:
                    Data["byline"] = para.text.replace(u'\xa0', u' ').replace("\n"," ").replace("Byline:", "")
                elif "Body" in para.text:
                    body = True

                headerDataDone = True # Once we reach bold words we can assume header info has been processed
                if len(headerData) >= 1:
                    Data["publisher"] = headerData[1]
                if len(headerData) >=2:
                    Data["date"] = headerData[2]
                

        if (not headerDataDone) and (para.text != ""): #data in header is not part of body paragraph, stored differently
            headerData.append(para.text)


        if not para.text: #if line is empty, ignore
            continue

        elif ("Load-Date:" in para.text) or ("End of Document" in  para.text): #if includes either, we do not add to body text
            continue
        elif ("Body" in para.text) or ("PDF" in para.text):
            continue

        elif body == True: #we are in the body paragraph
                if (para.text != mainHeader) and (para.text != ""): #check that text is not header text nore empty
                    lineToAdd = para.text.replace("\n", " ") #remove newlines
                    fullText.append(lineToAdd) #add to fullTest list, will be joined later



        
    Data["body"] = (" ".join(fullText)) 
    #########################################  section that writes to csv file
    with open("finalData.csv", "a", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writerows([Data])
    #########################################
    return Data



if __name__ == "__main__":

    with open("finalData.csv", "w") as csvfile:
        fieldnames = ["title", "byline", "date", "section", "publisher", "length", "body"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        


    onlyfiles = os.listdir()

    for file in onlyfiles:
        if file.endswith("docx"):
            getText(file)