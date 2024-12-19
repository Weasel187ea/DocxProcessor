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
        # print(headerData)
        # print()
        if mainHeader == "": #this section extracts header
            for header in iter_headings(doc.paragraphs):
                Data["title"] = header.text

        for run in para.runs:
            if run.bold:
                if "Section" in run.text:
                    Data["section"] = para.text.replace(u'\xa0', u' ').replace("\n"," ").replace("Section:", "")
                elif "Length" in run.text:
                    # text = 
                    Data["length"] = para.text.replace(u'\xa0', u' ').replace("\n"," ").replace("Length:", "")

                elif "Body" in para.text:
                    body = True
                headerDataDone = True # Once we reach bold words we can assume header info has been processed
                if len(headerData) >= 2:
                    Data["publisher"] = headerData[1]
                

        if (not headerDataDone) and (para.text != ""):
            headerData.append(para.text)

        if para.text == "":
            continue

        elif ("Load-Date:" in para.text) or ("End of Document" in  para.text):
            continue


        elif body == True: #we are in the body paragraph
                if (para.text != mainHeader) and (para.text != ""):
                    lineToAdd = para.text.replace("\n", " ")
                    # fullText.append(re.sub('\s+', ' ', lineToAdd)) # look this over
                    fullText.append(lineToAdd)

        
    Data["body"] = (" ".join(fullText)) 
    #########################################
    with open("finalData.csv", "a", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writerows([Data])
    #########################################
    return Data



if __name__ == "__main__":

    with open("finalData.csv", "w") as csvfile:
        fieldnames = ["title", "section", "publisher", "length", "body"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        


    onlyfiles = os.listdir()

    for file in onlyfiles:
        if file.endswith("docx"):
            getText(file)
    # for data in FileData[0]:
    #     print(data)
    #     print()

    # print(FileData[0])
    # for char in FileData[0][7]:
    #     print(char, end="")
    #     if char == ".":
    #         print()


    # for data in FileData:
    #     # print(data)
    #     for index in data:
    #         print(index)
    #         print()
    #     break
        