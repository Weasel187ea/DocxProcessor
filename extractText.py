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
    Data = {} # --> [header, length, section, text]


    for para in doc.paragraphs:
        if mainHeader == "": #this section extracts header
            for header in iter_headings(doc.paragraphs):
                mainHeader = header.text

        for run in para.runs:
            if run.bold:
                if "Section" in run.text:
                    Data["section"] = para.text.replace(u'\xa0', u' ').replace("\n"," ")
                elif "Length" in run.text:
                    Data["length"] = para.text.replace(u'\xa0', u' ').replace("\n"," ")
                elif "Body" in para.text:
                    body = True


        if para.text == "":
            continue

        # elif para.text.startswith("Copyright") or para.text.startswith("Load-Date"):
        #     continue
        elif para.text.startswith("CopyRight"):
            continue
        elif ("Load-Date:" in para.text) or ("End of Document" in  para.text):
            continue


        elif body == True: #we are in the body paragraph
                if (para.text != mainHeader) and (para.text != ""):
                    lineToAdd = para.text.replace("\n", " ")
                    # fullText.append(re.sub('\s+', ' ', lineToAdd)) # look this over
                    fullText.append(lineToAdd)
        # else:
        #     Data.append(para.text.replace(u'\xa0', u' ').replace("\n"," ")) #gets rid of \xa0

        
    Data["body"] = (" ".join(fullText))
    return Data



if __name__ == "__main__":


    onlyfiles = os.listdir()
    FileData = []

    for file in onlyfiles:
        if file.endswith("docx"):
            FileData.append(getText(file))
    # for data in FileData[0]:
    #     print(data)
    #     print()

    print(FileData[0])
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
        