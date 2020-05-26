import docx2txt
import nltk.corpus
from nltk import Tree
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from io import StringIO
import re
from datetime import datetime
from dateutil import relativedelta

def pdf_Reader(TotalFilename):
     rsrcmgr = PDFResourceManager()
     sio = StringIO()
     # codec = 'utf-8'
     laparams = LAParams()
     device = TextConverter(rsrcmgr, sio, laparams=laparams)
     interpreter = PDFPageInterpreter(rsrcmgr, device)

     pdfname = TotalFilename

     fp = open(pdfname, 'rb')
     for page in PDFPage.get_pages(fp):
          interpreter.process_page(page)
     fp.close()
     text = sio.getvalue()
     # name = re.findall("(^[A-Z][A-Za-z]*\s+[A-Z][A-Za-z]*\s+[A-Z][A-Za-z]*)|(^[A-Z][A-Za-z]*\s+[A-Z][A-Za-z]*)",text)
     # print(name)
     return text


def doc_Reader(TotalFilename):

     mytxt=docx2txt.process(TotalFilename)
     with open("demo.txt",'w') as text_file:
          print(mytxt,file=text_file)
     #print()
     f=open("demo.txt")
     #print(f.read())
     document=f.read()
     # name = re.findall("(^[A-Z][A-Za-z]*\s+[A-Z][A-Za-z]*\s+[A-Z][A-Za-z]*)|(^[A-Z][A-Za-z]*\s+[A-Z][A-Za-z]*)", document)
     # print(name)
     return document


def extract_Experience(document,exp_Type):
     lines=[line.strip() for line in document.split("\n") if len(line) >0]
     lines=[nltk.word_tokenize(line) for line in lines]
     lines=[nltk.pos_tag(line) for line in lines]


# def extract_Experience(exp_Type,lines,sentences):
     for sentence in lines:
          sen=" ".join([words[0].lower() for words in sentence])
          if re.search(exp_Type,sen):
               sen_tokenised=nltk.word_tokenize(sen)
               # print(sen_tokenised)
               tagged=nltk.pos_tag(sen_tokenised)
               # print(tagged)
               entities=nltk.chunk.ne_chunk(tagged)
               for subtree in Tree.subtrees(entities):
                    for index in range(len(subtree)):
                         if subtree[index][1] == 'CD':
                              return subtree[index][0], subtree[index + 1][0]


def Experience_extracter(TotalFilename):
     fileType=TotalFilename[-4:]
     # print(fileType)
     if fileType == ".pdf":
          document=pdf_Reader(TotalFilename)
     elif fileType == "docx":
          document=doc_Reader(TotalFilename)
     # print(document)
     experience=extract_Experience(document,"Total experience")
     # print(ex)
     if experience is None:
          experience = extract_Experience(document, "experience")
          # print(experience)
     if experience is None:
          return 0, "months"
     return list(experience)


# def Name_Extracter(TotalFilename):
#      fileType=TotalFilename[-4:]
#      # print(fileType)
#      if fileType == ".pdf":
#           document=pdf_Reader(TotalFilename)
#      elif fileType == "docx":
#           document=doc_Reader(TotalFilename)
#
#      print(document.encode("utf-8"))
#
# TotalFilename = r"C:\Users\Shastransu\PycharmProjects\Resume\Resumes\Resume_prasoon.pdf"
# Name_Extracter(TotalFilename)