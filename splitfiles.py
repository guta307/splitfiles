import nltk
import os
import pdfplumber
import fileinput
import tkinter as tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename,askdirectory
from PyPDF2 import PdfFileWriter, PdfFileReader
from nltk import ne_chunk, pos_tag, word_tokenize, sent_tokenize
from nltk.tree import Tree



def pdf_get_name (page, pdf_file):

  '''  
  page é o número da página
  pdf_file é caminho até o arquivo original 
  ''' 

  #O método open retorna uma instância da classe pdfplumber.PDF.
  pdf_content = pdfplumber.open(pdf_file)

  #Seleciona a página.
  pdf_page = pdf_content.pages[page]

  #Extrai o texto dividido por quebras de linha  
  pdf_text = pdf_page.extract_text().split('\n')
  name = ''
  
  for text in pdf_text:
        nltk_results = ne_chunk(pos_tag(word_tokenize(text)))
        for nltk_result in nltk_results:
            print(nltk_result)
 
 
  return name
  

def pdf_sep (pdf_file, out_dir):
  '''
  pdf_file é o caminho do pdf orinal
  out_dir é a pasta onde os pdfs divididos serão salvos
  '''
  #Abre o pdf original no modo de leitura
  with open(pdf_file, 'rb') as pdf:

    #Cria dois objetos, o primeiro da classe PdfFileReader para leitura e o segundo, da classe PdfFileWriter, para escrita
    pdf_content = PdfFileReader(pdf_file)
    pdf_writer = PdfFileWriter()

    #Armazena o quantidade de páginas do pdf original
    num_pages = pdf_content.getNumPages()
    
    #Faz uma iteração para cada uma das páginas
    for page in range(num_pages):
      
      #Adiciona a página da iteração atual ao objeto para escrita do PDF
      pdf_writer.addPage(pdf_content.getPage(page))

      #Invoca a função pdf_get_name para extrair o nome contido na página atual
      pdf_get_name(page,pdf_file)
     

#Testando as funções

#root = tk.Tk()
#root.withdraw() # we don't want a full GUI, so keep the root window from appearing
#path = askopenfilename() # show an "Open" dialog box and return the path to the selected file

#root.withdraw()
#out_dir = askdirectory()

pdf_sep('C:/Users/gustavo.veiga/Downloads/holerite.pdf','C:/Users/gustavo.veiga/Desktop/files')