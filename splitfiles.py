
import os
import pandas
import pdfplumber
from charset_normalizer import md__mypyc
import tkinter as tk     # from tkinter import Tk for Python 3.x
from tkinter import simpledialog
from tkinter.filedialog import askopenfilename,askdirectory
from PyPDF2 import PdfFileWriter, PdfFileReader


df= pandas.read_excel('./Pasta1.xlsx',usecols = 'B')


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
              for index,row in df.iterrows():
                if row['FUNCIONÁRIO'].lower() in text.lower():
                  name=row['FUNCIONÁRIO'] 
  print(name)
  if name == '':
    name = 'nome_desconhecido ['+str(page)+']'
  return name
  

def pdf_sep (pdf_file, out_dir):
  '''
  pdf_file é o caminho do pdf original
  out_dir é a pasta onde os pdfs divididos serão salvos
  '''
  #Abre o pdf original no modo de leitura
  with open(pdf_file, 'rb') as pdf:

    #Cria dois objetos, o primeiro da classe PdfFileReader para leitura e o segundo, da classe PdfFileWriter, para escrita
    pdf_content = PdfFileReader(pdf_file)
    

    #Armazena o quantidade de páginas do pdf original
    num_pages = pdf_content.getNumPages()
    
    #Faz uma iteração para cada uma das páginas
    for page in range(num_pages):
      
      pdf_writer = PdfFileWriter()
      #Adiciona a página da iteração atual ao objeto para escrita do PDF
      pdf_writer.addPage(pdf_content.getPage(page))
      
      #Invoca a função pdf_get_name para extrair o nome contido na página atual
      pdf_name = pdf_get_name(page,pdf_file)
    
      #O médoto os.path.join() une o caminho para gravação, o nome e a extesão do arquivo pdf. 
      pdf_out = os.path.join(out_dir, pdf_name +'.pdf')

      #Grava o objeto de escrita no arquivo
      with open(pdf_out, 'wb') as pdf_named:
        pdf_writer.write(pdf_named)

      
      
     

#Testando as funções

root = tk.Tk()
root.withdraw() # we don't want a full GUI, so keep the root window from appearing
path = askopenfilename(defaultextension=".pdf",filetypes=[('pdf file', '*.pdf')]) # show an "Open" dialog box and return the path to the selected file
if path != '':
  root.withdraw()
  out_dir = askdirectory()
  if out_dir != '':
    pdf_sep(path,out_dir)