
import os
import pandas
import pdfplumber
import tkinter as tk     # from tkinter import Tk for Python 3.x
from tkinter import simpledialog
from tkinter.filedialog import askopenfilename,askdirectory
from PyPDF2 import PdfFileWriter, PdfFileReader
from playwright.async_api import async_playwright
from os import walk
import time
import asyncio
df= pandas.read_excel('./Pasta1.xlsx',usecols = 'B')

path = ''
out_dir = ''

async def uploadFile():

    f = []
    path= []
    for (dirpath, dirnames, filenames) in walk(out_dir):
        f = filenames
        path=[dirpath]
    print(path)
    async with async_playwright() as p:
        navegador = await p.chromium.launch(headless=False)
        pagina = await navegador.new_page()
        await pagina.goto('https://www.engecomp.ind.br/area-do-colaborador/manager/')
        await pagina.fill('xpath=/html/body/div/div/div/div/div[1]/div/form/div[1]/input', "Karla")
        await pagina.fill('xpath=/html/body/div/div/div/div/div[1]/div/form/div[2]/input',"KARLA1020")
        await pagina.locator('xpath=/html/body/div/div/div/div/div[1]/div/form/div[3]/div[1]/button').click()

        time.sleep(2)
        await pagina.locator('xpath=//html/body/div[2]/div[1]/nav/ul/li[5]/a').click()
        await pagina.locator('xpath=/html/body/div[2]/div[1]/nav/ul/li[5]/ul/li[1]/a').click()
        
        time.sleep(2)
        for file in filenames:
          name = file.replace('_Pagamento.pdf','')
          name = name.replace('_Adiantamento.pdf','')
          name = name.replace('_13º_1ª_Parcela.pdf','')
          name = name.replace('13º_2ª_Parcela.pdf','')
          name = name.replace('_PLR.pdf','')
          name = name.replace('_Demonstrativo_de_Rendimento.pdf','')
          name = name.replace('.pdf','')
          print(name)
          await pagina.fill('xpath=/html/body/div[2]/main/div/div/div[2]/div[2]/div/div[2]/label/input', name)
          try:
            await pagina.locator('xpath=/html/body/div[2]/main/div/div/div[2]/div[2]/div/table/tbody/tr/td[5]/a').click(timeout=2000)
            time.sleep(2)
            await pagina.locator('xpath=/html/body/div[2]/main/div/div/ul/li[2]/a').click(timeout=2000)
            
            with await pagina.expect_file_chooser() as file_chooser:
              await pagina.locator('xpath=/html/body/div[2]/main/div/div/div/div[3]/div[1]/div').click(timeout=2000) 
              file_chooser.value.setFiles(""+dirpath[0]+"/"+file+"")


            await pagina.go_back()
            time.sleep(2)
          except Exception as e: 
           print(e)

          time.sleep(2)

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
      pdf_out = os.path.join(out_dir, pdf_name + options.get()+'.pdf')

      #Grava o objeto de escrita no arquivo
      with open(pdf_out, 'wb') as pdf_named:
        pdf_writer.write(pdf_named)

      
      
def choose_file_folder():
    global path
    global out_dir     
    path = askopenfilename(defaultextension=".pdf",filetypes=[('pdf file', '*.pdf')]) # show an "Open" dialog box and return the path to the selected file

    out_dir = askdirectory()

def callUploadFunction():
  loop = asyncio.get_event_loop()
  loop.run_until_complete(uploadFile()) 
  loop.close()
#Testando as funções

choose_file_folder()

my_w = tk.Tk()
my_w.geometry("350x200")  # Size of the window 
my_w.title("www.plus2net.com")  # Adding a title

options = tk.StringVar(my_w)
options.set("_Pagamento") # default value

l1 = tk.Label(my_w,  text='Select One', width=10 )  
l1.grid(row=2,column=1) 
om1 =tk.OptionMenu(my_w, options, '_Pagamento','_Adiantamento','_13º_1ª_Parcela','13º_2ª_Parcela','_PLR','_Demonstrativo_de_Rendimento','')
om1.grid(row=3,column=1) 

b1= tk.Button(my_w, text ="SEPARAR ARQUIVOS", command = lambda: pdf_sep(path,out_dir) )
b1.grid(row=4,column=1)

b2= tk.Button(my_w, text ="TROCAR ARQUIVO E PASTA", command = lambda: choose_file_folder() )
b2.grid(row=5,column=1)

b2= tk.Button(my_w, text ="Upload arquivos", command = lambda: callUploadFunction())
b2.grid(row=6,column=1)  


my_w.mainloop()
