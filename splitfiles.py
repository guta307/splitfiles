
import os
import pandas
import pdfplumber
import tkinter as tk     # from tkinter import Tk for Python 3.x
import customtkinter
from PIL import Image
from PIL import ImageTk
from tkinter.messagebox import showinfo, showerror
from tkinter.filedialog import askopenfilename,askdirectory
from PyPDF2 import PdfWriter, PdfReader
from playwright.sync_api import sync_playwright
from os import walk
import time
df= pandas.read_excel('./Pasta1.xlsx',usecols = 'B')

path = ''
out_dir = ''

def uploadFile(root):
    f = []
    path= []
    error = []
    for (dirpath, dirnames, filenames) in walk(out_dir):
        f = filenames
    
        path=dirpath
    print(path)
    with sync_playwright() as p:
        navegador =  p.firefox.launch(headless=False)
        pagina =  navegador.new_page()
        pagina.goto('https://www.engecomp.ind.br/area-do-colaborador/manager/')
        pagina.fill('xpath=/html/body/div/div/div/div/div[1]/div/form/div[1]/input', "Karla")
        pagina.fill('xpath=/html/body/div/div/div/div/div[1]/div/form/div[2]/input',"KARLA1020")
        pagina.locator('xpath=/html/body/div/div/div/div/div[1]/div/form/div[3]/div[1]/button').click()

        time.sleep(2)
        pagina.locator('xpath=//html/body/div[2]/div[1]/nav/ul/li[5]/a').click()
        pagina.locator("text="+ company.get()+"").click()
        
        time.sleep(2)
        for file in filenames:
          name = file.replace('-Pagamento.pdf','')
          name = name.replace('-Adiantamento.pdf','')
          name = name.replace('-13º-1ª-Parcela.pdf','')
          name = name.replace('-13º-2ª-Parcela.pdf','')
          name = name.replace('-PLR.pdf','')
          name = name.replace('-Demonstrativo-de-Rendimento.pdf','')
          name = name.replace('-Vale-Transporte.pdf','')
          name = name.replace('.pdf','')
          name = name.replace('-',' ')
        
          pagina.fill('xpath=/html/body/div[2]/main/div/div/div[2]/div[2]/div/div[2]/label/input', name)
          pagina.click('tr:has-text("Sim") >> a')
          
          try:
             pagina.locator('/html/body/div[2]/main/div/div/ul/li[2]/a').click()
             with  pagina.expect_file_chooser() as fc_info:

               file_chooser = fc_info.value
               file_chooser.set_files(""+dirpath+"/"+file+"")
             time.sleep(5)
             os.remove(""+dirpath+"/"+file+"")
             
             pagina.go_back()
          
          except Exception as e:
            print(name)
            error.append("Erro no nome: "+name+"")
            

          
        showinfo("Upload concluído", "Processo finalizado")
        if len(error) > 0:
          showerror("Erro no upload",
            '\n '.join(map(' '.join, error))
            )
          navegador.close()
          root.destroy()

def pdf_get_data (page, pdf_file):

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
    pdf_content = PdfReader(pdf_file)
    

    #Armazena o quantidade de páginas do pdf original
    num_pages = len(pdf_content.pages)
    
    #Faz uma iteração para cada uma das páginas
    for page in range(num_pages):
      
      pdf_writer = PdfWriter()
      #Adiciona a página da iteração atual ao objeto para escrita do PDF
      pdf_writer.add_page(pdf_content.pages[page])
      
      #Invoca a função pdf_get_name para extrair o nome contido na página atual
      pdf_name = pdf_get_data(page,pdf_file).replace(" ","-")
    
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

def callFunction(function):
  root = tk.Tk().eval('tk::PlaceWindow %s center' % tk.winfo_pathname(tk.winfo_id()))
  root.withdraw()
  root.attributes("-topmost", True)
  
  root.geometry("350x200")
  root.title("Example")
  label = tk.Label(root, text="Waiting for task to finish.")
  label.pack()

  root.after(200, lambda : function(root))
  
  root.mainloop()
#Testando as funções
my_w = tk.Tk()
my_w.withdraw()
choose_file_folder()
my_w.deiconify()

#my_w.attributes("-fullscreen", True)
my_w.geometry("350x200")
my_w.title("ENGECOMP")  # Adding a title
my_w.config(bg='#0e3764')
frame = tk.Frame(my_w)
frame.place(relx=0.5, anchor="c")
frame.pack(fill='y',expand=1)
frame.config(bg='#0e3764')

options = tk.StringVar(my_w)
options.set("-Pagamento") # default value
operations = ('-Pagamento','-Adiantamento','-13-1ª-Parcela','-13º-2ª-Parcela','-PLR','-Demonstrativo-de-Rendimento','-Vale-Transporte')
menu_width = len(max(operations, key=len))
l1 = customtkinter.CTkLabel(frame,  text='tipo de arquivo', justify='center',text_color='white')
l1.grid(row=0,column=1)
frame.grid_rowconfigure(0,weight=14) 
om1 =customtkinter.CTkOptionMenu(frame,variable=options,values=operations,fg_color="#404040",dropdown_hover_color="#89CFF0") 

om1.grid(row=1,column=1)
frame.grid_rowconfigure(1,weight=14)
fileImage= customtkinter.CTkImage(Image.open('./imgs/file.png').resize((20,20),Image.ANTIALIAS))
b1= customtkinter.CTkButton(frame,image=fileImage,text="separar", command = lambda: pdf_sep(path,out_dir),fg_color="#03C03C", hover_color='white',width=190,height=40,compound='left',text_color='black')


b1.grid(row=2,column=1)
frame.grid_rowconfigure(2,weight=14)
folderImage= customtkinter.CTkImage(Image.open('./imgs/folder.png').resize((20,20),Image.ANTIALIAS))
b2= customtkinter.CTkButton(frame,image=folderImage,text="alterar",command = lambda: choose_file_folder(),fg_color="#fabb3d", hover_color='white',width=190,height=40,compound='left',text_color='black')

b2.grid(row=3,column=1)

frame.grid_rowconfigure(3,weight=14)

company = tk.StringVar(my_w)
company.set("ENGECOMP COMERCIO DE MATERIAIS E SERVICOS EIRELI") # default value
companies= ('ENGECOMP COMERCIO DE MATERIAIS E SERVICOS EIRELI','ENGECOMP FACILITIES E MANUTENCAO INDUSTRIAL EIRELI','ENGECOMP INSTALACOES INDUSTRIAIS ERELI','ENGECOMP MANUTENCAO INDUSTRIAL','ENGECOMP MONTAGENS E SERVIÇOS INDUSTRIAIS','ENGECOMP SERVICOS INDUSTRIAIS EIRELI (FILIAL)','ENGECOMP SERVICOS INDUSTRIAIS EIRELI (MATRIZ)','PJ')
l2 = customtkinter.CTkLabel(frame,  text='Empresa', justify='center',text_color='white') 
l2.grid(row=4,column=1)
frame.grid_rowconfigure(4,weight=14)
menu_width = len(max(companies, key=len))
om2 = customtkinter.CTkOptionMenu(frame,variable=company,values=companies,fg_color="#404040",dropdown_hover_color="#89CFF0") 
om2.grid(row=5,column=1) 
frame.grid_rowconfigure(5,weight=14)
uploadImage= customtkinter.CTkImage(Image.open('./imgs/upload.png').resize((20,20),Image.ANTIALIAS))
b2=customtkinter.CTkButton(frame,image=uploadImage,text="Upload",  command = lambda: callFunction(uploadFile),fg_color="#03C03C", hover_color='#0096FF',width=190,height=40,compound='left',text_color='black')

b2.grid(row=6,column=1)  
frame.grid_rowconfigure(6,weight=14)

my_w.mainloop()
