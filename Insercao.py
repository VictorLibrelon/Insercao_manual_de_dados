import tkinter as tk
from tkinter import filedialog, messagebox
import os
from os import listdir
from os.path import isfile, join
import csv
from DBLib import SqlServer
import time
import datetime

from  Utility import Folder

sql = SqlServer()

class P:
    def ConvertFile(self, csvFile):
        fileToConvert = csvFile

        convertedFile = r"C:\caminho\do\arquivo\csv\\" + os.path.basename(fileToConvert).replace(".csv","") + "Converted.csv"

        updatedCount = 0
        categoryCount = 0
        cityCount = 0
        countryCount = 0


        with open(fileToConvert, encoding="ISO-8859-1") as csv_file:
            f = open(convertedFile, "w")
            csv_reader = csv.reader(csv_file, delimiter=',')
            for row in csv_reader:
                
                if (row[0] == "assignment_group") or (row[0] == "req_assignment_group"):
                    if(row[8]=="updated") or (row[8]=="sys_updated_on") or (row[8]=="req_sys_updated_on"):
                        updatedCount = 1
                    if(row[9]=="category"):
                        categoryCount = 1
                    if(row[11]=="req_location.city") or (row[11]=="location.city"):
                        cityCount = 1
                    if(row[12]=="req_location.country") or (row[12]=="location.country"):
                        countryCount = 1
                    continue
                try: 
                    assignmentGroup =row[0]
                    assignedTo=row[1]
                    number=row[2]
                    priority=row[3]
                    state=row[4]
                    opened=row[5]
                    if opened !='':
                        opened = datetime.datetime.strptime(opened, "%Y-%m-%d %H:%M:%S").strftime("%m/%d/%Y %H:%M")                      
                    closed = row[6]
                    if closed !='':
                        closed = datetime.datetime.strptime(closed, "%Y-%m-%d %H:%M:%S").strftime("%m/%d/%Y %H:%M")
                    updatedBy=row[7]  
                    if(updatedCount == 1):                  
                        updated=row[8]
                    else:
                        updated= "" 
                    if updated !="":
                        updated = datetime.datetime.strptime(updated, "%Y-%m-%d %H:%M:%S").strftime("%m/%d/%Y %H:%M")  
                    if(categoryCount == 1):
                        category=row[9]
                    else:
                        category=""
                    shortDescription=row[10].replace(',', ' ')
                    if(cityCount == 1):
                        city=row[11]
                    else:
                        city=""
                    if(countryCount == 1):
                        country=row[12]                     
                    cat6=""

                    row = (assignmentGroup  + ','
                            + assignedTo + ','
                            + number + ','
                            + priority + ','
                            + state + ','
                            + str(opened) + ','
                            + str(closed) + ','
                            + updatedBy + ','
                            + str(updated) + ','
                            + category + ','
                            + shortDescription + ','
                            + city + ','
                            + country + ','
                            + cat6
                            )

                    f.write(row + "\n")

                except Exception as e:
                    print(e)
                    exit()
        f.close()
        updatedCount = 0
        categoryCount = 0
        cityCount = 0
        countryCount = 0
        return convertedFile
    

def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename()

    return file_path

def converter_csv(tipo_arquivo_selecionado, cliente_selecionado):
    caminho_arquivo = selecionar_arquivo()
    if not caminho_arquivo:
        print('nenhum arquivo selecionado')
        return

    p = P()
    convertedFileName = ''

    convertedFileName = p.ConvertFile(caminho_arquivo)

    #Limpa a tabela do cliente selecionado 
    try:
        sql.TruncateClientTable(cliente_selecionado)
    except Exception as e:
        print("Erro no truncate client. ", e)

    #Faz a inserção dos dados na tabela cliente selecionado
    try:
        sql.InsertBulkDB(cliente_selecionado, convertedFileName)
    except Exception as e:
        print("Erro na inserção de dados. ", e)    

    #Roda a proc para tranferir de tabela cliente para OP_CONTROL_AUTO
    try:
        sql.TransferLoadToClient(cliente_selecionado, tipo_arquivo_selecionado)
    except Exception as e:
        print("Erro na execução da proc. ", e)
        
    
    time.sleep(1)    
    


def iniciar_conversao():

    def fechar_aplicacao():

        #Roda o job OP_CONTROL_UPDATE
        sql.runJob()

        # Fecha a janela principal
        root.destroy()  

    root = tk.Tk()
    root.title("Inserção manual - Automation")
    root.geometry("600x400")

    label_titulo = tk.Label(root, text="Bem-vindo(a)!", font=("Arial", 16), bg="#f2f2f2")

    label_intrucao = tk.Label(root, text="Para iniciar selecione o tipo de arquivo da inserção depois cliente e por fim selecione o arquivo,\n para encerrar a aplicação basta clicar em ENCERRAR APLICAÇÃO", font=("Arial", 10), bg="#f2f2f2")

    # Lista com as opções do primeiro dropdown menu
    opcoes_tipo_arquivo = ["INC", "WO"]

    # Variável que vai armazenar a opção selecionada do primeiro dropdown menu
    escolha_tipo_arquivo = tk.StringVar()

    # Label para o primeiro dropdown menu
    label_tipo_arquivo = tk.Label(root, text="Selecione o tipo de arquivo:")

    # Dropdown menu com as opções do primeiro dropdown menu
    dropdown_tipo_arquivo = tk.OptionMenu(root, escolha_tipo_arquivo, *opcoes_tipo_arquivo)

    # Lista com as opções do segundo dropdown menu
    opcoes_cliente = ["Nome_Cliente", "Nome_Cliente"]

    # Variável que vai armazenar a opção selecionada do segundo dropdown menu
    escolha_cliente = tk.StringVar()

    # Label para o segundo dropdown menu
    label_cliente = tk.Label(root, text="Selecione o cliente:")

    # Dropdown menu com as opções do segundo dropdown menu
    dropdown_cliente = tk.OptionMenu(root, escolha_cliente, *opcoes_cliente)

    # Botão para confirmar a seleção
    button = tk.Button(root, text="Selecionar", command= lambda: converter_csv(escolha_tipo_arquivo.get(), escolha_cliente.get()))

    button_fechar = tk.Button(root, text="Encerrar aplicação", command= lambda:fechar_aplicacao())
    
    label_titulo.pack(pady=15)
    label_intrucao.pack(pady=5)
    label_tipo_arquivo.pack(pady=10)
    dropdown_tipo_arquivo.pack(pady=5)
    label_cliente.pack(pady=10)
    dropdown_cliente.pack(pady=5)
    button.pack(pady=10)
    button_fechar.pack(pady=10)

    root.mainloop()

def Run():

    #Limpa a tabela OP_CONTROL_AUTO
    try:
        sql.TruncateOPCONTROLAUTOTable()
    except Exception as e:
        print("Erro na execução da proc. ", e)

    #Limpa a pasta de arquivos
    try:
        Folder.CleanFolder()
    except Exception as e:
        print("Erro na limpeza do diretorio de arquivos. ", e)

    #Inicia a interface grafica
    if __name__ == '__main__':
        iniciar_conversao()
    
    #Roda o job do BD
    '''try:
        sql.runJob()
    except Exception as e:
        print("Erro na execucao do job. ", e)'''

    

r = Run()