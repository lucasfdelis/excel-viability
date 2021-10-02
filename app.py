import pandas as pd
import numpy as np
import os.path
import os
import time
from tkinter import filedialog
from tkinter import *
import pyexcel
import pyexcel_xls
import pyexcel_xlsx

print("DIGITE O NÚMERO DE PLANILHAS PARA CONSULTA")
n = input()
while(n.isdigit()==False):  
    os.system('CLS')
    print("ENTRADA DE DADOS INVÁLIDA, DIGITE UM NÚMERO.")
    n = input()
n = int(n)
i = 0

while(i < n):
    
    folder_selected = 'nan'
    os.system('CLS')
    print("SELECIONE A",i+1,"PLANILHA PARA ANÁLISE.")
    print('')

    time.sleep(1)
    root = Tk()
    root.withdraw()
    folder_selected = filedialog.askopenfilename()
    if(folder_selected.count('.csv')>0):
        os.system('CLS')
        print("CONVERTENDO PLANILHA .CSV PARA O FORMATO .XLSX")
        print("AGUARDE, ESSE PROCESSO PODE DEMORAR UM POUCO...")
        sheet = pyexcel.get_sheet(file_name=folder_selected, delimiter=";")
        folder_selected = folder_selected.replace('.csv','.xlsx')
        sheet.save_as(folder_selected)
    while(folder_selected == ''):
        os.system('CLS')
        print("FALHA AO SELECIONAR ARQUIVO. POR FAVOR, SELECIONE O ARQUIVO DESEJADO NOVAMENTE.")
        time.sleep(1)
        root = Tk()
        root.withdraw()
        folder_selected = filedialog.askopenfilename()
        if(folder_selected.count('.csv')>0):
            os.system('CLS')
            print("CONVERTENDO PLANILHA .CSV PARA O FORMATO .XLSX")
            print("AGUARDE, ESSE PROCESSO PODE DEMORAR UM POUCO...")
            sheet = pyexcel.get_sheet(file_name=folder_selected, delimiter=";")
            folder_selected = folder_selected.replace('.csv','.xlsx')
            sheet.save_as(folder_selected)
        folder_divided = folder_selected.split('/')
        nome = folder_divided[len(folder_divided)-1]
    folder_divided = folder_selected.split('/')
    nome = folder_divided[len(folder_divided)-1]
    os.system('CLS')

    print("GERANDO PLANILHA DE VIABILIDADE - DFV AL-BA-PB-PE-SE")
    print("AGUARDE...")
    
    info_workbook = folder_selected
    initial_workbook = 'C://Users//adm//Downloads//dfv-viabilidade-pub//initial//dfvceps.xlsx'
    output_workbook = 'C://Users//adm//Downloads//dfv-viabilidade-pub//output//VIABILIDADE-'+str(nome)

    df_initial = pd.read_excel(initial_workbook)
    df_info = pd.read_excel(info_workbook)
    
    listanum = []
    ceps = []
    enderecos = []
    CONTADOR = 0
    
    if not 'CEP' in df_info.columns and not 'cep_1' in df_info.columns and not 'ceps' in df_info.columns and not 'cep' in df_info.columns:
        os.system('CLS')
        print("GERANDO COLUNA DE CEPS, AGUARDE...")
        df_info.rename(columns={'endereco_1':'endereco'}, inplace=True)
        endereco = df_info['endereco'].to_list()
        tam = len(endereco)
        s = 0
        
        while(s < tam):
                endereco[s] = str(endereco[s])
                endereco[s] = endereco[s].replace('-','')
                endereco[s] = endereco[s].replace('.','')
                endereco[s] = endereco[s].replace(':',' ')
                endereco[s] = endereco[s].replace(',',' ')
                
                if(len(str(endereco[s])) < 5):
                    ceps.insert(CONTADOR,'')
                    CONTADOR = CONTADOR+1
                    
                if(len(str(endereco[s])) > 10):
                    enderecos = endereco[s].split()
                    k = len(enderecos)
                    j = 0
                    while(j < k):
                        if(enderecos[j].isdigit()==True):
                            listanum.insert(j,enderecos[j])
                        j = j+1
                    listanum = sorted(listanum, key=int, reverse=True)
                    listanum[0] = int(listanum[0])
                    ceps.insert(CONTADOR,listanum[0])
                    CONTADOR = CONTADOR+1
                    listanum = []
                s = s+1
        df_info.insert(1, "cep", ceps)
        df_info.to_excel(folder_selected, index=False, sheet_name="Sheet1")
    def pesquisa():
       if('telcelular' in df_info.columns):
           df_info.rename(columns={'CEP':'cep'}, inplace=True)
           df_info.rename(columns={'cep_1':'cep'}, inplace=True)
           df_info.rename(columns={'ceps':'cep'}, inplace=True)
           df_3 = pd.merge(df_initial, df_info[['cep','cidade','nome','cpf','rg','telcelular','endereco']], on='cep', how='left')
           df_3.dropna(subset=['nome'], inplace=True)
       if('CLIENTE' in df_info.columns):
           df_info.rename(columns={'CEP':'cep'}, inplace=True)
           df_info.rename(columns={'cep_1':'cep'}, inplace=True)
           df_info.rename(columns={'ceps':'cep'}, inplace=True)
           df_3 = pd.merge(df_initial, df_info[['cep','Cidade','CLIENTE','CPF/CNPJ do cliente.','TELEFONE','ENDEREÇO']], on='cep', how='left')
           df_3.dropna(subset=['CLIENTE'], inplace=True)
       if('NOME' in df_info.columns):
           df_info.rename(columns={'CEP':'cep'}, inplace=True)
           df_info.rename(columns={'cep_1':'cep'}, inplace=True)
           df_info.rename(columns={'ceps':'cep'}, inplace=True)
           df_3 = pd.merge(df_initial, df_info[['cep','Cidade','NOME','CPF/CNPJ do cliente','TELCEL','ENDEREÇO']], on='cep', how='left')
           df_3.dropna(subset=['NOME'], inplace=True)
       if('contato_1' in df_info.columns):
           df_info.rename(columns={'CEP':'cep'}, inplace=True)
           df_info.rename(columns={'cep_1':'cep'}, inplace=True)
           df_info.rename(columns={'ceps':'cep'}, inplace=True)
           df_info.rename(columns={'endereco_1':'endereco'}, inplace=True)
           df_3 = pd.merge(df_initial, df_info[['cep','documento','tipo_documento','nome','contato_1','endereco']], on='cep', how='left')
           df_3.dropna(subset=['nome'], inplace=True)
       if('DOC' in df_info.columns):
           df_info.rename(columns={'CEP':'cep'}, inplace=True)
           df_info.rename(columns={'cep_1':'cep'}, inplace=True)
           df_info.rename(columns={'ceps':'cep'}, inplace=True)
           df_3 = pd.merge(df_initial, df_info[['cep','cidade','nome','DOC','telefone','endereco']], on='cep', how='left')
           df_3.dropna(subset=['nome'], inplace=True)
       if('NOME_CONTATO_1' in df_info.columns):
           df_info.rename(columns={'CEP':'cep'}, inplace=True)
           df_info.rename(columns={'cep_1':'cep'}, inplace=True)
           df_info.rename(columns={'ceps':'cep'}, inplace=True)
           df_3 = pd.merge(df_initial, df_info[['cep','NOME_CONTATO_1','CNPJ','RAZAO_SOCIAL','LOGRADOURO','CIDADE','UF','COMPLEMENTO','BAIRRO','CONTATO_EMPRESA_1','CONTATO_SOCIO_1','CONTATO_EMPRESA_5','TERMINAL_CLIENTE']], on='cep', how='left')
           df_3.dropna(subset=['NOME_CONTATO_1'], inplace=True)
       if('contato' in df_info.columns):
           df_info.rename(columns={'CEP':'cep'}, inplace=True)
           df_info.rename(columns={'cep_1':'cep'}, inplace=True)
           df_info.rename(columns={'ceps':'cep'}, inplace=True)
           df_3 = pd.merge(df_initial, df_info[['cpf','cep','nome','endereco','contato']], on='cep', how='left')
           df_3.dropna(subset=['contato'], inplace=True)
       df_3.to_excel(output_workbook, index=False)
       os.system('CLS')
       
    pesquisa()

    print("GERANDO PLANILHA DE VIABILIDADE - DFV MG")
    print("AGUARDE...")
    
    initial_workbook = 'C://Users//adm//Downloads//dfv-viabilidade-pub//initial//dfvcepsmg.xlsx'
    output_workbook = 'C://Users//adm//Downloads//dfv-viabilidade-pub//output//VIABILIDADE-MG-'+str(nome)


    df_initial = pd.read_excel(initial_workbook)
    pesquisa()
    
    i = i+1
print("OPERAÇAO CONCLUÍDA COM SUCESSO. AS PLANILHAS ESTÃO ARMAZENADAS NAS PASTA 'OUTPUT'")
time.sleep(5)

