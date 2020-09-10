import pandas as pd
import numpy as np
import os.path
import os

print("DIGITE O NÚMERO DE PLANILHAS PARA CONSULTA, LEMBRE-SE DE COLOCÁ-LAS NA PASTA 'info' DESTE PROGRAMA.")
n = input()
while(n.isdigit()==False):
    os.system('CLS')
    print("ENTRADA DE DADOS INVÁLIDA, DIGITE UM NÚMERO.")
    n = input()
n = int(n)
i = 0
while(i < n):
    os.system('CLS')
    print("ESCREVA O NOME EXATO DA",i+1,"PLANILHA QUE ESTÁ NA PASTA, INCLUINDO O FORMATO DELA (.xlsx , por exemplo)")
    print('')
    print("OBS 1: PLANILHAS COM OUTROS FORMATOS (EX: .csv) NÃO ESTÃO DISPONÍVEIS PARA CONSULTA NO MOMENTO.")
    print("OBS 2: O CEP DO ENDEREÇO DA PLANILHA DEVE ESTAR SEPARADO EM UMA COLUNA EXCLUSIVA")
    nome = input()
    nome = nome.replace('.XLSX','.xlsx')
    while(os.path.exists('C://Users//adm//Downloads//dfv-viabilidade-pub//info//'+str(nome))==False or nome==''):
        os.system('CLS')
        print("ESTE ARQUIVO NÃO EXISTE. VERIFIQUE SE O NOME ESCRITO ESTÁ CORRETO.")
        nome = input()
        nome = nome.replace('.XLSX','.xlsx')
    os.system('CLS')

    print("GERANDO PLANILHA DE VIABILIDADE - DFV AL-BA-PB-PE-SE")
    print("AGUARDE...")
    initial_workbook = 'C://Users//adm//Downloads//dfv-viabilidade-pub//initial//dfvceps.xlsx'

    #PLANILHA CONSULTA
    info_workbook = 'C://Users//adm//Downloads//dfv-viabilidade-pub//info//'+str(nome)

    #PLANILHA QUE VAI SER GERADA
    output_workbook = 'C://Users//adm//Downloads//dfv-viabilidade-pub//output//VIABILIDADE-'+str(nome)
    #output_workbook = '//Adm-televenda//televendas pf//MAILING CONCORRENCIA//mailing viabilidade//VIABILIDADE-'+str(nome)

    df_initial = pd.read_excel(initial_workbook)
    df_info = pd.read_excel(info_workbook)
    if('telcelular' in df_info.columns):
        df_info.rename(columns={'CEP':'cep'}, inplace=True)
        df_info.rename(columns={'cep_1':'cep'}, inplace=True)
        df_info.rename(columns={'ceps':'cep'}, inplace=True)
        df_3 = pd.merge(df_initial, df_info[['cep','cidade','nome','cpf','rg','telcelular','endereco']], on='cep', how='left')
        df_3.dropna(subset=['nome'], inplace=True)
    if('contato_1' in df_info.columns):
        df_info.rename(columns={'CEP':'cep'}, inplace=True)
        df_info.rename(columns={'cep_1':'cep'}, inplace=True)
        df_info.rename(columns={'ceps':'cep'}, inplace=True)
        df_3 = pd.merge(df_initial, df_info[['cep','documento','tipo_documento','nome','contato_1','endereco_1']], on='cep', how='left')
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
    df_3.to_excel(output_workbook, index=False)
    os.system('CLS')

    print("GERANDO PLANILHA DE VIABILIDADE - DFV MG")
    print("AGUARDE...")
    
    initial_workbook = 'C://Users//adm//Downloads//dfv-viabilidade-pub//initial//dfvcepsmg.xlsx'

    #PLANILHA QUE VAI SER GERADA
    output_workbook = 'C://Users//adm//Downloads//dfv-viabilidade-pub//output//VIABILIDADE-MG-'+str(nome)
    #output_workbook = '//Adm-televenda\televendas pf//MAILING CONCORRENCIA//mailing viabilidade//VIABILIDADE-MG-'+str(nome)

    df_initial = pd.read_excel(initial_workbook)
    df_info = pd.read_excel(info_workbook)
    if('telcelular' in df_info.columns):
        df_info.rename(columns={'CEP':'cep'}, inplace=True)
        df_info.rename(columns={'cep_1':'cep'}, inplace=True)
        df_info.rename(columns={'ceps':'cep'}, inplace=True)
        df_3 = pd.merge(df_initial, df_info[['cep','cidade','nome','cpf','rg','telcelular','endereco']], on='cep', how='left')
        df_3.dropna(subset=['nome'], inplace=True)
    if('DOC' in df_info.columns):
        df_info.rename(columns={'CEP':'cep'}, inplace=True)
        df_info.rename(columns={'cep_1':'cep'}, inplace=True)
        df_info.rename(columns={'ceps':'cep'}, inplace=True)
        df_3 = pd.merge(df_initial, df_info[['cep','cidade','nome','DOC','telefone','endereco']], on='cep', how='left')
        df_3.dropna(subset=['nome'], inplace=True)
    if('contato_1' in df_info.columns):
        df_info.rename(columns={'CEP':'cep'}, inplace=True)
        df_info.rename(columns={'cep_1':'cep'}, inplace=True)
        df_info.rename(columns={'ceps':'cep'}, inplace=True)
        df_3 = pd.merge(df_initial, df_info[['cep','documento','tipo_documento','nome','contato_1','endereco_1']], on='cep', how='left')
        df_3.dropna(subset=['nome'], inplace=True)
    if('NOME_CONTATO_1' in df_info.columns):
        df_info.rename(columns={'CEP':'cep'}, inplace=True)
        df_info.rename(columns={'cep_1':'cep'}, inplace=True)
        df_info.rename(columns={'ceps':'cep'}, inplace=True)
        df_3 = pd.merge(df_initial, df_info[['cep','NOME_CONTATO_1','CNPJ','RAZAO_SOCIAL','LOGRADOURO','CIDADE','UF','COMPLEMENTO','BAIRRO','CONTATO_EMPRESA_1','CONTATO_SOCIO_1','CONTATO_EMPRESA_5','TERMINAL_CLIENTE']], on='cep', how='left')
        df_3.dropna(subset=['NOME_CONTATO_1'], inplace=True)
    df_3.to_excel(output_workbook, index=False)
    os.system('CLS')
    
    i = i+1
print("OPERAÇAO CONCLUÍDA COM SUCESSO. AS PLANILHAS ESTÃO ARMAZENADAS NAS PASTA 'OUTPUT'")

