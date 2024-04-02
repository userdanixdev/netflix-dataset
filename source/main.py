import pandas as pd # Manipulação de análise de dados
import os  # manipular arquivos do sistema operacional
import glob # manipular arquivos em massa a partir da raiz
import openpyxl

# Criar uma variável caminho para ler os arquivos:
folder_path='netflix-dataset\\raw'

# Guardar os arquivos de excel:
excel_files=glob.glob(os.path.join(folder_path,'*.xlsx'))
if not excel_files:
    print('Nenhum arquivo Excel encontrado')
else:        

    # Data Frame = Uma tabela na memória para guardar os arquivos 
    dfs = []
    # Para cada arquivo colocar no data_frame temporário:
    for excel_file in excel_files:
        try:
            # Ler o nome do arquivo
            df_temp = pd.read_excel(excel_file) 
            # Abaixo pega o caminho do arquivo criando uma COLUNA A MAIS
            # para rastreá-lo.
            file_name = os.path.basename(excel_file) 

            # Adicionar coluna
            df_temp['fileName']=file_name
            
            if 'brasil' in file_name.lower():
                df_temp['location']='br'
            elif 'france' in file_name.lower():
                df_temp['location']='fr'
            elif 'italian' in file_name.lower():
                df_temp['location']='ita'
            # criar uma nova coluna chamada CAMPANHA:
            df_temp['CAMPANHA']=df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')                           
            dfs.append(df_temp)
        except Exception as e:
            print(f'Erro ao ler o arquivo {excel_file} : {e}')
            # caso não leia, exiba a mensagem
            # Concatenar um texto com variável = f'(f string)
if dfs:
    # Concatena todas as tabelas salvas no dfs em uma única tabela:
    result = pd.concat(dfs)            
    # Caminho de saída:
    output_file = os.path.join('netflix-dataset','ready','clean3.xlsx')
    # Configuração do motor da escrita:
    writer = pd.ExcelWriter(output_file)
    result.to_excel(writer,index=False)
    # Salva o arquivo de excel:
    writer._save()
else:
    print('Nenhum dado para ser salvo')



