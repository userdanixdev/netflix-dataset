import pandas as pd
import glob
import os
import openpyxl

# Criar uma variável caminho para ler os arquivos:
folder_path = 'source\\data-raw\\netflix-dataset\\raw'
# O glob vai listar todos os arquivos de excel:
excel_files = glob.glob(os.path.join(folder_path,'*.xlsx'))
#print(excel_files)
if not excel_files:
    print('Nenhum arquivo compatível encontrado.')
else:
    dfs=[]
    for excel_file in excel_files:
        try:
            df_temp=pd.read_excel(excel_file)
            file_name = os.path.basename(excel_file)
            if 'brasil' in file_name.lower():
                df_temp['Location'] = 'Brasil'
            elif 'france' in file_name.lower():
                df_temp['Location'] = 'France'
            elif 'italian' in file_name.lower():
                df_temp['Location'] = 'Italia' 

            # Criar uma nova coluna chamada campanha e extrair tudo o que vem depois de 'utm_campaign='
            df_temp['Campanha']=df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')     
            
            # Guardar os dados tratatos dentro do data frame:
            dfs.append(df_temp)        
        
        except Exception as e:
            print(f'Erro ao ler o arquivo {excel_file}:{e}')
if dfs:
    result = pd.concat(dfs)
    # Caminho de saída:
    output_file=os.path.join('source','data-ready','clean_3200.xlsx')
    writer=pd.ExcelWriter(output_file)
    
    # Resultado oficinal 'result':
    result.to_excel(writer)
    # salvar para efetivar o processo:
    writer._save()
else:
    print('Nenhum dado salvo')    
