import os
import pandas as pd
import time
from datetime import datetime

# Função para obter o estado a partir do código do gestor
def obter_estado(codigo):
    estados = {
        '11': 'Rondônia - RO',
        '12': 'Acre - AC',
        '13': 'Amazonas - AM',
        '14': 'Roraima - RR',
        '15': 'Pará - PA',
        '16': 'Amapá - AP',
        '17': 'Tocantins - TO',
        '21': 'Maranhão - MA',
        '22': 'Piauí - PI',
        '23': 'Ceará - CE',
        '24': 'Rio Grande do Norte - RN',
        '25': 'Paraíba - PB',
        '26': 'Pernambuco - PE',
        '27': 'Alagoas - AL',
        '28': 'Sergipe - SE',
        '29': 'Bahia - BA',
        '31': 'Minas Gerais - MG',
        '32': 'Espírito Santo - ES',
        '33': 'Rio de Janeiro - RJ',
        '35': 'São Paulo - SP',
        '41': 'Paraná - PR',
        '42': 'Santa Catarina - SC',
        '43': 'Rio Grande do Sul - RS',
        '50': 'Mato Grosso do Sul - MS',
        '51': 'Mato Grosso - MT',
        '52': 'Goiás - GO',
        '53': 'Distrito Federal - DF'
    }
    codigo_estado = codigo[:2]
    if codigo_estado in estados:
        return estados[codigo_estado]
    else:
        return 'Estado não encontrado'

# Pasta de entrada
pasta = 'PLANILHA'

# Verifica se a pasta existe e se contém arquivos válidos
if os.path.exists(pasta):
    arquivos = os.listdir(pasta)

    # Verifica se há arquivos na pasta
    if arquivos:
        dfs_dict_aba1 = {}
        dfs_dict_aba2 = {}
        dfs_dict_aba3 = {}

        for arquivo in arquivos:
            caminho_arquivo = os.path.join(pasta, arquivo)
            if os.path.isfile(caminho_arquivo) and caminho_arquivo.endswith('.xls*'):
                print(f'Lendo arquivo: {caminho_arquivo}')

                # Lê os dados da aba 1
                df_aba1 = pd.read_excel(caminho_arquivo, sheet_name='PLANEJADO')
                df_aba1['UF'] = ''
                for indice, linha in df_aba1.iterrows():
                    codigo_gestor = linha['COD_GESTOR']
                    uf = obter_estado(codigo_gestor)
                    df_aba1.at[indice, 'UF'] = uf

                dfs_dict_aba1[arquivo] = df_aba1

                # Lê os dados da aba 2
                df_aba2 = pd.read_excel(caminho_arquivo, sheet_name='FILAS')
                df_aba2['UF'] = ''
                for indice, linha in df_aba2.iterrows():
                    codigo_gestor = linha['COD_GESTOR']
                    uf = obter_estado(codigo_gestor)
                    df_aba2.at[indice, 'UF'] = uf

                dfs_dict_aba2[arquivo] = df_aba2

                # Lê os dados da aba 3
                df_aba3 = pd.read_excel(caminho_arquivo, sheet_name='CONSOLIDADO')
                df_aba3['UF'] = ''
                for indice, linha in df_aba3.iterrows():
                    codigo_gestor = linha['COD_GESTOR']
                    uf = obter_estado(codigo_gestor)
                    df_aba3.at[indice, 'UF'] = uf

                dfs_dict_aba3[arquivo] = df_aba3

        # Verifica se há dados nas abas
        if dfs_dict_aba1 and dfs_dict_aba2 and dfs_dict_aba3:
            # Concatena os dataframes de cada aba
            df_planilha_aba1 = pd.concat(dfs_dict_aba1.values(), ignore_index=True)
            df_planilha_aba2 = pd.concat(dfs_dict_aba2.values(), ignore_index=True)
            df_planilha_aba3 = pd.concat(dfs_dict_aba3.values(), ignore_index=True)

            # Continuação do código para processar e salvar os dados...
            # Por exemplo, a criação do arquivo Excel e a geração de relatórios

            print(f"[OK] GERANDO ARQUIVO PARA XLSX  =============================================>: {time.strftime('%H:%M:%S')}")
        else:
            print("Nenhuma informação encontrada nas abas.")
    else:
        print("Nenhum arquivo encontrado na pasta.")
else:
    print("A pasta de entrada não existe.")
