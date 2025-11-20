import os
import re
import pandas as pd

caminho_planilhas = r'D:\Documentos\PyCharmD\.venv\Projetos 2025\ServiçoAnddiapToledo\Toledo 2025_Edit\Planilha Corretores 2025'
caminho_nf = r'D:\Documentos\PyCharmD\.venv\Projetos 2025\ServiçoAnddiapToledo\Toledo 2025_Edit\N Fiscais Toledo 2025'

def extrair_nomes(lista):
    nomes = []
    for nome in lista:
        if nome.lower().endswith('.pdf'):
            nome = os.path.splitext(nome)[0]
            nome = re.sub(r'\(.*\)', "", nome)
            if nome.startswith('Planilha Comissão '):
                nome = nome.replace('Planilha Comissão ', '')
            elif nome.startswith('Prestação de Contas'):
                nome = nome.replace('Prestação de Contas ', '')
            elif nome.startswith('NF '):
                nome = nome.replace('NF ', '')
            nome = nome.strip()
            nomes.append(nome)
    return nomes

def tratar_espolio(lista):
    nomes_tratados = []
    for nome in lista:
        if 'Espólio' in nome:
            nome = nome.replace('Espólio', '').strip()
            nomes_tratados.append(nome + ' Espólio')
        elif 'Espolio' in nome:
            nome = nome.replace('Espolio', '').strip()
            nomes_tratados.append(nome + ' Espólio')
        else:
            nomes_tratados.append(nome)
    return nomes_tratados



arquivos_planilhas = os.listdir(caminho_planilhas)
arquivos_nf = os.listdir(caminho_nf)
arquivos = arquivos_planilhas + arquivos_nf

nomes = extrair_nomes(arquivos)
nomes = tratar_espolio(nomes)
nomes.sort()
nomes = list(set(nomes))
dataframe = pd.DataFrame({'nome': nomes})

print(len(nomes))
for nome in nomes:
    print(nome)

