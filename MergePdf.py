import os
import re
import PyPDF2

# Caminhos de origem
caminho_planilhas = r'D:\Documentos\JOBS\DATA\RelatórioAndiappToledo\Toledo 2025_Edit\Planilha Corretores 2025'
caminho_nf = r'D:\Documentos\JOBS\DATA\RelatórioAndiappToledo\Toledo 2025_Edit\N Fiscais Toledo 2025'

# Pasta destino
pasta_destino = r'D:\Documentos\JOBS\DATA\RelatórioAndiappToledo\Toledo 2025_Edit\Arquivos Tratados'
os.makedirs(pasta_destino, exist_ok=True)

# Lista com caminhos completos dos PDFs
lista_arquivos = []
for caminho in [caminho_planilhas, caminho_nf]:
    for arquivo in os.listdir(caminho):
        if arquivo.lower().endswith('.pdf'):
            lista_arquivos.append(os.path.join(caminho, arquivo))

# Processar nomes para identificar clientes
nome_clientes = []
for caminho_arquivo in lista_arquivos:
    arquivo = os.path.basename(caminho_arquivo)
    nome = arquivo.title()

    # Remover prefixos
    for prefixo in ['Planilha Comissão ', 'Prestação De Contas ', 'Nf ']:
        if nome.startswith(prefixo.title()):
            nome = nome.replace(prefixo.title(), '')

    # Limpar extensões e textos entre parênteses
    nome = nome.replace('.Pdf', '')
    nome = re.sub(r'\s*\(.*?\)', '', nome).strip()

    # Tratar "Espólio"
    if nome.lower().startswith(('espólio', 'espolio', 'espólio de', 'espolio de')):
        nome = re.sub(r'^(Esp[óo]lio( De)?)\s*', '', nome, flags=re.IGNORECASE) + ' - Espólio'

    nome_clientes.append((nome.strip(), caminho_arquivo))

# Separar por prefixo (PR, PL, NF)
lista_pr = []
lista_pl = []
lista_nf = []
for cliente, caminho_arquivo in nome_clientes:
    prefixo = os.path.basename(caminho_arquivo)[0:2].upper()
    if prefixo == 'PR':
        lista_pr.append((cliente, caminho_arquivo))
    elif prefixo == 'PL':
        lista_pl.append((cliente, caminho_arquivo))
    elif prefixo == 'NF':
        lista_nf.append((cliente, caminho_arquivo))

# Criar conjunto com todos os clientes
todos_clientes = sorted(set([c for c, _ in lista_pr + lista_pl + lista_nf]))

# Gerar lista_ok
lista_ok = []
for cliente in todos_clientes:
    origens = []
    if cliente in [c for c, _ in lista_pr]:
        origens.append('PR')
    if cliente in [c for c, _ in lista_pl]:
        origens.append('PL')
    if cliente in [c for c, _ in lista_nf]:
        origens.append('NF')

    if len(origens) >= 2:
        lista_ok.append(cliente)

# Mesclar PDFs de cada cliente
for cliente in lista_ok:
    merger = PyPDF2.PdfMerger()

    # Adicionar todos PDFs encontrados para esse cliente
    for lista in [lista_pr, lista_pl, lista_nf]:
        for nome, caminho_pdf in lista:
            if nome == cliente:
                merger.append(caminho_pdf)

    # Salvar PDF final
    caminho_saida = os.path.join(pasta_destino, f"{cliente}.pdf")
    with open(caminho_saida, 'wb') as f_out:
        merger.write(f_out)
    merger.close()

print(f"✅ Processamento concluído. Arquivos salvos em: {pasta_destino}")