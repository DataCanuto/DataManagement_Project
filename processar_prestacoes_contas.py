import os
import re
import pandas as pd
from PyPDF2 import PdfReader


def extrair_dados_prestacao(caminho_pdf):
    """Extrai nome do cliente e número do processo de um arquivo PDF"""
    try:
        reader = PdfReader(caminho_pdf)
        texto = ''.join([page.extract_text() or '' for page in reader.pages])

        # Extrair nome do cliente
        nome_match = re.search(r'À\s+([^\n]+)', texto)
        nome_cliente = nome_match.group(1).strip() if nome_match else "Nome não encontrado"

        # Extrair número do processo (padrões variados)
        processo_match = re.search(r'(?:Processo|N° Processo)[:\s]*([\d\.\-/]+)', texto, re.IGNORECASE)
        numero_processo = processo_match.group(1).strip() if processo_match else "Número não encontrado"

        return nome_cliente, numero_processo

    except Exception as e:
        print(f"Erro ao processar {os.path.basename(caminho_pdf)}: {str(e)}")
        return "Erro no processamento", ""


def processar_prestacoes(pasta):
    """Processa todos os arquivos de prestação de contas na pasta"""
    dados = []

    for arquivo in os.listdir(pasta):
        if not (arquivo.lower().endswith('.pdf') and
                ('prestacao de contas' in arquivo.lower() or
                 'prestação de contas' in arquivo.lower())):
            continue

        caminho_completo = os.path.join(pasta, arquivo)
        nome_cliente, numero_processo = extrair_dados_prestacao(caminho_completo)

        dados.append({
            'Nome do Cliente': nome_cliente,
            'Número do Processo': numero_processo,
            'Arquivo': arquivo
        })

    df = pd.DataFrame(dados)
    return df.sort_values('Nome do Cliente')


def exportar_excel(df, destino):
    """Exporta DataFrame em Excel para o caminho informado"""
    try:
        df.to_excel(destino, index=False)
        print(f"\n✅ Arquivo exportado com sucesso para: {destino}")
    except Exception as e:
        print(f"\n❌ Erro ao exportar Excel: {e}")


# ===== Execução =====
pasta = r'D:\Documentos\PyCharmD\.venv\Projetos 2025\ServiçoAnddiapToledo\Toledo 2025_Edit\Planilha Corretores 2025'
destino = r'D:\Documentos\PyCharmD\.venv\Projetos 2025\ServiçoAnddiapToledo\Toledo 2025_Edit\Relatórios\Relatório-prestacao_contas.xlsx'

df_prestacoes = processar_prestacoes(pasta)

print("\n=== DataFrame ===")
print(df_prestacoes)

print("\n=== ESTATÍSTICAS ===")
print("Total de clientes sem número de processo encontrado:",
      len(df_prestacoes[df_prestacoes['Número do Processo'] == "Número não encontrado"]))
print("Total de clientes sem nome encontrado:",
      len(df_prestacoes[df_prestacoes['Nome do Cliente'] == "Nome não encontrado"]))

# Exportar Excel para o caminho definido
exportar_excel(df_prestacoes, destino)
