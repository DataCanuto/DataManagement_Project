import os
import re
import pandas as pd

caminho = r'D:\Documentos\PyCharmD\.venv\Projetos 2025\ServiçoAnddiapToledo\Toledo 2025_Edit\Planilha Corretores 2025'
destino = r'D:\Documentos\PyCharmD\.venv\Projetos 2025\ServiçoAnddiapToledo\Toledo 2025_Edit\Relatórios\Relatório-planilha-comissao.xlsx'

arquivos = os.listdir(caminho)
planilha = [arquivo for arquivo in arquivos if arquivo.title().startswith('Planilha')]

dados = []
padrao = re.compile(r"Planilha Comissão (.*?)\s*\((.*?)\s+(.*?)\)\.pdf", re.IGNORECASE)

for arquivo in planilha:
    m = padrao.match(arquivo)
    if m:
        nome = m.group(1).strip()
        situacao = m.group(2).strip()
        corretor = m.group(3).strip()
    else:
        # Se não bater, tenta extrair algo ou deixa em branco
        nome = arquivo.replace("Planilha Comissão", "").replace(".pdf", "").strip()
        situacao = ""
        corretor = ""
    dados.append({
        "arquivo": arquivo,
        "nome": nome,
        "situacao": situacao,
        "corretor": corretor
    })

df = pd.DataFrame(dados)
print(f"\n🔢 Total de linhas no DataFrame: {len(df)}")
print(df)

def exportar_excel(df, destino):
    try:
        df.to_excel(destino, index=False)
        print(f"\n✅ Arquivo exportado com sucesso para: {destino}")
    except Exception as e:
        print(f"\n❌ Erro ao exportar Excel: {e}")

exportar_excel(df, destino)