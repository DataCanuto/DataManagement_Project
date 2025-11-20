import os
import fitz  # PyMuPDF

def analisar_e_limpar_pdfs(caminho_da_pasta):
    """
    Analisa todos os arquivos PDF em uma pasta, relata suas propriedades
    e oferece a opção de remover páginas em branco.
    Filtra o relatório para exibir apenas arquivos com mais de uma página.

    Args:
        caminho_da_pasta (str): O caminho para a pasta contendo os arquivos PDF.
    """
    if not os.path.isdir(caminho_da_pasta):
        print(f"Erro: O caminho '{caminho_da_pasta}' não é uma pasta válida.")
        return

    # --- ESTRUTURAS DE DADOS PARA ARMAZENAR OS RESULTADOS ---
    relatorio_paginas = {}
    total_paginas_por_arquivo = {} # Dicionário auxiliar para guardar o total de páginas
    arquivos_com_mais_de_uma_pagina = 0
    arquivos_com_paginas_em_branco = {}

    print(f"--- Iniciando Análise da Pasta: {caminho_da_pasta} ---\n")

    lista_de_arquivos = [f for f in os.listdir(caminho_da_pasta) if f.lower().endswith('.pdf')]

    if not lista_de_arquivos:
        print("Nenhum arquivo PDF encontrado na pasta.")
        return

    # --- FASE 1: ANÁLISE DOS ARQUIVOS ---
    for nome_arquivo in lista_de_arquivos:
        caminho_completo = os.path.join(caminho_da_pasta, nome_arquivo)
        
        try:
            doc = fitz.open(caminho_completo)
            total_paginas = doc.page_count
            total_paginas_por_arquivo[nome_arquivo] = total_paginas # Armazena o total
            paginas_em_branco_indices = []
            
            # Verifica se o arquivo tem mais de uma página
            if total_paginas > 1:
                arquivos_com_mais_de_uma_pagina += 1

            # Itera por cada página para análise
            for i in range(total_paginas):
                page = doc.load_page(i)
                # Critério de página em branco: não tem texto e não tem imagens.
                if not page.get_text() and not page.get_images():
                    paginas_em_branco_indices.append(i)

            # Armazena a contagem de páginas que atendem ao critério (não estão em branco)
            paginas_validas = total_paginas - len(paginas_em_branco_indices)
            relatorio_paginas[nome_arquivo] = paginas_validas
            
            # Se encontrou páginas em branco, armazena quais são
            if paginas_em_branco_indices:
                arquivos_com_paginas_em_branco[nome_arquivo] = [p + 1 for p in paginas_em_branco_indices]

            doc.close()

        except Exception as e:
            print(f"Não foi possível processar o arquivo '{nome_arquivo}'. Erro: {e}")

    # --- FASE 2: APRESENTAÇÃO DOS RESULTADOS (COM FILTRO) ---
    print("--- Relatório da Análise Concluído ---\n")
    
    # MODIFICAÇÃO AQUI: Adicionado filtro para mostrar apenas arquivos com mais de 1 página
    print("1. Quantidade de páginas válidas (não-brancas) por arquivo (somente arquivos com >1 página):")
    if relatorio_paginas:
        arquivos_filtrados_exibidos = False
        for nome, qtd in relatorio_paginas.items():
            # A condição do filtro é aplicada aqui
            if total_paginas_por_arquivo.get(nome, 0) > 1:
                print(f"  - {nome}: {qtd} página(s)")
                arquivos_filtrados_exibidos = True
        if not arquivos_filtrados_exibidos:
             print("  Nenhum arquivo com mais de uma página foi encontrado para listar.")
    else:
        print("  Nenhum arquivo processado.")
    print("-" * 30)

    print(f"2. Total de arquivos com mais de uma página: {arquivos_com_mais_de_uma_pagina}")
    print("-" * 30)
    
    print("3. Arquivos que contêm páginas em branco:")
    if arquivos_com_paginas_em_branco:
        for nome, paginas in arquivos_com_paginas_em_branco.items():
            print(f"  - {nome} (Páginas em branco: {paginas})")
    else:
        print("  Nenhum arquivo com páginas em branco foi encontrado.")
    print("-" * 30)

    print("\n--- Verificação Manual Completa ---")


    print("\n-- ANÁLISE ---\n")
    print("Os arquivos com mais de uma página listados acima contêm informações válidas.")
    print("Os demais arquivos foram limpos de páginas em branco, se houveram.")
    
    # --- FASE 3: AÇÃO DE APAGAR PÁGINAS (COM CONFIRMAÇÃO) ---
    if not arquivos_com_paginas_em_branco:
        print("\nProcesso finalizado. Nenhuma ação de modificação necessária.")
        return

    prosseguir = input("\nDeseja apagar as páginas em branco listadas acima? (s/n): ").lower()

    if prosseguir == 's':
        print("\n--- Iniciando a remoção de páginas em branco ---")
        
        pasta_saida = os.path.join(caminho_da_pasta, "arquivos_limpos")
        os.makedirs(pasta_saida, exist_ok=True)
        print(f"Os arquivos modificados serão salvos em: '{pasta_saida}'")

        for nome_arquivo, paginas in arquivos_com_paginas_em_branco.items():
            try:
                caminho_original = os.path.join(caminho_da_pasta, nome_arquivo)
                doc = fitz.open(caminho_original)
                
                indices_para_remover = sorted([p - 1 for p in paginas], reverse=True)
                
                for indice in indices_para_remover:
                    doc.delete_page(indice)
                
                caminho_novo_arquivo = os.path.join(pasta_saida, nome_arquivo)
                doc.save(caminho_novo_arquivo, garbage=4, deflate=True)
                doc.close()
                print(f"  - Páginas removidas de '{nome_arquivo}'. Nova versão salva.")

            except Exception as e:
                print(f"  - Erro ao modificar o arquivo '{nome_arquivo}': {e}")
        
        print("\n--- Processo de limpeza concluído! ---")
    else:
        print("\nNenhuma modificação foi realizada. Processo encerrado.")


# --- COMO USAR O SCRIPT ---
if __name__ == "__main__":
    # IMPORTANTE: Substitua o caminho abaixo pela pasta onde estão seus PDFs.
    pasta_de_pdfs = "D:\Documentos\PyCharmD\.venv\Projetos 2025\ServiçoAnddiapToledo\Toledo 2025_Edit\Planilha Corretores 2025"
    
    analisar_e_limpar_pdfs(pasta_de_pdfs)

