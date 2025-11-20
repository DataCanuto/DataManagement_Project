import os
import fitz  # PyMuPDF

import os
import fitz  # PyMuPDF

def apagar_ultima_pagina(diretorio, nome_arquivo):
    """
    Remove a última página de um arquivo PDF específico.
    Este método usa um arquivo temporário para garantir a compatibilidade e segurança,
    evitando o erro de "salvamento incremental".
    """
    caminho_completo = os.path.join(diretorio, nome_arquivo)
    
    # Define um nome para o arquivo temporário
    nome_temp = os.path.join(diretorio, "temp_" + nome_arquivo)

    # --- Verificações de Segurança ---
    if not os.path.isfile(caminho_completo):
        print(f"\nERRO: O arquivo '{nome_arquivo}' não foi encontrado no diretório especificado.")
        return

    try:
        # Abre o documento PDF original
        doc = fitz.open(caminho_completo)

        total_paginas = doc.page_count
        
        if total_paginas <= 1:
            print(f"\nAVISO: O arquivo '{nome_arquivo}' possui apenas uma página. Nenhuma ação foi tomada.")
            doc.close()
            return

        print(f"\nProcessando '{nome_arquivo}'...")
        print(f"O arquivo possui {total_paginas} páginas. A última página (nº {total_paginas}) será removida.")

        # Remove a última página
        doc.delete_page(total_paginas - 1)

        # --- NOVA LÓGICA DE SALVAMENTO ---
        # 1. Salva as alterações em um novo arquivo temporário
        doc.save(nome_temp, garbage=4, deflate=True)
        doc.close() # Fecha o documento original

        # 2. Remove o arquivo original antigo
        os.remove(caminho_completo)

        # 3. Renomeia o arquivo temporário para o nome do original
        os.rename(nome_temp, caminho_completo)

        print(f"\nSUCESSO! A última página do arquivo '{nome_arquivo}' foi removida.")

    except Exception as e:
        print(f"\nERRO: Ocorreu um problema ao processar o arquivo '{nome_arquivo}': {e}")
        # Se um erro ocorrer, verifica se um arquivo temporário ficou para trás e o remove
        if os.path.exists(nome_temp):
            os.remove(nome_temp)


# --- PONTO DE EXECUÇÃO PRINCIPAL ---
if __name__ == "__main__":
    # 1. CONFIGURE O DIRETÓRIO AQUI
    # IMPORTANTE: Substitua o caminho abaixo pelo caminho ABSOLUTO da pasta dos seus PDFs.
    # Exemplo Windows: dir = "C:\\Users\\SeuUsuario\\Desktop\\Meus Documentos"
    # Exemplo Linux/Mac: dir = "/home/seu_usuario/documentos/pdfs_para_limpar"
    dir = r"D:\Documentos\PyCharmD\.venv\Projetos 2025\ServiçoAnddiapToledo\Toledo 2025_Edit\Planilha Corretores 2025"
    
    # Verifica se o caminho foi alterado
    if "coloque/o/caminho" in dir:
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print("!!! ATENÇÃO: Você precisa editar o script e configurar a     !!!")
        print("!!! variável 'dir' com o caminho para a sua pasta de PDFs.  !!!")
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    else:
        # 2. PEÇA PARA O USUÁRIO INFORMAR O NOME DO ARQUIVO
        print("--- Ferramenta para Remover a Última Página de um PDF ---")
        print(f"Procurando arquivos no diretório: {dir}")
        print("AVISO: Esta ação irá modificar o arquivo original. Tenha um backup.\n")
        
        nome_do_arquivo = input("Informe o nome do arquivo PDF (ex: relatorio.pdf): ")

        # 3. CHAME A FUNÇÃO PARA FAZER A LIMPEZA
        if nome_do_arquivo:
            apagar_ultima_pagina(dir, nome_do_arquivo)
        else:
            print("Nenhum nome de arquivo foi inserido. Encerrando o programa.")