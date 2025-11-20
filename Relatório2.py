import os
import zipfile
import win32com.client as win32
import pandas as pd
from lxml import etree as ET
from typing import List, Dict, Optional, Tuple
import re
from pathlib import Path
import unidecode

# Namespaces para XML do Word
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'v': 'urn:schemas-microsoft-com:vml',
    'o': 'urn:schemas-microsoft-com:office:office',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

SITUACAO_KEYWORDS = [
    "incontroverso",
    "valor total",
    "acordo",
    "homologado",
    "transitado em julgado",
    "sentença",
    "improcedente",
    "procedente",
    "extinto",
    "arquivado",
]

def normalize_text(text: str) -> str:
    """Retira acentos e coloca em upper case para busca mais eficiente."""
    return unidecode.unidecode(text or "").upper()

def find_situacao(text: str) -> Optional[str]:
    """
    Busca eficiente da situação no texto usando padrões flexíveis.
    Retorna a palavra-chave encontrada ou None.
    """
    norm_text = normalize_text(text)
    for keyword in SITUACAO_KEYWORDS:
        pattern = re.compile(r"\b" + re.escape(normalize_text(keyword)) + r"\b", re.IGNORECASE)
        if pattern.search(norm_text):
            return keyword.upper()
    return None

def extract_xml_texts(xml_content: bytes, xpath: str, namespaces: dict) -> List[str]:
    """Extrai textos de um XML usando xpath."""
    root = ET.fromstring(xml_content)
    return [
        (elem.text or "").strip()
        for elem in root.xpath(xpath, namespaces=namespaces)
        if elem.text and elem.text.strip()
    ]

def extract_all_text_from_docx(docx_path: Path) -> str:
    """Extrai todo texto de todas partes XML do documento."""
    texts = []
    with zipfile.ZipFile(docx_path) as archive:
        for name in archive.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                try:
                    xml_content = archive.read(name)
                    texts += extract_xml_texts(xml_content, ".//w:t", NS)
                except ET.XMLSyntaxError:
                    continue
    return " ".join(texts)

def extract_wordart_texts(docx_path: Path) -> List[str]:
    """Extrai textos de WordArt/VML/drawings do documento."""
    texts = []
    with zipfile.ZipFile(docx_path) as archive:
        for name in archive.namelist():
            if (name.startswith("word/header") or name.startswith("word/footer") or name == "word/document.xml") and name.endswith(".xml"):
                try:
                    xml_content = archive.read(name)
                    # WordArt em textboxes
                    texts += extract_xml_texts(xml_content, ".//v:textbox//w:t", NS)
                    # WordArt em shapes
                    texts += [
                        t.strip() for t in extract_xml_texts(xml_content, ".//v:shape//w:t", NS)
                    ]
                    # WordArt em textpath
                    root = ET.fromstring(xml_content)
                    for shape in root.xpath('.//v:shape[contains(@style, "mso-word-art")]', namespaces=NS):
                        for tpath in shape.xpath('.//v:textpath', namespaces=NS):
                            text = tpath.get('string', '')
                            if text.strip():
                                texts.append(text.strip())
                    # Desenhos
                    texts += extract_xml_texts(xml_content, ".//w:drawing//w:t", NS)
                except ET.XMLSyntaxError:
                    continue
    return texts

def extract_table_cell_text(table, row: int, col: int) -> Optional[str]:
    """Extrai texto de uma célula da tabela do Word."""
    try:
        if row <= table.Rows.Count and col <= table.Columns.Count:
            return clean_text(table.Cell(row, col).Range.Text)
    except Exception:
        pass
    return None

def clean_text(text: str) -> str:
    """Limpa texto de caracteres indesejados."""
    if not text:
        return ""
    cleaned = re.sub(r'[\x00-\x1F\x7F\r\x07]', '', text.strip())
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned

def extract_corretor_name(raw_text: str) -> Optional[str]:
    """Extrai nome do corretor por padrões flexíveis."""
    if not raw_text:
        return None
    cleaned = clean_text(raw_text)
    patterns = [
        r'%[\s]*(.+)',
        r'corretor[:\s]*(.+)',
        r'([A-ZÀ-Ü][a-zà-ü]+\s+[A-ZÀ-Ü][a-zà-ü]+)'
    ]
    for pattern in patterns:
        match = re.search(pattern, cleaned, re.IGNORECASE)
        if match:
            return clean_text(match.group(1))
    return cleaned

def extract_docx_data(docx_path: Path, word_app) -> Dict[str, Optional[str]]:
    """Extrai todos os dados relevantes do documento."""
    filename = docx_path.name
    result = {
        "Arquivo": filename,
        "Nome do Cliente": None,
        "Nome do Corretor": None,
        "Situação do Processo": None,
    }

    # Extrair textos (WordArt e gerais)
    wordart_texts = extract_wordart_texts(docx_path)
    all_text = extract_all_text_from_docx(docx_path)

    # Situação: busca eficiente, priorizando WordArt, depois texto geral
    situacao = None
    for text in wordart_texts:
        situacao = find_situacao(text)
        if situacao:
            break
    if not situacao:
        situacao = find_situacao(all_text)
    result["Situação do Processo"] = situacao

    # Extrair tabelas via COM
    abs_path = str(docx_path.resolve())
    doc = word_app.Documents.Open(abs_path, ReadOnly=True, Visible=False)
    try:
        for table in doc.Tables:
            rows, cols = table.Rows.Count, table.Columns.Count
            # Cliente (exemplo 4x14)
            if rows == 4 and cols == 14 and not result["Nome do Cliente"]:
                result["Nome do Cliente"] = extract_table_cell_text(table, 4, 1)
            # Corretor (exemplo 8x3)
            if rows == 8 and cols == 3 and not result["Nome do Corretor"]:
                raw_corretor = extract_table_cell_text(table, 7, 1)
                if raw_corretor:
                    result["Nome do Corretor"] = extract_corretor_name(raw_corretor)
    finally:
        doc.Close(SaveChanges=False)

    return result

def process_folder_docx(folder_path: Path) -> List[Dict[str, Optional[str]]]:
    """Processa todos os arquivos .docx da pasta."""
    arquivos = [p for p in folder_path.glob("*.docx") if p.is_file()]
    if not arquivos:
        raise FileNotFoundError(f"Não há arquivos .docx em {folder_path}")
    print(f"Encontrados {len(arquivos)} arquivos para processar...")

    word_app = win32.Dispatch("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = False

    resultados = []
    try:
        for i, arquivo_path in enumerate(arquivos, 1):
            print(f"[{i}/{len(arquivos)}] Processando: {arquivo_path.name}")
            try:
                dados = extract_docx_data(arquivo_path, word_app)
                resultados.append(dados)
                print(f"  ✓ Cliente: {dados['Nome do Cliente'] or 'N/A'}")
                print(f"  ✓ Corretor: {dados['Nome do Corretor'] or 'N/A'}")
                print(f"  ✓ Situação: {dados['Situação do Processo'] or 'N/A'}")
            except Exception as e:
                print(f"  ✗ Erro: {e}")
                resultados.append({
                    "Arquivo": arquivo_path.name,
                    "Nome do Cliente": None,
                    "Nome do Corretor": None,
                    "Situação do Processo": None
                })
    finally:
        word_app.Quit()
    return resultados

def save_results_to_excel(resultados: List[Dict[str, Optional[str]]], output_path: Path):
    """Salva os resultados em um arquivo Excel."""
    df = pd.DataFrame(resultados)
    df.to_excel(output_path, index=False)
    print(f"\n✓ Processamento concluído! Salvo em: {output_path}")
    situacoes = df['Situação do Processo'].notna().sum()
    print(f"Situações encontradas: {situacoes}/{len(df)}")

def main():
    pasta = Path(r"D:\Documentos\JOBS\DATA\RelatórioAndiappToledo\Dados\Ariel Anddiap")
    if not pasta.exists():
        raise SystemExit(f"Pasta não encontrada: {pasta}")
    resultados = process_folder_docx(pasta)
    if resultados:
        save_results_to_excel(resultados, Path("resumo_processos.xlsx"))
    else:
        print("Nenhum resultado foi gerado.")

if __name__ == "__main__":
    main()