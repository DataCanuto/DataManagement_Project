import os
import win32com.client as win32
import pandas as pd

def extract_docx(path, word_app):
    """
    Abre um .docx via COM e extrai:
      - Nome do Cliente (tabela 1: célula 4×1)
      - Nome do Corretor (tabela 2: célula 6×1, texto após '% ')
      - Situação do Processo (primeiro texto de WordArt no cabeçalho ou corpo)
    Retorna um dict com esses campos.
    """
    def clean(txt):
        return txt.rstrip("\r\x07").strip()

    doc = word_app.Documents.Open(os.path.abspath(path), ReadOnly=True)

    # 1) Extrai todos os WordArt/textos dos headers
    watermarks = []
    for section in doc.Sections:
        # itera todos os tipos de cabeçalho
        for i in range(1, section.Headers.Count + 1):
            try:
                header = section.Headers(i)
                for shp in header.Shapes:
                    try:
                        t = clean(shp.TextFrame.TextRange.Text)
                    except Exception:
                        continue
                    if t:
                        watermarks.append(t)
            except Exception:
                continue

    # fallback: varre shapes do corpo, se não achar watermark no header
    if not watermarks:
        for shp in doc.Shapes:
            try:
                t = clean(shp.TextFrame.TextRange.Text)
            except Exception:
                continue
            if t:
                watermarks.append(t)

    # 2) Identifica tabela 1 (4×14) e tabela 2 (8×3)
    tbl1 = tbl2 = None
    for tbl in doc.Tables:
        r, c = tbl.Rows.Count, tbl.Columns.Count
        if (r, c) == (4, 14):
            tbl1 = tbl
        elif (r, c) == (8, 3):
            tbl2 = tbl
        if tbl1 and tbl2:
            break

    # helper para ler célula sem marcadores de rodapé
    def cell_text(tab, row, col):
        return clean(tab.Cell(row, col).Range.Text)

    # 3) Extrai Cliente da tabela 1, célula (4,1)
    cliente = cell_text(tbl1, 4, 1) if tbl1 else None

    # 4) Extrai Corretor da tabela 2, célula (7,1) e faz parsing após '% '
    raw = cell_text(tbl2, 7, 1) if tbl2 else ""
    if "% " in raw:
        corretor = raw.split("% ", 1)[1].strip()
    else:
        corretor = raw or None

    # 5) Situação: primeiro watermark encontrado
    situacao = watermarks[0] if watermarks else None

    doc.Close(False)
    return {
        "Arquivo": os.path.basename(path),
        "Nome do Cliente": cliente,
        "Nome do Corretor": corretor,
        "Situação do Processo": situacao
    }


if __name__ == "__main__":
    pasta = r"D:\Documentos\JOBS\DATA\RelatórioAndiappToledo\Dados\Ariel Anddiap"
    resultados = []

    # somente .docx
    arquivos = [f for f in os.listdir(pasta) if f.lower().endswith(".docx")]
    if not arquivos:
        raise SystemExit(f"Não há arquivos .docx em {pasta}")

    word = win32.Dispatch("Word.Application")
    word.Visible = False

    for nome in arquivos:
        full_path = os.path.join(pasta, nome)
        try:
            info = extract_docx(full_path, word)
            print(f"[✔] {nome} → {info}")
            resultados.append(info)
        except Exception as e:
            print(f"[✘] {nome} → ERRO: {e}")

    word.Quit()

    df = pd.DataFrame(resultados)
    print("\n=== DataFrame final ===")
    print(df)
    df.to_excel("resumo_processos.xlsx", index=False)