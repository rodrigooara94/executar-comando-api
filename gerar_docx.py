from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.table import WD_ROW_HEIGHT_RULE

def gerar_etiquetas_formatadas(df, nome_arquivo="etiquetas_final.docx"):
    doc = Document()

    # Configuração da página
    section = doc.sections[0]
    section.page_height = Inches(11.0)
    section.page_width = Inches(8.5)
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(0.4)
    section.right_margin = Cm(0.4)

    etiquetas_por_linha = 2
    etiquetas_por_pagina = 14

    for i, (_, row) in enumerate(df.iterrows()):
        if i % etiquetas_por_linha == 0:
            table = doc.add_table(rows=1, cols=etiquetas_por_linha)
            table.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            table.autofit = False
            table.allow_autofit = False

            row_obj = table.rows[0]
            row_obj.height = Cm(3.4)
            row_obj.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

            for cell in row_obj.cells:
                cell.width = Cm(10.16)

        col = i % etiquetas_por_linha
        cell = table.rows[0].cells[col]
        cell._element.clear_content()

        # Dividir cutter manualmente se contiver '\n'
        cutter_linhas = str(row['Código Cutter']).splitlines()
        linhas = [str(row['ID_Acervo'])] + cutter_linhas + [str(row['Ano']), "", str(row['Exemplar'])]

        for j, linha in enumerate(linhas):
            p = cell.add_paragraph()
            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            p_format = p.paragraph_format
            p_format.left_indent = Cm(4.8)
            p_format.first_line_indent = Cm(0)
            p_format.space_before = Pt(10) if j == 0 else Pt(0)
            p_format.space_after = Pt(0)
            # Espaçamento simples (padrão)

            run = p.add_run(str(linha))
            run.bold = True
            run.font.size = Pt(10)
            run.font.name = 'Times New Roman'
            r = run._element
            r.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

        if (i + 1) % etiquetas_por_pagina == 0:
            doc.add_page_break()

    doc.save(nome_arquivo)
    print(f"✅ Arquivo gerado: {nome_arquivo}")
