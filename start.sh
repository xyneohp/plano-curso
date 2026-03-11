#!/usr/bin/env python3
"""
Generate Plano de Curso DOCX from form data.
Usage: python3 generate_docx.py <json_input_file> <output_file>
"""
import json
import sys
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_borders(cell, color='000000', size=4):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), str(size))
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)

def set_col_width(cell, width_twips):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(width_twips))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def add_bold_normal_para(cell, bold_text, normal_text='', align=WD_ALIGN_PARAGRAPH.LEFT):
    from docx.text.paragraph import Paragraph as DocxParagraph
    if isinstance(cell, DocxParagraph):
        p = cell
    else:
        p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    if bold_text:
        run = p.add_run(bold_text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(10)
    if normal_text:
        run2 = p.add_run(normal_text)
        run2.bold = False
        run2.font.name = 'Arial'
        run2.font.size = Pt(10)
    return p

def add_multiline_para(cell, text, size=10):
    lines = (text or '').strip().split('\n')
    first = True
    for line in lines:
        if first:
            p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
            first = False
        else:
            p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(line)
        run.font.name = 'Arial'
        run.font.size = Pt(size)

def add_section_header_row(table, title):
    row = table.add_row()
    cell = row.cells[0]
    # Merge all columns
    for i in range(1, len(row.cells)):
        cell = cell.merge(row.cells[i])
    set_cell_bg(cell, 'D0D8E8')
    set_cell_borders(cell)
    add_bold_normal_para(cell, title)
    return row

def merge_row_cells(row_obj):
    cells = row_obj.cells
    if len(cells) > 1:
        cells[0].merge(cells[-1])
    return cells[0]

def generate(data, output_path):
    doc = Document()
    
    # Page setup
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    # Page width in twips (content area)
    content_w = int((21 - 3) * 1440 / 2.54)  # ~11339 twips

    # Create main table with 2 columns
    col1 = int(content_w * 0.78)  # ~8844
    col2 = content_w - col1        # ~2495
    
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    table.allow_autofit = False
    
    # Set column widths
    for col_idx, width in enumerate([col1, col2]):
        for cell in table.columns[col_idx].cells:
            set_col_width(cell, width)

    # ── HEADER ROW: Logo + Institution
    hrow = table.add_row()
    logo_cell = hrow.cells[0]
    title_cell = hrow.cells[1]
    
    # Merge title cell spans - we'll use first cell for logo, merge rest for title
    logo_cell.merge(title_cell)
    big_cell = hrow.cells[0]
    set_cell_borders(big_cell)
    
    # Logo + text in same cell
    logo_path = os.path.join(os.path.dirname(__file__), 'logo.png')
    p_logo = big_cell.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(4)
    p_logo.paragraph_format.space_after = Pt(2)
    
    if os.path.exists(logo_path):
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Cm(3))
        p_logo.add_run('\t')
    
    run_name = p_logo.add_run('UNIVERSIDADE FEDERAL DO ACRE')
    run_name.bold = True
    run_name.font.name = 'Arial'
    run_name.font.size = Pt(12)
    
    p_proreitor = big_cell.add_paragraph('PRÓ-REITORIA DE GRADUAÇÃO')
    p_proreitor.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_proreitor.paragraph_format.space_before = Pt(0)
    p_proreitor.paragraph_format.space_after = Pt(4)
    r = p_proreitor.runs[0] if p_proreitor.runs else p_proreitor.add_run('PRÓ-REITORIA DE GRADUAÇÃO')
    r.bold = True
    r.font.name = 'Arial'
    r.font.size = Pt(11)

    # ── PLANO DE CURSO Title
    trow = table.add_row()
    tc = merge_row_cells(trow)
    set_cell_bg(tc, 'D0D8E8')
    set_cell_borders(tc)
    tp = tc.paragraphs[0]
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_before = Pt(3)
    tp.paragraph_format.space_after = Pt(3)
    tr = tp.add_run('PLANO DE CURSO')
    tr.bold = True
    tr.font.name = 'Arial'
    tr.font.size = Pt(12)

    # ── CENTRO
    crow = table.add_row()
    cc = merge_row_cells(crow)
    set_cell_borders(cc)
    add_bold_normal_para(cc, 'Centro: ', data.get('centro', ''))

    # ── CURSO
    curow = table.add_row()
    cuc = merge_row_cells(curow)
    set_cell_borders(cuc)
    add_bold_normal_para(cuc, 'Curso: ', data.get('curso', ''))

    # ── DISCIPLINA + CRÉDITOS
    drow = table.add_row()
    dc1 = drow.cells[0]
    dc2 = drow.cells[1]
    set_cell_borders(dc1)
    set_cell_borders(dc2)
    add_bold_normal_para(dc1, 'Disciplina: ', data.get('disciplina_cod', ''))
    add_bold_normal_para(dc2, 'Créditos: ', str(data.get('creditos', '')))

    # ── PRÉ-REQUISITOS + CO-REQUISITOS
    prow = table.add_row()
    pc1 = prow.cells[0]
    pc2 = prow.cells[1]
    set_cell_borders(pc1)
    set_cell_borders(pc2)
    add_bold_normal_para(pc1, 'Pré-requisitos: ', data.get('prerequisitos', 'Não há'))
    add_bold_normal_para(pc2, 'Co-requisitos: ', data.get('corequisitos', 'Não há'))

    # ── CH + CH ACEX + ENCONTROS (3 cols - split first cell)
    chrow = table.add_row()
    # We'll use a nested table approach via merge: col1 split into 2
    chc1 = chrow.cells[0]
    chc2 = chrow.cells[1]
    set_cell_borders(chc1)
    set_cell_borders(chc2)
    acex_val = data.get('acex', 0)
    acex_str = f"{acex_val}h" if acex_val and int(acex_val) > 0 else 'Não há'
    ch_text = f"Carga Horária: {data.get('ch', '')}h    |    CH de Acex: {acex_str}    |    Encontros: {data.get('encontros', '')}"
    add_bold_normal_para(chc1, '', ch_text)
    add_bold_normal_para(chc2, 'Período: ', data.get('semestre', ''))

    # ── DIAS/HORÁRIOS
    dh_row = table.add_row()
    dhc = merge_row_cells(dh_row)
    set_cell_borders(dhc)
    add_bold_normal_para(dhc, 'Dias/Horários de Aula: ', data.get('dias_horarios', ''))

    # ── PROFESSOR
    prow2 = table.add_row()
    pc = merge_row_cells(prow2)
    set_cell_borders(pc)
    add_bold_normal_para(pc, 'Professor(a): ', data.get('professor', ''))

    # ── I - EMENTA
    add_section_header_row(table, 'I - Ementa')
    er = table.add_row()
    ec = merge_row_cells(er)
    set_cell_borders(ec)
    add_multiline_para(ec, data.get('ementa', ''))

    # ── II - OBJETIVOS
    add_section_header_row(table, 'II - Objetivos de Ensino')
    
    og_row = table.add_row()
    og_c = merge_row_cells(og_row)
    set_cell_bg(og_c, 'F0F0F0')
    set_cell_borders(og_c)
    add_bold_normal_para(og_c, '1 - Objetivos Gerais')
    
    ogv_row = table.add_row()
    ogv_c = merge_row_cells(ogv_row)
    set_cell_borders(ogv_c)
    add_multiline_para(ogv_c, data.get('obj_gerais', ''))

    oe_row = table.add_row()
    oe_c = merge_row_cells(oe_row)
    set_cell_bg(oe_c, 'F0F0F0')
    set_cell_borders(oe_c)
    add_bold_normal_para(oe_c, '2 - Objetivos Específicos')

    oev_row = table.add_row()
    oev_c = merge_row_cells(oev_row)
    set_cell_borders(oev_c)
    add_multiline_para(oev_c, data.get('obj_especificos', ''))

    # ── III - CONTEÚDOS
    add_section_header_row(table, 'III - Conteúdos de Ensino')
    
    # Header row for units
    uh_row = table.add_row()
    uhc1 = uh_row.cells[0]
    uhc2 = uh_row.cells[1]
    set_cell_bg(uhc1, 'F0F0F0')
    set_cell_bg(uhc2, 'F0F0F0')
    set_cell_borders(uhc1)
    set_cell_borders(uhc2)
    add_bold_normal_para(uhc1, 'Unidades Temáticas')
    p_ch = uhc2.paragraphs[0]
    p_ch.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ch = p_ch.add_run('C/H')
    r_ch.bold = True
    r_ch.font.name = 'Arial'
    r_ch.font.size = Pt(10)

    unidades = data.get('unidades', [])
    if not unidades:
        unidades = [{'titulo': '', 'subtemas': '', 'ch': ''}]
    
    for i, u in enumerate(unidades):
        ur = table.add_row()
        uc1 = ur.cells[0]
        uc2 = ur.cells[1]
        set_cell_borders(uc1)
        set_cell_borders(uc2)
        
        titulo = u.get('titulo', '')
        subtemas = u.get('subtemas', '')
        
        up = uc1.paragraphs[0]
        up.paragraph_format.space_before = Pt(2)
        up.paragraph_format.space_after = Pt(1)
        rb = up.add_run(f'Unidade {i+1} - ')
        rb.bold = True
        rb.font.name = 'Arial'
        rb.font.size = Pt(10)
        rt = up.add_run(titulo)
        rt.font.name = 'Arial'
        rt.font.size = Pt(10)
        
        if subtemas:
            for line in subtemas.split('\n'):
                if line.strip():
                    sp = uc1.add_paragraph(f'  • {line.strip()}')
                    sp.paragraph_format.space_before = Pt(0)
                    sp.paragraph_format.space_after = Pt(0)
                    if sp.runs:
                        sp.runs[0].font.name = 'Arial'
                        sp.runs[0].font.size = Pt(10)
        
        ch_p = uc2.paragraphs[0]
        ch_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ch_r = ch_p.add_run(str(u.get('ch', '')))
        ch_r.font.name = 'Arial'
        ch_r.font.size = Pt(10)

    # ── IV - METODOLOGIA
    add_section_header_row(table, 'IV - Metodologia de Ensino')
    mr = table.add_row()
    mc = merge_row_cells(mr)
    set_cell_borders(mc)
    add_multiline_para(mc, data.get('metodologia', ''))

    # ── V - RECURSOS
    add_section_header_row(table, 'V - Recursos Didáticos')
    rr = table.add_row()
    rc = merge_row_cells(rr)
    set_cell_borders(rc)
    add_multiline_para(rc, data.get('recursos', ''))

    # ── VI - AVALIAÇÃO
    add_section_header_row(table, 'VI - Avaliação da Aprendizagem')
    avr = table.add_row()
    avc = merge_row_cells(avr)
    set_cell_borders(avc)
    add_multiline_para(avc, data.get('avaliacao', ''))

    # ── VII - BIBLIOGRAFIA
    add_section_header_row(table, 'VII - Bibliografia')
    
    for label, key in [
        ('1 - Bibliografia Básica', 'biblio_basica'),
        ('2 - Bibliografia Complementar', 'biblio_complementar'),
        ('3 - Bibliografia Sugerida', 'biblio_sugerida'),
    ]:
        lrow = table.add_row()
        lc = merge_row_cells(lrow)
        set_cell_bg(lc, 'F0F0F0')
        set_cell_borders(lc)
        add_bold_normal_para(lc, label)
        
        vrow = table.add_row()
        vc = merge_row_cells(vrow)
        set_cell_borders(vc)
        add_multiline_para(vc, data.get(key, ''))

    # ── VIII - CRONOGRAMA
    add_section_header_row(table, 'VIII - Cronograma da Disciplina')
    
    pr_row = table.add_row()
    pr_c = merge_row_cells(pr_row)
    set_cell_borders(pr_c)
    add_bold_normal_para(pr_c, 'Período de realização: ', data.get('periodo_realiz', ''))
    p_dh = pr_c.add_paragraph()
    add_bold_normal_para(p_dh, 'Dia e Horário de Execução: ', data.get('dias_horarios', ''))

    # Cronograma unidades header
    cu_h = table.add_row()
    cuh1 = cu_h.cells[0]
    cuh2 = cu_h.cells[1]
    set_cell_bg(cuh1, 'F0F0F0')
    set_cell_bg(cuh2, 'F0F0F0')
    set_cell_borders(cuh1)
    set_cell_borders(cuh2)
    add_bold_normal_para(cuh1, 'Unidades Temáticas')
    p_it = cuh2.paragraphs[0]
    p_it.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rit = p_it.add_run('Início / Término')
    rit.bold = True
    rit.font.name = 'Arial'
    rit.font.size = Pt(10)

    for i, u in enumerate(unidades):
        ur2 = table.add_row()
        uc2_1 = ur2.cells[0]
        uc2_2 = ur2.cells[1]
        set_cell_borders(uc2_1)
        set_cell_borders(uc2_2)
        add_bold_normal_para(uc2_1, f'Unidade {i+1}: ', u.get('titulo', ''))
        p_dates = uc2_2.paragraphs[0]
        p_dates.alignment = WD_ALIGN_PARAGRAPH.CENTER
        inicio = u.get('inicio', '....../....../......')
        termino = u.get('termino', '....../....../......')
        r_dates = p_dates.add_run(f'{inicio}   /   {termino}')
        r_dates.font.name = 'Arial'
        r_dates.font.size = Pt(10)

    # Avaliações cronograma
    av_h = table.add_row()
    avh1 = av_h.cells[0]
    avh2 = av_h.cells[1]
    set_cell_bg(avh1, 'F0F0F0')
    set_cell_bg(avh2, 'F0F0F0')
    set_cell_borders(avh1)
    set_cell_borders(avh2)
    add_bold_normal_para(avh1, 'Avaliação da Aprendizagem')
    p_dr = avh2.paragraphs[0]
    p_dr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rdr = p_dr.add_run('Data de Realização')
    rdr.bold = True
    rdr.font.name = 'Arial'
    rdr.font.size = Pt(10)

    avaliacoes = data.get('avaliacoes', [])
    default_avals = [
        {'desc': 'Avaliação 1 - N1', 'data': ''},
        {'desc': 'Avaliação 2 - N1', 'data': ''},
        {'desc': 'Avaliação 1 - N2', 'data': ''},
        {'desc': 'Avaliação 2 - N2', 'data': ''},
        {'desc': 'Realização da Prova Final', 'data': ''},
    ]
    if not avaliacoes:
        avaliacoes = default_avals

    for a in avaliacoes:
        ar = table.add_row()
        ac1 = ar.cells[0]
        ac2 = ar.cells[1]
        set_cell_borders(ac1)
        set_cell_borders(ac2)
        add_bold_normal_para(ac1, '', a.get('desc', ''))
        pd2 = ac2.paragraphs[0]
        pd2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rd2 = pd2.add_run(a.get('data', ''))
        rd2.font.name = 'Arial'
        rd2.font.size = Pt(10)

    # ── APROVAÇÃO
    add_section_header_row(table, 'Aprovação do Colegiado de Curso')
    apr_row = table.add_row()
    apr_c = merge_row_cells(apr_row)
    set_cell_borders(apr_c)
    add_multiline_para(apr_c, data.get('aprovacao', ''))

    # ── ASSINATURA
    sig_row = table.add_row()
    sig_c = merge_row_cells(sig_row)
    set_cell_borders(sig_c)
    sig_c.paragraphs[0].paragraph_format.space_before = Pt(4)
    add_bold_normal_para(sig_c, '', f"Local e Data: {data.get('local_data', '')}")
    sig_c.add_paragraph('')
    sp2 = sig_c.add_paragraph('Nome e Assinatura do(a) Professor(a): _______________________________________________')
    if sp2.runs:
        sp2.runs[0].font.name = 'Arial'
        sp2.runs[0].font.size = Pt(10)
    sig_c.add_paragraph('')

    doc.save(output_path)
    return output_path

if __name__ == '__main__':
    with open(sys.argv[1], encoding='utf-8') as f:
        data = json.load(f)
    generate(data, sys.argv[2])
    print('OK')
