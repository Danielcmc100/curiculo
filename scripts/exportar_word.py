"""Exporta o currículo LaTeX deste projeto para Word editável.

Dependência: python-docx. O parser cobre os comandos usados neste currículo.
"""
from pathlib import Path
import re
from io import BytesIO
from PIL import Image
from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

ROOT = Path(__file__).resolve().parents[1]
TEX = (ROOT / 'Kleber_Daniel_Mattos_Viana.tex').read_text(encoding='utf-8')
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Mm(210), Mm(297)
sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Mm(16)
for name in ['Normal', 'Title', 'Subtitle', 'Heading 1']:
    st = doc.styles[name]
    st.font.name = 'Source Sans 3'
    st.font.color.rgb = RGBColor(0, 0, 0)
    for border in list(st.element.findall('.//' + qn('w:pBdr'))):
        border.getparent().remove(border)
    st.paragraph_format.space_after = Pt(0)
    st.paragraph_format.line_spacing = Pt(14)
doc.styles['Normal'].font.size = Pt(11)
doc.styles['Title'].font.size = Pt(19)
doc.styles['Title'].font.bold = True
doc.styles['Title'].paragraph_format.line_spacing = Pt(23)
doc.styles['Subtitle'].font.size = Pt(14)
doc.styles['Subtitle'].font.italic = False
doc.styles['Subtitle'].paragraph_format.line_spacing = Pt(18)
doc.styles['Heading 1'].font.size = Pt(12)
doc.styles['Heading 1'].font.bold = True
doc.styles['Heading 1'].paragraph_format.space_before = Pt(14)
doc.styles['Heading 1'].paragraph_format.space_after = Pt(6)
doc.styles['Heading 1'].paragraph_format.keep_with_next = True
doc.core_properties.title = 'Kleber Daniel Mattos Viana Resume'
doc.core_properties.author = 'Kleber Daniel Mattos Viana'

def group(s, pos):
    assert s[pos] == '{'
    depth, start = 1, pos + 1
    pos += 1
    while depth:
        depth += (s[pos] == '{') - (s[pos] == '}')
        pos += 1
    return s[start:pos-1], pos

def inline(p, s, bold=None, italic=None, url=None):
    i = 0
    plain = ''
    def flush():
        nonlocal plain
        if not plain:
            return
        text = plain.replace('~', ' ').replace('--', '–')
        r = p.add_run(text)
        if bold is not None: r.bold = bold
        if italic is not None: r.italic = italic
        if url:
            link = OxmlElement('w:hyperlink')
            link.set(qn('r:id'), p.part.relate_to(url, RT.HYPERLINK, is_external=True))
            link.append(r._r)
            p._p.append(link)
        plain = ''
    while i < len(s):
        if s.startswith('\\\\', i):
            flush(); p.add_run().add_break(); i += 2; continue
        match = re.match(r'\\(textbf|textit|linked)\{', s[i:])
        if match:
            flush()
            name = match.group(1)
            value, j = group(s, i + len(match.group(0))-1)
            if name == 'linked':
                label, j = group(s, j)
                inline(p, label, bold, italic, value)
            else:
                inline(p, value, True if name == 'textbf' else bold,
                       True if name == 'textit' else italic, url)
            i = j
        else:
            plain += s[i]; i += 1
    flush()

def para(text='', style=None):
    p = doc.add_paragraph(style=style)
    inline(p, text)
    return p

def entry(label, date, city=''):
    labels = label.split('\\\\', 1)
    p = para(labels[0])
    p.paragraph_format.tab_stops.add_tab_stop(Mm(178), WD_TAB_ALIGNMENT.RIGHT)
    inline(p, '\t' + date)
    p.paragraph_format.keep_with_next = True
    if len(labels) > 1:
        p = para(labels[1])
        p.paragraph_format.tab_stops.add_tab_stop(Mm(178), WD_TAB_ALIGNMENT.RIGHT)
        if city: inline(p, '\t' + city)
        p.paragraph_format.keep_with_next = True
    elif city:
        p = para(city)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.keep_with_next = True

# Foto original com geometria oval; texto restante permanece nativo do Word.
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(8)
p.paragraph_format.line_spacing = 1
photo = BytesIO()
Image.open(ROOT / 'assets/foto.jpg').convert('RGB').save(photo, format='JPEG')
photo.seek(0)
pic = p.add_run().add_picture(photo, width=Mm(34.5))
geom = pic._inline.find('.//' + qn('a:prstGeom'))
geom.set('prst', 'ellipse')
for text, style in [
    ('Kleber Daniel Mattos Viana', 'Title'),
    ('Software Engineer | Python Backend | Automação de Estudos Elétricos', 'Subtitle')]:
    p = para(text, style)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(7)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
for n, (label, url) in enumerate([
    ('danielcmc100@gmail.com', 'mailto:danielcmc100@gmail.com'),
    ('+55 21 97517-6594', 'tel:+5521975176594'),
    ('https://github.com/Danielcmc100', 'https://github.com/Danielcmc100')]):
    if n: p.add_run('    ')
    inline(p, label, url=url)
p.paragraph_format.space_after = Pt(7)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
inline(p, 'https://www.linkedin.com/in/kleber-daniel-mattos-viana/',
       url='https://www.linkedin.com/in/kleber-daniel-mattos-viana/')

sections = re.split(r'\\cvsection\{([^}]+)\}', TEX.split('\\cvsection{Perfil}', 1)[1])
# Reintroduz a primeira seção removida pelo split.
sections = ['Perfil', sections[0]] + sections[1:]
pending_break = False
for title, body in zip(sections[::2], sections[1::2]):
    p = para(title, 'Heading 1')
    border = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    for k, v in {'val':'single','sz':'8','space':'2','color':'000000'}.items():
        bottom.set(qn('w:'+k), v)
    border.append(bottom); p._p.get_or_add_pPr().append(border)
    if title == 'Habilidades':
        # Reproduz as duas colunas do currículo de referência.
        table = doc.add_table(rows=2, cols=2)
        table.autofit = False
        for col in table.columns:
            col.width = Mm(89)
        blocks = re.findall(r'\\begin\{minipage\}\[t\]\{[^}]+\}\\raggedright(.*?)\\end\{minipage\}', body, re.S)
        for col, block in enumerate(blocks):
            chunks = re.split(r'\\par\\vspace\{[^}]+\}', block)
            for row, chunk in enumerate(chunks):
                p = table.cell(row, col).paragraphs[0]
                inline(p, re.sub(r'\s*\n\s*', ' ', chunk).strip())
                p.paragraph_format.space_after = Pt(10)
        continue
    body = re.sub(r'%[^\n]*', '', body)
    body = re.sub(r'\\begin\{minipage\}\[t\]\{[^}]+\}', '', body)
    body = re.sub(r'\\(?:end\{minipage\}|raggedright|hfill)', '', body)
    body = re.sub(r'\\(?:begin|end)\{itemize\}', '', body)
    body = re.sub(r'\\vspace\{[^}]+\}', '\n', body)
    body = body.replace('\\par', '\n\n').replace('\\end{document}', '')
    lines = [x.strip().removesuffix('\\\\') for x in body.splitlines() if x.strip()]
    for line in lines:
        if line == '\\newpage':
            pending_break = True
        elif line.startswith('\\entry'):
            args = []; j = len('\\entry')
            for _ in range(3):
                v, j = group(line, j); args.append(v)
            entry(*args)
        elif line.startswith('\\item '):
            p = para('• ' + line[6:])
            p.paragraph_format.left_indent = Pt(10)
            p.paragraph_format.first_line_indent = Pt(-10)
            p.paragraph_format.right_indent = Mm(37)
        else:
            p = para(line)
            if pending_break:
                p.paragraph_format.page_break_before = True
                pending_break = False
            if title == 'Perfil':
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            elif title == 'Habilidades':
                p.paragraph_format.space_after = Pt(6)
            elif title in ['Projetos', 'Cursos']:
                p.paragraph_format.space_after = Pt(5)

out = ROOT / 'output/word/Kleber_Daniel_Mattos_Viana.docx'
out.parent.mkdir(parents=True, exist_ok=True)
doc.save(out)
print(out)
