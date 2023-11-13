import docx
from docx.text.paragraph import Paragraph
from docx.oxml.xmlchemy import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, Inches

def create_new_paragraph(paragraph, text):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    new_para.add_run(text)
    return new_para

def format_paragraph(paragraph):
    paragraph.paragraph_format.first_line_indent = Inches(.5)
    paragraph.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph.runs:
        run.font.size = Pt(14)
    return paragraph

def insert_paragraph_after(paragraph, text):
    new_para = create_new_paragraph(paragraph, text)
    format_paragraph(new_para)
    return new_para

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def bold_selection(para, text, size):
    text = text.strip()
    content = para.text
    para.text = ""
