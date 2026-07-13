from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt


def set_font(style, size, bold=False, italic=False, east_asia="SimSun"):
    style.font.name = "Times New Roman"
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = italic
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii", "Times New Roman")
    rfonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi", "Times New Roman")
    rfonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cs", "Times New Roman")
    rfonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", east_asia)


def add_style(doc, name, base="Normal", size=12, bold=False, italic=False,
              alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=0, space_before=0,
              space_after=0, line_spacing=1.0):
    styles = doc.styles
    if name in styles:
        style = styles[name]
    else:
        style = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    if base and base in styles:
        style.base_style = styles[base]
    set_font(style, size, bold=bold, italic=italic)
    pf = style.paragraph_format
    pf.alignment = alignment
    pf.first_line_indent = Cm(first_indent) if first_indent else None
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    return style


def article_reference(path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2)

    normal = doc.styles["Normal"]
    set_font(normal, 12)
    pf = normal.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.first_line_indent = Cm(1.25)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.15

    for name, size, bold in [("Title", 14, True), ("Heading 1", 12, True), ("Heading 2", 12, True)]:
        st = doc.styles[name]
        set_font(st, size, bold=bold)
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER if name == "Title" else WD_ALIGN_PARAGRAPH.LEFT
        st.paragraph_format.first_line_indent = None
        st.paragraph_format.space_before = Pt(10 if name != "Title" else 0)
        st.paragraph_format.space_after = Pt(6)
        st.paragraph_format.keep_with_next = True

    add_style(doc, "Article Title", size=14, bold=True,
              alignment=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0, space_after=8)
    add_style(doc, "Article Meta", size=10,
              alignment=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0, space_after=6)
    add_style(doc, "Abstract", size=10,
              alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.0, space_after=3)
    add_style(doc, "Keywords", size=10,
              alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.0, space_after=6)
    add_style(doc, "English Title", size=12, bold=True,
              alignment=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0, space_after=6)
    add_style(doc, "Table Caption", size=10, bold=True,
              alignment=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.0, space_before=6, space_after=3)
    add_style(doc, "Table Note", size=9,
              alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.0, space_after=4)
    add_style(doc, "Bibliography", size=10,
              alignment=WD_ALIGN_PARAGRAPH.LEFT, first_indent=-0.75,
              line_spacing=1.0, space_after=2)

    for foot_name in ["Footnote Text", "Footnote Reference"]:
        if foot_name in doc.styles:
            set_font(doc.styles[foot_name], 10)
    doc.save(path)


def matrix_reference(path):
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width = Cm(29.7)
    sec.page_height = Cm(21)
    sec.top_margin = Cm(1.2)
    sec.bottom_margin = Cm(1.2)
    sec.left_margin = Cm(1.2)
    sec.right_margin = Cm(1.2)

    normal = doc.styles["Normal"]
    set_font(normal, 8.5)
    pf = normal.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.first_line_indent = None
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0

    set_font(doc.styles["Title"], 13, bold=True)
    doc.styles["Title"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(doc.styles["Heading 1"], 10, bold=True)
    add_style(doc, "Matrix Intro", size=9,
              alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.0, space_after=4)
    add_style(doc, "Table Note", size=8,
              alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.0, space_after=3)
    doc.save(path)


if __name__ == "__main__":
    article_reference("reference_article.docx")
    matrix_reference("reference_matrix.docx")
