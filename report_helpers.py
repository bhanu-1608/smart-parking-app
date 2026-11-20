"""
DeccanPulse Internship Report - Final Professional Generation Script
Matches exact structure, font, and layout of the sample Safeguard internship report PDF.
Incorporates:
 - MECS logo on cover and certificate pages
 - New high-resolution certificate image (from Bhanu_Datta_Nalam_.pdf)
 - Times New Roman 12pt body, 14pt headings, 11pt body text 
 - Proper header "Summer Internship Report | <Title>" running header on every chapter page
 - Footer "Matrusri Engineering College | Dept. of Computer Science & Engineering  Page X"
 - Roman numerals ii-ix for front matter, Arabic 1-N for chapters
 - Page break before every chapter
"""
import os, sys, shutil
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Emu, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

# ─── Paths ───────────────────────────────────────────────────────────────────
DOCS = Path("C:/smart_parking_app/docs")
SS   = DOCS / "screenshots"
CERT = str(DOCS / "certificate.jpg")          # 2000x1414 from new PDF
LOGO = str(DOCS / "mecs_logo.jpg")            # 411x353 MECS emblem
OUT  = Path("C:/smart_parking_app/DeccanPulse_Internship_Report.docx")
DL   = Path("C:/Users/dell/Downloads/DeccanPulse_Internship_Report.docx")

# ─── Colors ───────────────────────────────────────────────────────────────────
BLACK   = RGBColor(0,   0,   0)
NAVY    = RGBColor(15,  51,  71)
SAFFRON = RGBColor(192, 80,  32)
GOLD    = RGBColor(156, 100, 18)
CHARCOAL= RGBColor(30,  41,  59)
GRAY    = RGBColor(100, 116, 139)
WHITE   = RGBColor(255, 255, 255)

# ─── Utility helpers ─────────────────────────────────────────────────────────
def cell_bg(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd  = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hexcolor}" w:val="clear"/>')
    tcPr.append(shd)

def cell_margins(cell, top=100, bottom=100, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for side, val in [('top',top),('bottom',bottom),('left',left),('right',right)]:
        n = OxmlElement(f'w:{side}')
        n.set(qn('w:w'),    str(val))
        n.set(qn('w:type'), 'dxa')
        tcMar.append(n)
    tcPr.append(tcMar)

def tbl_borders(table, color="B0B0B0", sz="4"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top    w:val="single" w:sz="{sz}" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:color="{color}"/>'
        f'<w:insideV w:val="none"/>'
        f'<w:left w:val="none"/><w:right w:val="none"/>'
        f'</w:tblBorders>')
    tblPr.append(borders)

def no_tbl_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/><w:bottom w:val="none"/>'
        f'<w:left w:val="none"/><w:right w:val="none"/>'
        f'<w:insideH w:val="none"/><w:insideV w:val="none"/>'
        f'</w:tblBorders>')
    tblPr.append(borders)

def thick_line_above(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    bdr  = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="12" w:color="1E3A5F"/>'
        f'</w:tcBorders>')
    tcPr.append(bdr)

def add_field(parent_elem, instr):
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), instr)
    run = OxmlElement('w:r')
    fld.append(run)
    parent_elem.append(fld)

class Doc:
    """Slim wrapper – keeps state so callers don't juggle the raw docx object."""
    def __init__(self):
        self.d = Document()
        sec = self.d.sections[0]
        sec.top_margin    = Inches(1.0)
        sec.bottom_margin = Inches(1.0)
        sec.left_margin   = Inches(1.25)
        sec.right_margin  = Inches(1.0)
        sec.different_first_page_header_footer = True
        # Front-matter: Roman numerals (ii onward, cover = no number)
        sectPr = sec._sectPr
        pnt = OxmlElement('w:pgNumType')
        pnt.set(qn('w:fmt'),  'lowerRoman')
        pnt.set(qn('w:start'), '1')      # cover = i (hidden), cert = ii, etc.
        sectPr.append(pnt)
        # Roman footer for front matter
        ftr = sec.footer
        p   = ftr.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_field(p._p, 'PAGE')

    # ── New chapter section (Arabic, with running header+footer) ─────────────
    def start_body(self):
        sec2 = self.d.add_section(WD_SECTION.NEW_PAGE)
        sec2.top_margin    = Inches(1.0)
        sec2.bottom_margin = Inches(1.0)
        sec2.left_margin   = Inches(1.25)
        sec2.right_margin  = Inches(1.0)
        sec2.header.is_linked_to_previous = False
        sec2.footer.is_linked_to_previous = False
        sectPr = sec2._sectPr
        pnt = OxmlElement('w:pgNumType')
        pnt.set(qn('w:fmt'),  'decimal')
        pnt.set(qn('w:start'), '1')
        sectPr.append(pnt)
        # Running header
        hdr = sec2.header
        ph  = hdr.paragraphs[0]
        ph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        ph.paragraph_format.space_after = Pt(2)
        self._run(ph, "Summer Internship Report  |  DeccanPulse: Hyderabad Smart Parking Platform",
                  sz=8.5, italic=True, color=GRAY)
        # Running footer
        ftr = sec2.footer
        pf  = ftr.paragraphs[0]
        pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pf.paragraph_format.space_before = Pt(2)
        r = pf.add_run("Matrusri Engineering College  |  Dept. of Computer Science & Engineering\u2003\u2003\u2003Page\u00a0")
        r.font.name  = "Times New Roman"
        r.font.size  = Pt(8.5)
        r.font.color.rgb = GRAY
        add_field(pf._p, 'PAGE')

    # ── Primitive paragraph builders ─────────────────────────────────────────
    def _para(self, text="", align=WD_ALIGN_PARAGRAPH.LEFT, sz=12,
              bold=False, italic=False, color=BLACK, name="Times New Roman",
              space_before=0, space_after=6, line_spacing=None, keep_next=False):
        p = self.d.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after  = Pt(space_after)
        if line_spacing:
            p.paragraph_format.line_spacing = line_spacing
        if keep_next:
            p.paragraph_format.keep_with_next = True
        if text:
            self._run(p, text, sz=sz, bold=bold, italic=italic, color=color, name=name)
        return p

    def _run(self, para, text, sz=12, bold=False, italic=False,
             color=BLACK, name="Times New Roman", underline=False):
        r = para.add_run(text)
        r.font.name    = name
        r.font.size    = Pt(sz)
        r.font.bold    = bold
        r.font.italic  = italic
        r.font.underline = underline
        r.font.color.rgb = color
        return r

    def page_break(self):
        self.d.add_page_break()

    # ── Standard structural elements ─────────────────────────────────────────
    def h1(self, text):
        """Chapter heading – 15pt Times New Roman Bold Navy, centered underline"""
        p = self._para(text, align=WD_ALIGN_PARAGRAPH.CENTER, sz=15,
                       bold=True, color=NAVY, space_before=0, space_after=14, keep_next=True)
        # Underline the whole paragraph via border
        pPr = p._p.get_or_add_pPr()
        bdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'<w:bottom w:val="single" w:sz="6" w:color="0F3347" w:space="2"/>'
            f'</w:pBdr>')
        pPr.append(bdr)
        return p

    def h2(self, text, space_before=14):
        return self._para(text, sz=12.5, bold=True, color=NAVY,
                          space_before=space_before, space_after=6, keep_next=True)

    def h3(self, text):
        return self._para(text, sz=12, bold=True, italic=True, color=CHARCOAL,
                          space_before=10, space_after=4, keep_next=True)

    def body(self, text, space_after=8, justify=True, bold_prefix=None):
        p = self.d.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after  = Pt(space_after)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = Pt(14)
        if bold_prefix:
            r = p.add_run(bold_prefix)
            r.font.name = "Times New Roman"; r.font.size = Pt(12)
            r.font.bold = True; r.font.color.rgb = BLACK
        r = p.add_run(text)
        r.font.name = "Times New Roman"; r.font.size = Pt(12)
        r.font.color.rgb = BLACK
        return p

    def bullet(self, text, bold_prefix=None):
        p = self.d.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after  = Pt(4)
        p.paragraph_format.line_spacing = Pt(14)
        if bold_prefix:
            r = p.add_run(bold_prefix)
            r.font.name = "Times New Roman"; r.font.size = Pt(12)
            r.font.bold = True; r.font.color.rgb = BLACK
        r = p.add_run(text)
        r.font.name = "Times New Roman"; r.font.size = Pt(12)
        r.font.color.rgb = BLACK
        return p

    def image(self, path, caption, width=Inches(5.6)):
        if not os.path.exists(path):
            self.body(f"[IMAGE NOT FOUND: {path}]")
            return
        p = self.d.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after  = Pt(4)
        p.paragraph_format.keep_with_next = True
        p.add_run().add_picture(path, width=width)
        cp = self.d.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_after = Pt(10)
        r = cp.add_run(caption)
        r.font.name = "Times New Roman"; r.font.size = Pt(10)
        r.font.italic = True; r.font.color.rgb = CHARCOAL

    def callout(self, text, title="NOTE"):
        tbl = self.d.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0,0)
        cell_bg(cell, "F0F7FA")
        cell_margins(cell, top=100, bottom=100, left=180, right=120)
        tcPr = cell._tc.get_or_add_tcPr()
        bdr  = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:left w:val="single" w:sz="24" w:color="0F3347"/>'
            f'</w:tcBorders>')
        tcPr.append(bdr)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(14)
        self._run(p, f"[{title}]  ", sz=10, bold=True, color=NAVY)
        self._run(p, text, sz=10, italic=True, color=CHARCOAL)
        self.d.add_paragraph().paragraph_format.space_after = Pt(4)

    def add_table(self, headers, data, col_widths, caption=None, header_bg="1E3A5F"):
        if caption:
            cp = self.d.add_paragraph()
            cp.paragraph_format.space_before = Pt(8)
            cp.paragraph_format.space_after  = Pt(4)
            cp.paragraph_format.keep_with_next = True
            self._run(cp, caption, sz=10, bold=True, color=NAVY)

        tbl = self.d.add_table(rows=len(data)+1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_borders(tbl, color="A0A0A0", sz="4")

        # Header row
        hrow = tbl.rows[0]
        for ci, h in enumerate(headers):
            cell = hrow.cells[ci]
            cell.width = Inches(col_widths[ci])
            cell_bg(cell, header_bg)
            cell_margins(cell)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            self._run(p, h, sz=9.5, bold=True, color=WHITE)

        # Data rows
        for ri, row_data in enumerate(data):
            row = tbl.rows[ri+1]
            bg  = "F3F6FA" if ri % 2 == 1 else "FFFFFF"
            for ci, val in enumerate(row_data):
                cell = row.cells[ci]
                cell.width = Inches(col_widths[ci])
                cell_bg(cell, bg)
                cell_margins(cell, top=70, bottom=70)
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = Pt(13)
                self._run(p, str(val), sz=9.5, color=CHARCOAL)

        self.d.add_paragraph().paragraph_format.space_after = Pt(4)

    def code(self, text, caption=None):
        if caption:
            cp = self.d.add_paragraph()
            cp.paragraph_format.space_before = Pt(8)
            cp.paragraph_format.space_after  = Pt(3)
            cp.paragraph_format.keep_with_next = True
            self._run(cp, caption, sz=10, bold=True, color=NAVY)
        tbl  = self.d.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0,0)
        cell_bg(cell, "F4F7F9")
        cell_margins(cell, top=100, bottom=100, left=150, right=120)
        tcPr = cell._tc.get_or_add_tcPr()
        bdr  = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top    w:val="single" w:sz="4" w:color="CBD5E1"/>'
            f'<w:bottom w:val="single" w:sz="4" w:color="CBD5E1"/>'
            f'<w:left   w:val="single" w:sz="18" w:color="0F3347"/>'
            f'<w:right  w:val="single" w:sz="4" w:color="CBD5E1"/>'
            f'</w:tcBorders>')
        tcPr.append(bdr)
        for i, line in enumerate(text.strip().split('\n')):
            p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(12)
            r = p.add_run(line)
            r.font.name  = "Consolas"
            r.font.size  = Pt(8.5)
            r.font.color.rgb = CHARCOAL
        self.d.add_paragraph().paragraph_format.space_after = Pt(4)

    def toc_entry(self, tbl, idx, label, pg, is_major):
        row = tbl.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(5.6); c1.width = Inches(0.75)
        for c in (c0, c1):
            cell_margins(c, top=15, bottom=15, left=30, right=30)
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        p0.paragraph_format.line_spacing = Pt(13)
        self._run(p0, label, sz=11 if is_major else 10,
                  bold=is_major, color=NAVY if is_major else CHARCOAL)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_after = Pt(0)
        p1.paragraph_format.line_spacing = Pt(13)
        self._run(p1, pg, sz=11 if is_major else 10,
                  bold=is_major, color=NAVY if is_major else CHARCOAL)

    def lof_entry(self, tbl, idx, num, cap, pg):
        row = tbl.rows[idx]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width = Inches(0.9); c1.width = Inches(5.0); c2.width = Inches(0.6)
        for c in (c0, c1, c2):
            cell_margins(c, top=12, bottom=12, left=20, right=20)
        for cell, txt, align in [(c0, num, WD_ALIGN_PARAGRAPH.LEFT),
                                  (c1, cap, WD_ALIGN_PARAGRAPH.LEFT),
                                  (c2, pg,  WD_ALIGN_PARAGRAPH.RIGHT)]:
            p = cell.paragraphs[0]
            p.alignment = align
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(13)
            self._run(p, txt, sz=10, color=CHARCOAL)

    def save(self):
        self.d.save(str(OUT))
        shutil.copy2(str(OUT), str(DL))
        print(f"Saved DOCX: {OUT}  ({OUT.stat().st_size:,} bytes)")
        print(f"Copied to:  {DL}")

print("helpers.py loaded OK")
