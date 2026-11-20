import os
import sys
import shutil
from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=60, bottom=60, left=90, right=90):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="CBD5E1", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

class ReportBuilder:
    def __init__(self):
        self.doc = Document()
        self.NAVY = RGBColor(15, 51, 71)        # #0F3347
        self.SAFFRON = RGBColor(192, 80, 32)    # #C05020
        self.GOLD = RGBColor(156, 100, 18)      # #9C6412
        self.CHARCOAL = RGBColor(30, 41, 59)    # #1E293B
        self.GRAY = RGBColor(100, 116, 139)     # #64748B
        self.BLACK = RGBColor(0, 0, 0)
        
        # Section 1: Front Matter (Roman numerals)
        sec1 = self.doc.sections[0]
        sec1.top_margin = Inches(0.80)
        sec1.bottom_margin = Inches(0.80)
        sec1.left_margin = Inches(1.10)
        sec1.right_margin = Inches(0.80)
        sec1.different_first_page_header_footer = True
        
        # Format Roman numerals for front matter
        sectPr1 = sec1._sectPr
        pgNumType1 = OxmlElement('w:pgNumType')
        pgNumType1.set(qn('w:fmt'), 'lowerRoman')
        sectPr1.append(pgNumType1)

        # Section 1 Footer (center Roman page number)
        ftr1 = sec1.footer
        p_ftr1 = ftr1.paragraphs[0]
        p_ftr1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ftr1.paragraph_format.space_before = Pt(0)
        p_ftr1.paragraph_format.space_after = Pt(0)
        fldSimple1 = OxmlElement('w:fldSimple')
        fldSimple1.set(qn('w:instr'), 'PAGE')
        p_ftr1._p.append(fldSimple1)

    def start_main_body(self):
        # Section 2: Main Chapters (Arabic numerals starting at 1)
        sec2 = self.doc.add_section(WD_SECTION.NEW_PAGE)
        sec2.top_margin = Inches(0.80)
        sec2.bottom_margin = Inches(0.80)
        sec2.left_margin = Inches(1.10)
        sec2.right_margin = Inches(0.80)
        
        sec2.different_first_page_header_footer = False
        sec2.header.is_linked_to_previous = False
        sec2.footer.is_linked_to_previous = False

        sectPr2 = sec2._sectPr
        pgNumType2 = OxmlElement('w:pgNumType')
        pgNumType2.set(qn('w:start'), '1')
        pgNumType2.set(qn('w:fmt'), 'decimal')
        sectPr2.append(pgNumType2)

        # Header for Section 2
        hdr2 = sec2.header
        p_hdr = hdr2.paragraphs[0]
        p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_hdr.paragraph_format.space_after = Pt(2)
        r_hdr = p_hdr.add_run("Summer Internship Report | DeccanPulse: Hyderabad Smart Parking Platform")
        r_hdr.font.name = "Times New Roman"
        r_hdr.font.size = Pt(8.5)
        r_hdr.font.italic = True
        r_hdr.font.color.rgb = self.GRAY

        # Footer for Section 2
        ftr2 = sec2.footer
        p_ftr = ftr2.paragraphs[0]
        p_ftr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_ftr.paragraph_format.space_before = Pt(2)
        r_ftr = p_ftr.add_run("Matrusri Engineering College | Dept. of Computer Science & Engineering  •  Page ")
        r_ftr.font.name = "Times New Roman"
        r_ftr.font.size = Pt(8.5)
        r_ftr.font.color.rgb = self.GRAY
        
        fldSimple2 = OxmlElement('w:fldSimple')
        fldSimple2.set(qn('w:instr'), 'PAGE')
        p_ftr._p.append(fldSimple2)

    def add_h1(self, text, space_before=10, space_after=4):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = self.NAVY
        return p

    def add_h2(self, text, space_before=7, space_after=3):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(11.5)
        r.font.bold = True
        r.font.color.rgb = self.SAFFRON
        return p

    def add_h3(self, text, space_before=5, space_after=2):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = self.CHARCOAL
        return p

    def add_p(self, text, bold_prefix=None, space_after=3.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = self.doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.12
        if bold_prefix:
            r_b = p.add_run(bold_prefix)
            r_b.font.name = "Times New Roman"
            r_b.font.size = Pt(9.5)
            r_b.font.bold = True
            r_b.font.color.rgb = self.CHARCOAL
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(9.5)
        r.font.color.rgb = self.CHARCOAL
        return p

    def add_bullet(self, text, bold_prefix=None):
        p = self.doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2.5)
        p.paragraph_format.line_spacing = 1.10
        if bold_prefix:
            r_b = p.add_run(bold_prefix)
            r_b.font.name = "Times New Roman"
            r_b.font.size = Pt(9.5)
            r_b.font.bold = True
            r_b.font.color.rgb = self.CHARCOAL
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(9.5)
        r.font.color.rgb = self.CHARCOAL
        return p

    def add_callout(self, text, title="TECHNICAL ARCHITECTURE NOTE"):
        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        tr = table.rows[0]._tr.get_or_add_trPr()
        tr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        cell = table.cell(0, 0)
        cell.width = Inches(6.6)
        set_cell_background(cell, "F0F9FF")
        set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
        
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="BAE6FD"/>'
            f'  <w:left w:val="single" w:sz="18" w:space="0" w:color="0284C7"/>'
            f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="BAE6FD"/>'
            f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="BAE6FD"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)

        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r_t = p.add_run(f"✦  {title.upper()}: ")
        r_t.font.name = "Times New Roman"
        r_t.font.size = Pt(9.0)
        r_t.font.bold = True
        r_t.font.color.rgb = RGBColor(3, 105, 161)
        
        r_body = p.add_run(text)
        r_body.font.name = "Times New Roman"
        r_body.font.size = Pt(9.0)
        r_body.font.color.rgb = RGBColor(15, 23, 42)

    def add_image(self, img_path, caption, width=Inches(4.3)):
        if os.path.exists(img_path):
            p_img = self.doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(3)
            p_img.paragraph_format.space_after = Pt(2)
            p_img.paragraph_format.keep_with_next = True
            r_img = p_img.add_run()
            r_img.add_picture(img_path, width=width)
            
            p_cap = self.doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(0)
            p_cap.paragraph_format.space_after = Pt(3)
            r_cap = p_cap.add_run(caption)
            r_cap.font.name = "Times New Roman"
            r_cap.font.size = Pt(8.5)
            r_cap.font.bold = True
            r_cap.font.color.rgb = self.GRAY
        else:
            self.add_callout(f"IMAGE NOT FOUND AT: {img_path}", title="MISSING ASSET")

    def add_paired_images(self, img1_path, cap1, img2_path, cap2, width=Inches(4.3)):
        self.add_image(img1_path, cap1, width=width)
        self.add_image(img2_path, cap2, width=width)

    def add_table(self, headers, data, col_widths, caption=None):
        if caption:
            p_cap = self.doc.add_paragraph()
            p_cap.paragraph_format.space_before = Pt(5)
            p_cap.paragraph_format.space_after = Pt(2)
            p_cap.paragraph_format.keep_with_next = True
            r_cap = p_cap.add_run(caption)
            r_cap.font.name = "Times New Roman"
            r_cap.font.size = Pt(8.5)
            r_cap.font.bold = True
            r_cap.font.color.rgb = self.NAVY

        table = self.doc.add_table(rows=len(data) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table, color="CBD5E1", sz="4")

        # Header Row
        hdr_row = table.rows[0]
        hdr_tr = hdr_row._tr.get_or_add_trPr()
        hdr_tr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        hdr_tr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        for i, h_text in enumerate(headers):
            cell = hdr_row.cells[i]
            cell.width = Inches(col_widths[i])
            set_cell_background(cell, "0F3347")
            set_cell_margins(cell, top=45, bottom=45, left=80, right=80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(h_text)
            r.font.name = "Times New Roman"
            r.font.size = Pt(8.5)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

        # Data Rows
        for r_idx, row_data in enumerate(data):
            row = table.rows[r_idx + 1]
            r_tr = row._tr.get_or_add_trPr()
            r_tr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            bg_color = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
            for c_idx, cell_value in enumerate(row_data):
                cell = row.cells[c_idx]
                cell.width = Inches(col_widths[c_idx])
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, top=35, bottom=35, left=80, right=80)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.02
                r = p.add_run(str(cell_value))
                r.font.name = "Times New Roman"
                r.font.size = Pt(8.0)
                r.font.color.rgb = self.CHARCOAL

    def add_code_listing(self, code_text, caption):
        p_cap = self.doc.add_paragraph()
        p_cap.paragraph_format.space_before = Pt(5)
        p_cap.paragraph_format.space_after = Pt(2)
        p_cap.paragraph_format.keep_with_next = True
        r_cap = p_cap.add_run(caption)
        r_cap.font.name = "Times New Roman"
        r_cap.font.size = Pt(8.5)
        r_cap.font.bold = True
        r_cap.font.color.rgb = self.NAVY

        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        tr = table.rows[0]._tr.get_or_add_trPr()
        tr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        cell = table.cell(0, 0)
        cell.width = Inches(6.6)
        set_cell_background(cell, "F8FAFC")
        set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
        
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            f'  <w:left w:val="single" w:sz="16" w:space="0" w:color="0F3347"/>'
            f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)

        lines = code_text.strip().split("\n")
        for i, line in enumerate(lines):
            p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run(line)
            r.font.name = "Consolas"
            r.font.size = Pt(7.5)
            r.font.color.rgb = self.CHARCOAL

    def add_page_break(self):
        self.doc.add_page_break()

print("ReportBuilder successfully updated.")
