"""
Joni Kumar Biswas — ATS-Optimized CV Generator
Requirements: pip install python-docx
Run:          python generate_cv.py
Output:       Joni_Kumar_Biswas_CV_ATS.docx
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── Colour Palette ────────────────────────────────────────────────────────────
NAVY  = RGBColor(0x1B, 0x3A, 0x6B)
DARK  = RGBColor(0x1A, 0x1A, 0x2E)
MID   = RGBColor(0x44, 0x44, 0x44)
LIGHT = RGBColor(0x66, 0x66, 0x66)


# ── Low-level XML helpers ─────────────────────────────────────────────────────

def set_paragraph_spacing(para, before_pt=0, after_pt=0, line_rule=None):
    """Set space-before / space-after in points (converted to twips internally)."""
    pPr = para._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), str(int(before_pt * 20)))
    spacing.set(qn("w:after"),  str(int(after_pt  * 20)))
    if line_rule:
        spacing.set(qn("w:lineRule"), line_rule)
    existing = pPr.find(qn("w:spacing"))
    if existing is not None:
        pPr.remove(existing)
    pPr.append(spacing)


def add_bottom_border(para, color_hex="1B3A6B", size=10, space=4):
    """Add a bottom border to a paragraph."""
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    str(size))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), color_hex)
    pBdr.append(bottom)
    existing = pPr.find(qn("w:pBdr"))
    if existing is not None:
        pPr.remove(existing)
    pPr.append(pBdr)


def add_top_bottom_border(para, color_hex="1B3A6B", size=6, space=2):
    """Add top + bottom borders to a paragraph."""
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    for side in ("top", "bottom"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   "single")
        el.set(qn("w:sz"),    str(size))
        el.set(qn("w:space"), str(space))
        el.set(qn("w:color"), color_hex)
        pBdr.append(el)
    existing = pPr.find(qn("w:pBdr"))
    if existing is not None:
        pPr.remove(existing)
    pPr.append(pBdr)


def add_right_tab(para, position_twips=9360):
    """Add a right-aligned tab stop at the given twips position."""
    pPr = para._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(position_twips))
    tabs.append(tab)
    existing = pPr.find(qn("w:tabs"))
    if existing is not None:
        pPr.remove(existing)
    pPr.append(tabs)


def add_hyperlink(para, url, display_text, color=NAVY, font_size=9.5):
    """Insert a clickable hyperlink run into an existing paragraph."""
    part = para.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    run_elem = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    # Font
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"),    "Calibri")
    rFonts.set(qn("w:hAnsi"),    "Calibri")
    rPr.append(rFonts)

    # Size
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(font_size * 2)))
    rPr.append(sz)

    # Color
    color_el = OxmlElement("w:color")
    color_el.set(qn("w:val"), f"{color[0]:02X}{color[1]:02X}{color[2]:02X}")
    rPr.append(color_el)

    # Underline
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)

    run_elem.append(rPr)
    t = OxmlElement("w:t")
    t.text = display_text
    run_elem.append(t)
    hyperlink.append(run_elem)
    para._p.append(hyperlink)


def set_bullet_indent(para):
    """Set hanging indent for bullet paragraphs."""
    pPr = para._p.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"),    "480")
    ind.set(qn("w:hanging"), "240")
    existing = pPr.find(qn("w:ind"))
    if existing is not None:
        pPr.remove(existing)
    pPr.append(ind)


def add_run(para, text, bold=False, italic=False, color=MID, font_size=10,
            font_name="Calibri"):
    """Add a formatted run to a paragraph and return it."""
    run = para.add_run(text)
    run.bold   = bold
    run.italic = italic
    run.font.name     = font_name
    run.font.size     = Pt(font_size)
    run.font.color.rgb = color
    return run


# ── High-level building blocks ────────────────────────────────────────────────

def section_header(doc, text):
    para = doc.add_paragraph()
    set_paragraph_spacing(para, before_pt=11, after_pt=4)
    add_bottom_border(para, color_hex="1B3A6B", size=10, space=4)
    run = para.add_run(text.upper())
    run.bold = True
    run.font.name       = "Calibri"
    run.font.size       = Pt(11)
    run.font.color.rgb  = NAVY
    # Character spacing (40/100 pt → in EMUs via XML)
    rPr = run._r.get_or_add_rPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:val"), "20")   # 20 = 1pt in half-points
    rPr.append(spacing)
    return para


def add_bullet(doc, sentence, bold_prefix=None):
    para = doc.add_paragraph()
    set_paragraph_spacing(para, before_pt=2, after_pt=2)
    set_bullet_indent(para)
    # bullet character
    run_bullet = para.add_run("▪  ")
    run_bullet.font.name      = "Calibri"
    run_bullet.font.size      = Pt(10)
    run_bullet.font.color.rgb = NAVY
    # optional bold label
    if bold_prefix:
        add_run(para, bold_prefix + " ", bold=True, color=DARK, font_size=10)
    add_run(para, sentence, bold=False, color=MID, font_size=10)
    return para


def add_job_header(doc, title, company, period, location):
    # Row 1: Job title (left) + Period (right)
    p1 = doc.add_paragraph()
    set_paragraph_spacing(p1, before_pt=8, after_pt=1)
    add_right_tab(p1)
    add_run(p1, title,  bold=True,  color=DARK,  font_size=11)
    add_run(p1, "\t",   bold=False, color=DARK,  font_size=10)
    add_run(p1, period, bold=False, italic=True, color=LIGHT, font_size=9.5)
    # Row 2: Company (left) + Location (right)
    p2 = doc.add_paragraph()
    set_paragraph_spacing(p2, before_pt=0, after_pt=3)
    add_right_tab(p2)
    add_run(p2, company,  bold=True,  color=NAVY,  font_size=10)
    add_run(p2, "\t",     bold=False, color=DARK,  font_size=10)
    add_run(p2, location, bold=False, italic=True, color=LIGHT, font_size=9.5)


def add_edu_row(doc, degree, institution, score, period):
    p1 = doc.add_paragraph()
    set_paragraph_spacing(p1, before_pt=8, after_pt=1)
    add_right_tab(p1)
    add_run(p1, degree,  bold=True,  color=DARK,  font_size=11)
    add_run(p1, "\t",    bold=False, color=DARK,  font_size=10)
    add_run(p1, period,  bold=False, italic=True, color=LIGHT, font_size=9.5)

    p2 = doc.add_paragraph()
    set_paragraph_spacing(p2, before_pt=0, after_pt=2)
    add_right_tab(p2)
    add_run(p2, institution, bold=True,  color=NAVY,  font_size=10)
    add_run(p2, "\t",        bold=False, color=DARK,  font_size=10)
    add_run(p2, score,       bold=False, italic=True, color=LIGHT, font_size=9.5)


def add_labeled_para(doc, label, value):
    para = doc.add_paragraph()
    set_paragraph_spacing(para, before_pt=3, after_pt=3)
    add_run(para, label + ": ", bold=True,  color=DARK, font_size=10)
    add_run(para, value,        bold=False, color=MID,  font_size=10)


# ── Main CV builder ───────────────────────────────────────────────────────────

def build_cv(output_path="Joni_Kumar_Biswas_CV_ATS.docx"):
    doc = Document()

    # Page margins (0.75 in top/bottom, ~0.9 in left/right)
    for section in doc.sections:
        section.top_margin    = Inches(0.625)
        section.bottom_margin = Inches(0.625)
        section.left_margin   = Inches(0.75)
        section.right_margin  = Inches(0.75)

    # Remove default paragraph spacing from Normal style
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after  = Pt(0)

    # ── NAME ─────────────────────────────────────────────────────────────────
    name_para = doc.add_paragraph()
    name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(name_para, before_pt=0, after_pt=3)
    run = name_para.add_run("JONI KUMAR BISWAS")
    run.bold = True
    run.font.name       = "Calibri"
    run.font.size       = Pt(26)
    run.font.color.rgb  = NAVY
    # Letter spacing via XML
    rPr = run._r.get_or_add_rPr()
    sp  = OxmlElement("w:spacing")
    sp.set(qn("w:val"), "40")
    rPr.append(sp)

    # ── SUB-TITLE ─────────────────────────────────────────────────────────────
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(title_para, before_pt=0, after_pt=6)
    add_run(title_para,
            "Finance & Accounting Professional  |  Business Development",
            italic=True, color=MID, font_size=11)

    # ── CONTACT BAR ──────────────────────────────────────────────────────────
    contact = doc.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(contact, before_pt=4, after_pt=4)
    add_top_bottom_border(contact, color_hex="1B3A6B", size=6, space=2)

    sep = "   |   "
    add_run(contact, "Dhaka, Bangladesh",        color=MID,  font_size=9.5)
    add_run(contact, sep,                        color=NAVY, font_size=9.5)
    add_run(contact, "+880 1875049536",           color=MID,  font_size=9.5)
    add_run(contact, sep,                        color=NAVY, font_size=9.5)
    add_run(contact, "jonikumarbiswas@gmail.com", color=MID,  font_size=9.5)
    add_run(contact, sep,                        color=NAVY, font_size=9.5)
    add_hyperlink(contact,
                  "https://www.linkedin.com/in/joni-kumar-biswas-1ba7a3349/",
                  "LinkedIn", color=NAVY, font_size=9.5)
    add_run(contact, sep, color=NAVY, font_size=9.5)
    add_hyperlink(contact,
                  "https://github.com/jonikumarbiswas-create",
                  "GitHub", color=NAVY, font_size=9.5)

    # ── PROFESSIONAL SUMMARY ─────────────────────────────────────────────────
    section_header(doc, "Professional Summary")
    summary = doc.add_paragraph()
    set_paragraph_spacing(summary, before_pt=4, after_pt=4)
    add_run(summary,
            "Results-driven BBA graduate with a Major in Accounting and hands-on experience "
            "spanning corporate banking, financial operations, and digital business development. "
            "Proven ability to manage financial records, reconcile accounts, and drive client "
            "acquisition in fast-paced, multicultural environments. Committed to delivering "
            "accuracy, compliance, and operational excellence—attributes sought by multinational "
            "organizations across finance and business functions.",
            color=MID, font_size=10)

    # ── PROFESSIONAL EXPERIENCE ──────────────────────────────────────────────
    section_header(doc, "Professional Experience")

    # Job 1 — Vnture AI
    add_job_header(doc,
                   "Business Development Executive",
                   "Vnture AI",
                   "Feb 2026 – Present",
                   "Dhaka, Bangladesh")
    add_bullet(doc,
               "Managed the company's Fiverr marketplace profile with optimized service listings "
               "to boost platform visibility and inbound lead volume.",
               "Profile & Listing Management:")
    add_bullet(doc,
               "Communicated with international clients across diverse markets, converting inbound "
               "inquiries into billable projects through consultative selling.",
               "Client Acquisition:")
    add_bullet(doc,
               "Coordinated cross-functional teams across design, development, and delivery to "
               "ensure on-time project completion and high client satisfaction.",
               "Project Coordination:")
    add_bullet(doc,
               "Tracked key business metrics including lead conversion rate, client retention, and "
               "project pipeline to inform strategic decisions.",
               "Performance Reporting:")

    # Job 2 — Authentic Software
    add_job_header(doc,
                   "Junior Accountant",
                   "Authentic Software",
                   "Jan 2025 – Aug 2025",
                   "Dhaka, Bangladesh")
    add_bullet(doc,
               "Managed daily financial transactions, ensuring accurate ledger maintenance, "
               "invoicing for software services, and tracking of operational expenses.",
               "Financial Reporting & Bookkeeping:")
    add_bullet(doc,
               "Conducted monthly bank and account reconciliations, identifying discrepancies and "
               "ensuring 100% financial accuracy.",
               "Account Reconciliation:")
    add_bullet(doc,
               "Maintained a high standard of professional integrity and diligent work ethic, "
               "ensuring all financial records complied with internal company policies and "
               "accounting standards.",
               "Integrity & Compliance:")
    add_bullet(doc,
               "Collaborated with the broader team to process payroll, manage accounts "
               "payable/receivable, and support financial planning for ongoing web and software "
               "projects.",
               "Cross-Departmental Support:")

    # Job 3 — Dhaka Bank
    add_job_header(doc,
                   "Intern – Finance & Banking Operations",
                   "Dhaka Bank PLC",
                   "Feb 2024 – Jul 2024",
                   "Dhaka, Bangladesh")
    add_bullet(doc,
               "Supported high-volume daily transactions, processing financial documentation in "
               "compliance with regulatory standards.",
               "Banking Operations:")
    add_bullet(doc,
               "Assisted front-line teams with client queries, gaining practical knowledge of KYC "
               "protocols and corporate banking procedures.",
               "Customer Service:")
    add_bullet(doc,
               "Collaborated with senior relationship managers on file reviews, accelerating "
               "turnaround times for client requests.",
               "Documentation Support:")

    # ── CORE COMPETENCIES ────────────────────────────────────────────────────
    section_header(doc, "Core Competencies")
    add_labeled_para(doc,
                     "Financial & Accounting Operations",
                     "Financial record keeping, bookkeeping, monthly reconciliations, accounts "
                     "payable/receivable, data entry, QuickBooks")
    add_labeled_para(doc,
                     "Technical Proficiency",
                     "MS Excel (advanced formulas, pivot tables, data analysis), MS Word, "
                     "MS PowerPoint, QuickBooks, Odoo ERP (financial modules, data entry, reporting)")
    add_labeled_para(doc,
                     "Business Development",
                     "Lead generation, inbound lead conversion, international client communication, "
                     "marketplace profile optimization, cross-functional coordination")
    add_labeled_para(doc,
                     "Professional Attributes",
                     "Analytical thinking, numerical accuracy, attention to detail, strict "
                     "confidentiality, cross-functional collaboration, adaptability in multicultural "
                     "environments")
    add_labeled_para(doc,
                     "Digital Marketing",
                     "Search Engine Optimization (SEO), digital campaign management, content strategy")

    # ── EDUCATION ────────────────────────────────────────────────────────────
    section_header(doc, "Education")
    add_edu_row(doc,
                "Bachelor of Business Administration — Major: Accounting",
                "Daffodil International University, Dhaka",
                "CGPA: 3.74 / 4.00",
                "2020 – 2024")
    add_edu_row(doc,
                "Higher Secondary Certificate — Business Studies",
                "Halishahar Cantonment Public School & College, Chittagong",
                "GPA: 4.50 / 5.00",
                "2017 – 2019")

    # ── CERTIFICATIONS ───────────────────────────────────────────────────────
    section_header(doc, "Certifications & Training")
    add_bullet(doc,
               "Advanced MS Excel for Data Analysis — BITM "
               "(Bangladesh Institute of Technology Management)")
    add_bullet(doc,
               "Digital Marketing: SEO & Campaign Management — SR DreamIT & Ghuri Learning")

    # ── LEADERSHIP & ACTIVITIES ──────────────────────────────────────────────
    section_header(doc, "Leadership & Activities")
    add_bullet(doc,
               "General Secretary & Technical Supporter — Business Club, Daffodil International "
               "University; organized flagship university events including Job Utshob and Annual "
               "Convocation.")
    add_bullet(doc,
               "2nd Place — University Business Quiz Competition, demonstrating competitive "
               "analytical and strategic thinking.")

    # ── KEY SKILLS (ATS keywords) ─────────────────────────────────────────────
    section_header(doc, "Key Skills")
    ks = doc.add_paragraph()
    set_paragraph_spacing(ks, before_pt=4, after_pt=4)
    add_run(ks,
            "Financial Reporting  |  Account Reconciliation  |  Bookkeeping  |  QuickBooks  |  "
            "Odoo ERP  |  Data Entry  |  MS Excel  |  Data Analysis  |  Business Development  |  "
            "Client Relationship Management  |  Banking Operations  |  KYC & Compliance  |  "
            "Digital Marketing  |  SEO  |  Cross-functional Collaboration  |  Attention to Detail",
            color=MID, font_size=9.5)

    # ── Save ──────────────────────────────────────────────────────────────────
    doc.save(output_path)
    print(f"✅  CV saved → {output_path}")


if __name__ == "__main__":
    build_cv("Joni_Kumar_Biswas_CV_ATS.docx")
