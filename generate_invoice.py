"""
Boring Science LLC — Invoice DOCX Generator
Run: python3 generate_invoice.py
Outputs: invoice-template.docx
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── BRAND COLOURS ──────────────────────────────────────────────
BG_DARK    = RGBColor(0x0a, 0x0a, 0x0a)
BG_CARD    = RGBColor(0x16, 0x16, 0x16)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT     = RGBColor(0xC8, 0xFF, 0x00)
MUTED      = RGBColor(0x88, 0x88, 0x88)
DIM        = RGBColor(0xC8, 0xC8, 0xC8)
BORDER     = RGBColor(0x22, 0x22, 0x22)
BLACK      = RGBColor(0x00, 0x00, 0x00)

MONO  = "Courier New"   # fallback for JetBrains Mono
SANS  = "Calibri"       # fallback for DM Sans
SERIF = "Georgia"       # fallback for Instrument Serif

# ── HELPERS ────────────────────────────────────────────────────
def set_cell_bg(cell, rgb: RGBColor):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), str(rgb))
    tcPr.append(shd)

def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), val.get('val', 'single'))
            el.set(qn('w:sz'), str(val.get('sz', 4)))
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), val.get('color', 'FFFFFF'))
            tcBorders.append(el)
    tcPr.append(tcBorders)

def set_table_no_border(table):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'none')
        tblBorders.append(el)
    tblPr.append(tblBorders)

def run(para, text, bold=False, font=SANS, size=10, color=WHITE, italic=False):
    r = para.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.name = font
    r.font.size = Pt(size)
    r.font.color.rgb = color
    return r

def label_para(cell_or_doc, text):
    """Accent uppercase mono label"""
    if hasattr(cell_or_doc, 'paragraphs'):
        p = cell_or_doc.add_paragraph()
    else:
        p = cell_or_doc
    p.clear()
    r = p.add_run(text.upper())
    r.font.name = MONO
    r.font.size = Pt(7.5)
    r.font.color.rgb = ACCENT
    r.font.bold = True
    p.paragraph_format.space_after = Pt(3)
    return p

def value_para(cell, text, size=10, color=WHITE, font=SANS, bold=False, align=None):
    p = cell.add_paragraph()
    r = p.add_run(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.bold = bold
    p.paragraph_format.space_after = Pt(2)
    if align:
        p.alignment = align
    return p

def set_col_width(table, col_idx, width_cm):
    for row in table.rows:
        row.cells[col_idx].width = Cm(width_cm)

def add_rule(doc, color=BORDER):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), str(color))
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p

# ── BUILD DOC ──────────────────────────────────────────────────
doc = Document()

# Page setup — A4, tight margins
section = doc.sections[0]
section.page_width  = Cm(21.0)
section.page_height = Cm(29.7)
section.top_margin    = Cm(1.5)
section.bottom_margin = Cm(1.5)
section.left_margin   = Cm(1.8)
section.right_margin  = Cm(1.8)

# Set page background to dark
# (Word needs a fill on the body — we do it via doc background XML)
background = OxmlElement('w:background')
background.set(qn('w:color'), '0A0A0A')
doc.element.insert(0, background)
settings = doc.settings.element
disp = OxmlElement('w:displayBackgroundShape')
settings.insert(0, disp)

# ── HEADER TABLE ────────────────────────────────────────────────
ht = doc.add_table(rows=1, cols=2)
set_table_no_border(ht)
ht.alignment = WD_TABLE_ALIGNMENT.CENTER
ht.columns[0].width = Cm(11)
ht.columns[1].width = Cm(7)

left  = ht.cell(0, 0)
right = ht.cell(0, 1)
set_cell_bg(left,  BG_DARK)
set_cell_bg(right, BG_DARK)
left.vertical_alignment  = WD_ALIGN_VERTICAL.BOTTOM
right.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM

# Brand name — left
brand_p = left.add_paragraph()
run(brand_p, 'BORING ', bold=True, font=MONO, size=18, color=WHITE)
run(brand_p, 'SCIENCE', bold=True, font=MONO, size=18, color=ACCENT)
brand_p.paragraph_format.space_after = Pt(2)

sub_p = left.add_paragraph()
run(sub_p, 'boringscience.bio  ·  Sunnyvale, CA', font=MONO, size=7.5, color=MUTED)
sub_p.paragraph_format.space_after = Pt(0)

# Invoice number — right
right.add_paragraph()  # spacer
lbl_p = right.add_paragraph()
lbl_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run(lbl_p, 'INVOICE', font=MONO, size=7.5, color=ACCENT, bold=True)
lbl_p.paragraph_format.space_after = Pt(2)

num_p = right.add_paragraph()
num_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run(num_p, '#INV-2026-001', bold=True, font=MONO, size=16, color=WHITE)
num_p.paragraph_format.space_after = Pt(0)

add_rule(doc, BORDER)

# ── PARTIES TABLE ────────────────────────────────────────────────
pt = doc.add_table(rows=1, cols=2)
set_table_no_border(pt)
pt.alignment = WD_TABLE_ALIGNMENT.CENTER

fc = pt.cell(0, 0)
tc = pt.cell(0, 1)
set_cell_bg(fc, BG_DARK)
set_cell_bg(tc, BG_DARK)

label_para(fc, 'From')
value_para(fc, 'Boring Science LLC', size=12, bold=True)
value_para(fc, '1234 Innovation Drive, Suite 100', size=9, color=MUTED)
value_para(fc, 'Sunnyvale, CA 94086  ·  United States', size=9, color=MUTED)
value_para(fc, 'inquiries@boringscience.bio  ·  408.368.9547', size=9, color=MUTED)

label_para(tc, 'Bill To')
value_para(tc, 'Client Name / Organization', size=12, bold=True)
value_para(tc, 'Street Address', size=9, color=MUTED)
value_para(tc, 'City, State ZIP  ·  Country', size=9, color=MUTED)
value_para(tc, 'client@example.com', size=9, color=MUTED)

p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6)
set_cell_bg(p._p, BG_DARK) if hasattr(p._p, '_tc') else None

# ── DATES TABLE ─────────────────────────────────────────────────
dt = doc.add_table(rows=2, cols=4)
set_table_no_border(dt)
dt.alignment = WD_TABLE_ALIGNMENT.CENTER

dates = [
    ('Issue Date',      '2026-03-12',    WHITE),
    ('Due Date',        '2026-04-11',    ACCENT),
    ('Project / PO Ref','BS-PROJECT-001', WHITE),
    ('Terms',           'Net 30',        WHITE),
]

for i, (lbl, val, col) in enumerate(dates):
    c = dt.cell(0, i)
    set_cell_bg(c, BG_CARD)
    lp = c.add_paragraph()
    run(lp, lbl.upper(), font=MONO, size=7, color=MUTED)
    lp.paragraph_format.space_after = Pt(2)

    c2 = dt.cell(1, i)
    set_cell_bg(c2, BG_CARD)
    vp = c2.add_paragraph()
    run(vp, val, font=MONO, size=10, color=col, bold=(col == ACCENT))
    vp.paragraph_format.space_after = Pt(0)

# padding rows
for row in dt.rows:
    for cell in row.cells:
        cell.width = Cm(4.5)

p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(8)

# ── LINE ITEMS ───────────────────────────────────────────────────
lbl_p2 = doc.add_paragraph()
run(lbl_p2, 'SERVICES RENDERED', font=MONO, size=7.5, color=ACCENT, bold=True)
lbl_p2.paragraph_format.space_after = Pt(4)

items_table = doc.add_table(rows=1, cols=4)
set_table_no_border(items_table)
items_table.alignment = WD_TABLE_ALIGNMENT.CENTER

col_widths = [9.0, 2.8, 3.0, 3.0]
headers = ['Description', 'Qty / Hrs', 'Rate (USD)', 'Amount (USD)']
header_row = items_table.rows[0]
for i, (hdr, w) in enumerate(zip(headers, col_widths)):
    c = header_row.cells[i]
    set_cell_bg(c, BG_DARK)
    set_cell_border(c, bottom={'val': 'single', 'sz': 4, 'color': '222222'})
    c.width = Cm(w)
    hp = c.add_paragraph()
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i > 0 else WD_ALIGN_PARAGRAPH.LEFT
    run(hp, hdr.upper(), font=MONO, size=7, color=MUTED)
    hp.paragraph_format.space_after = Pt(4)

line_items = [
    ('Computational Biology Consulting',
     'SE(3)-GNN docking pipeline setup and validation',
     '16', '$250.00', '$4,000.00'),
    ('NGS Pipeline Engineering',
     'Nextflow DSL2 WGS pipeline — containerization & CI',
     '24', '$250.00', '$6,000.00'),
    ('bsdock License — Annual',
     'Single-site, unlimited users. Includes support & updates.',
     '1', '$3,500.00', '$3,500.00'),
    ('Data Report & Documentation',
     'Technical report, methodology writeup, reproducibility package',
     '8', '$200.00', '$1,600.00'),
]

for desc, sub, qty, rate, amt in line_items:
    row = items_table.add_row()
    cells = row.cells
    for i, c in enumerate(cells):
        set_cell_bg(c, BG_DARK)
        set_cell_border(c, bottom={'val': 'single', 'sz': 4, 'color': '1a1a1a'})
        c.width = Cm(col_widths[i])
        c.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # Description cell
    dp = cells[0].add_paragraph()
    run(dp, desc, font=SANS, size=10, color=WHITE, bold=True)
    dp.paragraph_format.space_after = Pt(1)
    sp = cells[0].add_paragraph()
    run(sp, sub, font=SANS, size=8.5, color=MUTED)
    sp.paragraph_format.space_after = Pt(6)

    for i, val in enumerate([qty, rate, amt], 1):
        vp = cells[i].add_paragraph()
        vp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run(vp, val, font=MONO, size=9, color=WHITE if i == 3 else DIM, bold=(i==3))
        vp.paragraph_format.space_after = Pt(6)

p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)

# ── TOTALS ───────────────────────────────────────────────────────
tot_table = doc.add_table(rows=4, cols=2)
set_table_no_border(tot_table)
tot_table.alignment = WD_TABLE_ALIGNMENT.RIGHT

totals = [
    ('Subtotal',   '$15,100.00', DIM,   False),
    ('Tax (0%)',   '$0.00',      DIM,   False),
    ('Discount',   '—',          DIM,   False),
    ('Total Due',  '$15,100.00', WHITE, True),
]

tot_col_w = [4.5, 3.0]
for i, (lbl, val, col, bold) in enumerate(totals):
    lc = tot_table.rows[i].cells[0]
    vc = tot_table.rows[i].cells[1]
    lc.width = Cm(tot_col_w[0])
    vc.width = Cm(tot_col_w[1])
    set_cell_bg(lc, BG_DARK)
    set_cell_bg(vc, BG_DARK)

    lbl_p3 = lc.add_paragraph()
    run(lbl_p3, lbl.upper(), font=MONO, size=7.5 if not bold else 8,
        color=ACCENT if bold else MUTED)
    lbl_p3.paragraph_format.space_after = Pt(3)

    val_p = vc.add_paragraph()
    val_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run(val_p, val, font=MONO, size=13 if bold else 9,
        color=col, bold=bold)
    val_p.paragraph_format.space_after = Pt(3)

    if i == 2:  # line above grand total
        set_cell_border(lc, bottom={'val': 'single', 'sz': 4, 'color': '222222'})
        set_cell_border(vc, bottom={'val': 'single', 'sz': 4, 'color': '222222'})

p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6)

# ── PAYMENT ──────────────────────────────────────────────────────
pay_outer = doc.add_table(rows=1, cols=1)
set_table_no_border(pay_outer)
pay_outer.alignment = WD_TABLE_ALIGNMENT.CENTER
pay_cell = pay_outer.cell(0, 0)
set_cell_bg(pay_cell, BG_CARD)
set_cell_border(pay_cell,
    top   ={'val':'single','sz':4,'color':'222222'},
    bottom={'val':'single','sz':4,'color':'222222'},
    right ={'val':'single','sz':4,'color':'222222'},
    left  ={'val':'single','sz':12,'color':'C8FF00'})

label_para(pay_cell, 'Payment Instructions')

pay_grid = [
    ('Bank',           'Bank Name'),
    ('Account Name',   'Boring Science LLC'),
    ('Account No.',    'XXXX-XXXX-XXXX'),
    ('Routing/SWIFT',  'XXXXXXXXX'),
    ('Wire Reference', '#INV-2026-001'),
    ('Also Accepts',   'ACH · Check · Stripe'),
]

inner = pay_cell.add_table(rows=3, cols=4)
set_table_no_border(inner)
for row_i in range(3):
    for col_i in range(4):
        idx = row_i * 2 + col_i // 2
        if idx >= len(pay_grid): break
        c = inner.rows[row_i].cells[col_i]
        set_cell_bg(c, BG_CARD)
        if col_i % 2 == 0:
            lp = c.add_paragraph()
            run(lp, pay_grid[idx][0].upper(), font=MONO, size=7, color=MUTED)
            lp.paragraph_format.space_after = Pt(1)
        else:
            vp = c.add_paragraph()
            run(vp, pay_grid[idx][1], font=MONO, size=9, color=WHITE)
            vp.paragraph_format.space_after = Pt(4)

p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6)

# ── NOTES ────────────────────────────────────────────────────────
lbl_notes = doc.add_paragraph()
run(lbl_notes, 'NOTES', font=MONO, size=7.5, color=ACCENT, bold=True)
lbl_notes.paragraph_format.space_after = Pt(3)

notes_p = doc.add_paragraph()
run(notes_p,
    'Payment is due within 30 days of the invoice date. '
    'Late payments are subject to a 1.5% monthly interest charge. '
    'Please include the invoice number in your payment reference. '
    'For questions, contact inquiries@boringscience.bio.',
    font=SANS, size=8.5, color=MUTED)
notes_p.paragraph_format.space_after = Pt(10)

add_rule(doc, BORDER)

# ── FOOTER ───────────────────────────────────────────────────────
ft = doc.add_table(rows=1, cols=2)
set_table_no_border(ft)
fl = ft.cell(0, 0)
fr = ft.cell(0, 1)
set_cell_bg(fl, BG_DARK)
set_cell_bg(fr, BG_DARK)

lp2 = fl.add_paragraph()
run(lp2, 'Boring Science LLC  ·  Incorporated 2026  ·  Sunnyvale, CA  ·  boringscience.bio',
    font=MONO, size=7.5, color=MUTED)

rp2 = fr.add_paragraph()
rp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run(rp2, "Science is unpredictable. Your invoices shouldn't be.",
    font=SERIF, size=9.5, color=MUTED, italic=True)

# ── SAVE ─────────────────────────────────────────────────────────
out = 'invoice-template.docx'
doc.save(out)
print(f'✓  Saved: {out}')
