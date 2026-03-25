"""
Générateur de CV — Conseiller·ère en Insertion Professionnelle
Louiza Hadid — Mission Locale du Pays de Lorient
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT   = "Calibri"
BLACK  = RGBColor(0, 0, 0)
GRAY   = RGBColor(110, 110, 110)
OUTPUT = "/home/user/Lettre-de-motivation/CV_Louiza_Hadid_CIP.docx"


# ─── Helpers ────────────────────────────────────────────────────────────────

def set_run(run, size=10.5, bold=False, italic=False, color=None):
    run.font.name      = FONT
    run.font.size      = Pt(size)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.color.rgb = color if color else BLACK


def _clear_borders(element, tag):
    bd = OxmlElement(tag)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"),   "none")
        b.set(qn("w:sz"),    "0")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        bd.append(b)
    ex = element.find(qn(tag))
    if ex is not None:
        element.remove(ex)
    element.append(bd)


def borderless_table(doc, rows, cols):
    tbl = doc.add_table(rows=rows, cols=cols)
    tbl.style = "Normal Table"
    tblPr = tbl._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl._tbl.insert(0, tblPr)
    _clear_borders(tblPr, "w:tblBorders")
    for row in tbl.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            _clear_borders(tcPr, "w:tcBorders")
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            for p in list(cell.paragraphs):
                p._element.getparent().remove(p._element)
    return tbl


def set_col_width(cell, cm):
    tcPr = cell._tc.get_or_add_tcPr()
    w = OxmlElement("w:tcW")
    w.set(qn("w:w"),    str(int(cm * 567)))
    w.set(qn("w:type"), "dxa")
    ex = tcPr.find(qn("w:tcW"))
    if ex is not None:
        tcPr.remove(ex)
    tcPr.append(w)


def cell_para(cell, text="", bold=False, italic=False, size=10.5,
              align=WD_ALIGN_PARAGRAPH.LEFT, sb=0, sa=2, color=None):
    p = cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    if text:
        r = p.add_run(text)
        set_run(r, size=size, bold=bold, italic=italic,
                color=color if color else BLACK)
    return p


def section_title(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(3)
    r = p.add_run(title.upper())
    set_run(r, size=10.5, bold=True)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)


def bullet(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before  = Pt(0)
    p.paragraph_format.space_after   = Pt(2)
    p.paragraph_format.left_indent   = Cm(0.5)
    r = p.add_run(f"\u2013 {text}")
    set_run(r, size=10)
    return p


def exp_header(doc, poste, structure, lieu, dates):
    """Ligne : poste (gauche, gras) | dates (droite, gris) + structure sous."""
    t = borderless_table(doc, 1, 2)
    cl = t.cell(0, 0)
    cr = t.cell(0, 1)
    set_col_width(cl, 11.5)
    set_col_width(cr, 5.0)

    p = cl.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after  = Pt(1)
    set_run(p.add_run(poste), size=10.5, bold=True)

    p2 = cl.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after  = Pt(2)
    set_run(p2.add_run(f"{structure} \u2014 {lieu}"),
            size=10, italic=True, color=GRAY)

    pd = cr.add_paragraph()
    pd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pd.paragraph_format.space_before = Pt(7)
    pd.paragraph_format.space_after  = Pt(1)
    set_run(pd.add_run(dates), size=10, color=GRAY)


# ═══════════════════════════════════════════════════════════════════════════
#  Document
# ═══════════════════════════════════════════════════════════════════════════

doc = Document()
sec = doc.sections[0]
sec.top_margin    = Cm(1.8)
sec.bottom_margin = Cm(1.8)
sec.left_margin   = Cm(2.2)
sec.right_margin  = Cm(2.2)

normal = doc.styles["Normal"]
normal.font.name = FONT
normal.font.size = Pt(10.5)
normal.paragraph_format.space_before = Pt(0)
normal.paragraph_format.space_after  = Pt(0)


# ═══════════════════════════════════════════════════════════════════════════
#  1. EN-TÊTE
# ═══════════════════════════════════════════════════════════════════════════

t = borderless_table(doc, 1, 2)
cl = t.cell(0, 0)
cr = t.cell(0, 1)
set_col_width(cl, 10.5)
set_col_width(cr, 6.0)

p = cl.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(3)
set_run(p.add_run("Louiza Hadid"), size=22, bold=True)

cell_para(cl, "Conseillère en insertion professionnelle",
          size=11, italic=True, color=GRAY, sa=0)

for line in ["06 50 37 56 47", "louizahdd@gmail.com",
             "3 rue de Ventspils \u2014 56100 Lorient", "Permis B"]:
    cell_para(cr, line, size=10, align=WD_ALIGN_PARAGRAPH.RIGHT,
              color=GRAY, sa=1)

# Filet séparateur
p_hr = doc.add_paragraph()
p_hr.paragraph_format.space_before = Pt(8)
p_hr.paragraph_format.space_after  = Pt(0)
pPr = p_hr._p.get_or_add_pPr()
pBdr = OxmlElement("w:pBdr")
bot = OxmlElement("w:bottom")
bot.set(qn("w:val"),   "single")
bot.set(qn("w:sz"),    "6")
bot.set(qn("w:space"), "1")
bot.set(qn("w:color"), "000000")
pBdr.append(bot)
pPr.append(pBdr)


# ═══════════════════════════════════════════════════════════════════════════
#  2. PROFIL
# ═══════════════════════════════════════════════════════════════════════════

section_title(doc, "Profil")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after  = Pt(0)
set_run(p.add_run(
    "Diplômée d\u2019une Licence Sciences de l\u2019\u00c9ducation (Rennes\u00a02, 2025) "
    "avec une sp\u00e9cialisation en action \u00e9ducative et dispositifs d\u2019insertion, "
    "j\u2019ai d\u00e9velopp\u00e9 une exp\u00e9rience solide dans l\u2019accompagnement "
    "de publics diversifi\u00e9s, l\u2019entretien individuel et le suivi de parcours. "
    "Originaire de Lorient, je souhaite mettre ces comp\u00e9tences au service "
    "des jeunes de la Mission Locale du Pays de Lorient."
), size=10.5)


# ═══════════════════════════════════════════════════════════════════════════
#  3. EXPÉRIENCES PROFESSIONNELLES
# ═══════════════════════════════════════════════════════════════════════════

section_title(doc, "Expériences professionnelles")

exp_header(doc,
    "Téléopératrice Recouvrement & Auditrice (appels mystères)",
    "Télémaque \u2013 Groupe Pénélope", "Caudan (56)",
    "Sept. 2025 \u2013 présent")
bullet(doc, "Conduite d\u2019entretiens téléphoniques, écoute et reformulation des situations")
bullet(doc, "Évaluation de la qualité d\u2019accueil (appels mystères) : observation, notation, reporting")
bullet(doc, "Suivi administratif rigoureux des dossiers")

exp_header(doc,
    "Chargée de Relation Client",
    "Foncia", "Rennes",
    "Mai \u2013 sept. 2025")
bullet(doc, "Accueil et orientation du public, gestion des demandes en face-à-face et par téléphone")
bullet(doc, "Suivi individualisé des dossiers, rédaction de comptes rendus (Word, Excel)")

exp_header(doc,
    "Bénévole — Tutorat de français pour allophones",
    "Association SupÉducation", "Rennes",
    "Janv. \u2013 avr. 2024")
bullet(doc, "Accompagnement individualisé à distance en petits groupes (4\u20136 personnes)")
bullet(doc, "Adaptation pédagogique aux besoins et au niveau de chaque apprenant")
bullet(doc, "Suivi de la progression et ajustement des méthodes (différenciation, écoute active)")

exp_header(doc,
    "Animatrice Périscolaire",
    "Mairie de Rennes \u2013 École REP", "Rennes",
    "Janv. \u2013 mai 2023")
bullet(doc, "Accueil et encadrement d\u2019enfants en réseau d\u2019éducation prioritaire")
bullet(doc, "Gestion de publics diversifiés, adaptation aux besoins individuels")

exp_header(doc,
    "Équipière \u2192 Formatrice \u2192 Chef d\u2019équipe",
    "McDonald\u2019s", "Rennes",
    "Sept. 2021 \u2013 nov. 2022")
bullet(doc, "Formation et intégration des nouveaux collaborateurs")
bullet(doc, "Management d\u2019équipe, coordination opérationnelle, gestion sous pression")


# ═══════════════════════════════════════════════════════════════════════════
#  4. FORMATION
# ═══════════════════════════════════════════════════════════════════════════

section_title(doc, "Formation")

t = borderless_table(doc, 1, 2)
cl = t.cell(0, 0)
cr = t.cell(0, 1)
set_col_width(cl, 12.0)
set_col_width(cr, 4.5)

p = cl.add_paragraph()
p.paragraph_format.space_before = Pt(5)
p.paragraph_format.space_after  = Pt(1)
set_run(p.add_run("Licence Sciences de l\u2019Éducation"), size=10.5, bold=True)
cell_para(cl, "Université Rennes 2 \u2014 Option Action Éducation et Formation (2 semestres)",
          size=10, italic=True, color=GRAY, sa=1)
cell_para(cl, "Dispositifs d\u2019insertion, politiques de formation, enquêtes terrain "
              "(observations en classe, entretiens enseignants et élèves)",
          size=10, color=GRAY, sa=0)

pd = cr.add_paragraph()
pd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
pd.paragraph_format.space_before = Pt(5)
pd.paragraph_format.space_after  = Pt(1)
set_run(pd.add_run("2022 \u2013 2025"), size=10, color=GRAY)

t2 = borderless_table(doc, 1, 2)
cl2 = t2.cell(0, 0)
cr2 = t2.cell(0, 1)
set_col_width(cl2, 12.0)
set_col_width(cr2, 4.5)

p2 = cl2.add_paragraph()
p2.paragraph_format.space_before = Pt(5)
p2.paragraph_format.space_after  = Pt(1)
set_run(p2.add_run("Baccalauréat Général"), size=10.5, bold=True)
cell_para(cl2, "Lycée Dupuy de Lôme \u2014 Lorient  |  Spécialités : SES \u00b7 LLCER Anglais",
          size=10, italic=True, color=GRAY, sa=0)

pd2 = cr2.add_paragraph()
pd2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
pd2.paragraph_format.space_before = Pt(5)
pd2.paragraph_format.space_after  = Pt(1)
set_run(pd2.add_run("2018 \u2013 2021"), size=10, color=GRAY)


# ═══════════════════════════════════════════════════════════════════════════
#  5. COMPÉTENCES
# ═══════════════════════════════════════════════════════════════════════════

section_title(doc, "Compétences")

t3 = borderless_table(doc, 1, 3)
cols_data = [
    ("Langues",       ["Français \u2014 natif", "Anglais \u2014 C1"]),
    ("Informatique",  ["Pack Office (Word, Excel)", "Outils collaboratifs en ligne"]),
    ("Autres",        ["Permis B", "Dispositifs insertion (CEJ, CIVIS…)", "Connaissance public 16-25 ans"]),
]
for i, (titre, items) in enumerate(cols_data):
    cell = t3.cell(0, i)
    set_col_width(cell, 5.5)
    cell_para(cell, titre, bold=True, size=10.5, sb=4, sa=2)
    for item in items:
        cell_para(cell, f"\u2022 {item}", size=10, sa=1)


# ─── Sauvegarde ─────────────────────────────────────────────────────────────
doc.save(OUTPUT)
print(f"Fichier généré : {OUTPUT}")
