"""
Générateur de lettre de motivation — ISFEC d'Arradon
Format : noir et blanc, simple, une seule page
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT   = "Calibri"
BLACK  = RGBColor(0, 0, 0)
OUTPUT = "/home/user/Lettre-de-motivation/Lettre_motivation_ISFEC_Arradon_M2E.docx"


# ─── Helpers ────────────────────────────────────────────────────────────────

def set_run(run, size=11, bold=False, italic=False):
    run.font.name      = FONT
    run.font.size      = Pt(size)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.color.rgb = BLACK


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


def cell_para(cell, text="", bold=False, size=11,
              align=WD_ALIGN_PARAGRAPH.LEFT, sb=0, sa=2):
    p = cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    if text:
        r = p.add_run(text)
        set_run(r, size=size, bold=bold)
    return p


def body_para(doc, text, justify=True, sb=0, sa=8, indent=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    if indent:
        p.paragraph_format.first_line_indent = Cm(1)
    r = p.add_run(text)
    set_run(r)
    return p


# ─── Document ───────────────────────────────────────────────────────────────
doc = Document()

section = doc.sections[0]
section.top_margin    = Cm(2.2)
section.bottom_margin = Cm(2.0)
section.left_margin   = Cm(2.5)
section.right_margin  = Cm(2.5)

normal = doc.styles["Normal"]
normal.font.name = FONT
normal.font.size = Pt(11)
normal.paragraph_format.space_before = Pt(0)
normal.paragraph_format.space_after  = Pt(0)


# ═══════════════════════════════════════════════════════════════════════════
#  1. EN-TÊTE
#     Expéditrice (gauche) | Destinataire (droite)
#     Dernière ligne expéditrice alignée avec 1ère ligne destinataire
# ═══════════════════════════════════════════════════════════════════════════

SENDER = [
    "Louiza Hadid",
    "06-50-37-56-47",
    "louizahdd@gmail.com",
    "3 rue de Ventspils",
    "56100 Lorient",
]

RECIPIENT = [
    "ISFEC Bretagne",
    "Site d'Arradon",
    "3 allée des Fougères, BP 25",
    "56610 Arradon",
]

# Chaque ligne expéditrice ≈ 11pt texte + 2pt espace = 13pt
# Pour aligner la dernière des 5 lignes avec la 1ère destinataire :
# on pousse le destinataire de (5-1) × 13 = 52pt vers le bas
LINE_H   = 13
SPACER   = (len(SENDER) - 1) * LINE_H   # 52pt

t = borderless_table(doc, 1, 2)
c_left  = t.cell(0, 0)
c_right = t.cell(0, 1)
set_col_width(c_left,  8.0)
set_col_width(c_right, 8.0)

for text in SENDER:
    cell_para(c_left, text, size=11, sa=2)

for i, text in enumerate(RECIPIENT):
    cell_para(c_right, text, size=11,
              align=WD_ALIGN_PARAGRAPH.RIGHT,
              sb=SPACER if i == 0 else 0,
              sa=2)


# ═══════════════════════════════════════════════════════════════════════════
#  2. DATE  (droite)
# ═══════════════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p.paragraph_format.space_before = Pt(20)
p.paragraph_format.space_after  = Pt(0)
set_run(p.add_run("Lorient, le 1er mars 2026"))


# ═══════════════════════════════════════════════════════════════════════════
#  3. OBJET
# ═══════════════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(20)
p.paragraph_format.space_after  = Pt(14)
r1 = p.add_run("Objet\u00a0: ")
set_run(r1, bold=True)
r2 = p.add_run("Candidature au Master MEEF \u2013 Mention 1er degré")
set_run(r2, bold=True)


# ═══════════════════════════════════════════════════════════════════════════
#  4. SALUTATION
# ═══════════════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(12)
set_run(p.add_run("Madame, Monsieur,"))


# ═══════════════════════════════════════════════════════════════════════════
#  5. CORPS
# ═══════════════════════════════════════════════════════════════════════════
body_para(doc,
    "En fin de licence de Sciences de l'éducation à Rennes 2, je candidate au Master MEEF "
    "mention 1er degré pour préparer le CRPE et devenir enseignante du premier degré. "
    "J'ai choisi l'ISFEC d'Arradon pour ses promotions à taille humaine — un cadre qui me "
    "semble plus adapté à cette année de préparation — et parce que les valeurs portées "
    "par l'enseignement catholique sont proches des miennes.")

body_para(doc,
    "Ces trois années à Rennes 2 m'ont permis de construire des bases solides en "
    "pédagogie, en psychologie du développement et en didactique. Elles ont "
    "progressivement conforté ma conviction : c'est dans l'enseignement du premier degré "
    "que je veux m'investir. La préparation au CRPE intégrée à votre Master constitue "
    "une étape décisive pour y accéder dans les meilleures conditions.")

body_para(doc,
    "Sur le plan pratique, mes expériences au contact des enfants ont renforcé cette "
    "orientation. En tant qu'animatrice périscolaire dans le quartier prioritaire de "
    "Villejean à Rennes, j'ai appris à adapter mon approche à des publics diversifiés "
    "et à faire preuve de patience et de bienveillance. J'ai également exercé comme "
    "tutrice de français auprès d'apprenants étrangers, développant ainsi la "
    "reformulation, la différenciation pédagogique et l'écoute active — des compétences "
    "directement transposables dans l'exercice du métier d'enseignante.")

body_para(doc,
    "Ces expériences m'ont confirmé que l'enseignement va bien au-delà de la "
    "transmission de savoirs : il s'agit avant tout d'accompagner chaque élève dans "
    "son développement. C'est dans cette perspective que je souhaite m'investir "
    "pleinement dans votre formation, convaincue que l'ISFEC d'Arradon m'offrira les "
    "outils théoriques et pratiques pour exercer ce métier avec rigueur et engagement.",
    sa=0)


# ═══════════════════════════════════════════════════════════════════════════
#  6. FORMULE DE POLITESSE
# ═══════════════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(12)
p.paragraph_format.space_after  = Pt(8)
set_run(p.add_run(
    "Dans l'attente d'une réponse favorable, je me tiens à votre disposition "
    "pour tout entretien ou renseignement complémentaire."))

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(28)
set_run(p.add_run(
    "Veuillez recevoir, Madame, Monsieur, l'expression de mes salutations distinguées."))


# ═══════════════════════════════════════════════════════════════════════════
#  7. SIGNATURE
# ═══════════════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(0)
set_run(p.add_run("Louiza Hadid"), bold=True)


# ─── Sauvegarde ─────────────────────────────────────────────────────────────
doc.save(OUTPUT)
print(f"Fichier généré : {OUTPUT}")
