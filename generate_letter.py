"""
Générateur de lettre de motivation — ISFEC d'Arradon
----------------------------------------------------
Pour ajouter le logo ISFEC :
  1. Téléchargez le logo depuis https://www.isfec-bretagne.org/
     (clic droit sur le logo → Enregistrer l'image)
  2. Placez-le ici :  /home/user/Lettre-de-motivation/isfec_logo.png
  3. Relancez le script
"""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT      = "Calibri"
BLUE      = RGBColor(0x1B, 0x54, 0x8C)   # bleu ISFEC (sobre)
LOGO_PATH = "/home/user/Lettre-de-motivation/isfec_logo.png"
OUTPUT    = "/home/user/Lettre-de-motivation/Lettre_motivation_ISFEC_Arradon_M2E.docx"


# ─────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────

def set_run(run, size=11, bold=False, italic=False, color=None):
    run.font.name   = FONT
    run.font.size   = Pt(size)
    run.font.bold   = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color


def new_para(container, text="", align=WD_ALIGN_PARAGRAPH.LEFT,
             bold=False, italic=False, size=11, sb=0, sa=6,
             color=None, first_indent=None):
    p = container.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    if first_indent is not None:
        p.paragraph_format.first_line_indent = Cm(first_indent)
    if text:
        r = p.add_run(text)
        set_run(r, size=size, bold=bold, italic=italic, color=color)
    return p


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
    """Crée une table sans bordures, style Normal Table."""
    tbl = doc.add_table(rows=rows, cols=cols)
    tbl.style = "Normal Table"
    # Supprimer les bordures au niveau du tableau
    tblPr = tbl._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl._tbl.insert(0, tblPr)
    _clear_borders(tblPr, "w:tblBorders")
    # Supprimer les bordures au niveau de chaque cellule
    for row in tbl.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            _clear_borders(tcPr, "w:tcBorders")
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            # Vider les paragraphes vides par défaut
            for p in list(cell.paragraphs):
                p._element.getparent().remove(p._element)
    return tbl


def set_col_width(cell, cm):
    tcPr = cell._tc.get_or_add_tcPr()
    w = OxmlElement("w:tcW")
    w.set(qn("w:w"),    str(int(cm * 567)))   # 1 cm ≈ 567 twips
    w.set(qn("w:type"), "dxa")
    ex = tcPr.find(qn("w:tcW"))
    if ex is not None:
        tcPr.remove(ex)
    tcPr.append(w)


def cell_para(cell, text="", bold=False, italic=False, size=11,
              align=WD_ALIGN_PARAGRAPH.LEFT, sb=0, sa=2, color=None):
    p = cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    if text:
        r = p.add_run(text)
        set_run(r, size=size, bold=bold, italic=italic, color=color)
    return p


def horiz_rule(doc, color_hex="1B548C", thickness="6", sb=14, sa=14):
    """Ligne de séparation horizontale colorée."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    pPr  = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single")
    bot.set(qn("w:sz"),    thickness)
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), color_hex)
    pBdr.append(bot)
    ex = pPr.find(qn("w:pBdr"))
    if ex is not None:
        pPr.remove(ex)
    pPr.append(pBdr)


# ─────────────────────────────────────────────────────────────────────
#  Document
# ─────────────────────────────────────────────────────────────────────
doc = Document()

section = doc.sections[0]
section.top_margin    = Cm(1.8)
section.bottom_margin = Cm(1.8)
section.left_margin   = Cm(2.5)
section.right_margin  = Cm(2.5)

normal = doc.styles["Normal"]
normal.font.name = FONT
normal.font.size = Pt(11)
normal.paragraph_format.space_before = Pt(0)
normal.paragraph_format.space_after  = Pt(0)


# ═════════════════════════════════════════════════════════════════════
#  1. EN-TÊTE  — Expéditrice (gauche) | Logo (droite)
# ═════════════════════════════════════════════════════════════════════
# Page utile = 16 cm  →  sender 10 cm | logo 6 cm
t_header = borderless_table(doc, 1, 2)
c_name   = t_header.cell(0, 0)
c_logo   = t_header.cell(0, 1)
set_col_width(c_name, 10.0)
set_col_width(c_logo,  6.0)

# — Expéditrice
cell_para(c_name, "Louiza HADID",         bold=True, size=15, sa=5, color=BLUE)
cell_para(c_name, "3 rue de Ventspils",   size=10, sa=1)
cell_para(c_name, "56100 Lorient",        size=10, sa=1)
cell_para(c_name, "(+33) 6 50 37 56 47", size=10, sa=1)
cell_para(c_name, "louizahdd@gmail.com",  size=10, sa=0)

# — Logo ISFEC
if os.path.exists(LOGO_PATH):
    p = c_logo.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    p.add_run().add_picture(LOGO_PATH, width=Cm(4.5))
else:
    p = c_logo.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(0)
    r = p.add_run("[ Logo ISFEC Bretagne ]")
    set_run(r, size=9, italic=True, color=RGBColor(0xBB, 0xBB, 0xBB))


# ═════════════════════════════════════════════════════════════════════
#  2. LIGNE DE SÉPARATION
# ═════════════════════════════════════════════════════════════════════
horiz_rule(doc, color_hex="1B548C", thickness="6", sb=16, sa=16)


# ═════════════════════════════════════════════════════════════════════
#  3. DESTINATAIRE (gauche)  |  DATE (droite)
# ═════════════════════════════════════════════════════════════════════
t_addr = borderless_table(doc, 1, 2)
c_dest = t_addr.cell(0, 0)
c_date = t_addr.cell(0, 1)
set_col_width(c_dest, 9.5)
set_col_width(c_date, 6.5)

cell_para(c_dest, "ISFEC Bretagne",                 bold=True, size=11, sa=2)
cell_para(c_dest, "Site d'Arradon",                 size=11, sa=2)
cell_para(c_dest, "3 allée des Fougères, BP 25",    size=11, sa=2)
cell_para(c_dest, "56610 Arradon",                  size=11, sa=0)

cell_para(c_date, "Lorient, le 1er mars 2026",
          align=WD_ALIGN_PARAGRAPH.RIGHT, size=11, sa=0)


# ═════════════════════════════════════════════════════════════════════
#  4. OBJET
# ═════════════════════════════════════════════════════════════════════
new_para(doc, sb=20, sa=0)   # espace avant objet

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(18)
r1 = p.add_run("Objet : ")
set_run(r1, bold=True)
r2 = p.add_run("Candidature au Master MEEF – Mention 1er degré")
set_run(r2, bold=True)


# ═════════════════════════════════════════════════════════════════════
#  5. SALUTATION
# ═════════════════════════════════════════════════════════════════════
new_para(doc, "Madame, Monsieur,", sb=0, sa=14)


# ═════════════════════════════════════════════════════════════════════
#  6. CORPS
# ═════════════════════════════════════════════════════════════════════
paragraphes = [

    # § 1 — Choix de l'ISFEC d'Arradon (enrichi, détaillé)
    (
        "C'est avec un vif intérêt que je vous adresse ma candidature pour le Master MEEF, "
        "mention 1er degré, proposé par l'ISFEC Bretagne sur son site d'Arradon. Affilié à "
        "l'Université Catholique de l'Ouest pour la délivrance du diplôme, cet établissement "
        "s'inscrit pleinement dans le réseau national des ISFEC et se distingue par un "
        "accompagnement de proximité au sein de promotions à effectifs réduits. C'est ce "
        "cadre de formation, à la fois rigoureux et bienveillant, ancré dans les valeurs de "
        "l'enseignement catholique, qui en fait mon premier choix dans la perspective de "
        "préparer le CRPE et d'enseigner dans le premier degré."
    ),

    # § 2 — Parcours académique
    (
        "Mon parcours en Sciences de l'éducation à l'Université Rennes 2 m'a permis de "
        "construire des bases solides en pédagogie, en psychologie du développement et en "
        "didactique. Ces apprentissages ont progressivement conforté ma conviction : c'est "
        "dans l'enseignement du premier degré que je veux m'investir sur le long terme. La "
        "préparation au CRPE intégrée à votre Master constitue pour moi une étape décisive "
        "pour accéder à ce métier dans les meilleures conditions."
    ),

    # § 3 — Expériences (sans tiret cadratin)
    (
        "Sur le plan pratique, mes expériences au contact des enfants ont renforcé cette "
        "orientation. En tant qu'animatrice périscolaire dans le quartier prioritaire de "
        "Villejean à Rennes, j'ai appris à adapter mon approche à des publics diversifiés "
        "et à faire preuve de patience et de bienveillance dans des situations parfois "
        "complexes. J'ai également exercé comme tutrice de français auprès d'apprenants "
        "étrangers, ce qui m'a amenée à travailler la reformulation, la différenciation "
        "pédagogique et l'écoute active, des compétences directement transposables dans "
        "l'exercice du métier d'enseignante."
    ),

    # § 4 — Motivation et projection
    (
        "Ces expériences m'ont confirmé que l'enseignement va bien au-delà de la "
        "transmission de savoirs : il s'agit avant tout d'accompagner chaque élève dans "
        "son développement, de l'encourager et d'adapter sa pratique à ses besoins. C'est "
        "dans cette perspective que je souhaite m'investir pleinement dans votre formation, "
        "convaincue que l'ISFEC d'Arradon m'offrira les outils théoriques et pratiques "
        "pour exercer ce métier avec rigueur et engagement."
    ),
]

for texte in paragraphes:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(12)
    p.paragraph_format.first_line_indent = Cm(1)
    r = p.add_run(texte)
    set_run(r)


# ═════════════════════════════════════════════════════════════════════
#  7. FORMULE DE POLITESSE
# ═════════════════════════════════════════════════════════════════════
for texte, sa in [
    (
        "Dans l'attente d'une réponse favorable, je me tiens à votre disposition "
        "pour tout entretien ou renseignement complémentaire.",
        12,
    ),
    (
        "Veuillez recevoir, Madame, Monsieur, l'expression de mes salutations distinguées.",
        34,
    ),
]:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(sa)
    set_run(p.add_run(texte))


# ═════════════════════════════════════════════════════════════════════
#  8. SIGNATURE
# ═════════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(0)
set_run(p.add_run("Louiza Hadid"), bold=True)


# ─────────────────────────────────────────────────────────────────────
doc.save(OUTPUT)
print(f"Fichier généré : {OUTPUT}")

if not os.path.exists(LOGO_PATH):
    print()
    print("Logo ISFEC manquant. Pour l'ajouter :")
    print("  1. Allez sur https://www.isfec-bretagne.org/")
    print("  2. Clic droit sur le logo → 'Enregistrer l'image sous'")
    print(f"  3. Sauvegardez-le ici : {LOGO_PATH}")
    print("  4. Relancez :  python3 generate_letter.py")
