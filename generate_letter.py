from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Marges ──────────────────────────────────────────────────────────────────
section = doc.sections[0]
section.top_margin    = Cm(2)
section.bottom_margin = Cm(2)
section.left_margin   = Cm(2.5)
section.right_margin  = Cm(2.5)

# ── Helpers ──────────────────────────────────────────────────────────────────
def set_font(run, name="Times New Roman", size=11, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_para(doc, text="", align=WD_ALIGN_PARAGRAPH.LEFT,
             bold=False, size=11, space_before=0, space_after=6,
             italic=False, color=None):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    if text:
        run = p.add_run(text)
        set_font(run, size=size, bold=bold, color=color)
        run.italic = italic
    return p

def add_two_column_header(doc, left_lines, right_lines):
    """Crée un paragraphe à deux colonnes via une table sans bordure."""
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    # Supprimer les bordures
    for cell in table.rows[0].cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            border = OxmlElement(f"w:{side}")
            border.set(qn("w:val"), "none")
            tcBorders.append(border)
        tcPr.append(tcBorders)

    left_cell  = table.cell(0, 0)
    right_cell = table.cell(0, 1)

    # Vider les paragraphes par défaut
    for cell in [left_cell, right_cell]:
        for p in cell.paragraphs:
            p._element.getparent().remove(p._element)

    # Colonne gauche
    for i, (txt, bold, size) in enumerate(left_lines):
        p = left_cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        run = p.add_run(txt)
        set_font(run, size=size, bold=bold)

    # Colonne droite (alignement à droite)
    for txt, bold, size in right_lines:
        p = right_cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        run = p.add_run(txt)
        set_font(run, size=size, bold=bold)

    return table


# ════════════════════════════════════════════════════════════════════════════
#  EN-TÊTE : expéditrice (gauche)  |  destinataire (droite)
# ════════════════════════════════════════════════════════════════════════════
left_lines = [
    ("LOUIZA HADID",            True,  12),
    ("3 rue de Ventspils",      False, 11),
    ("56100 Lorient",           False, 11),
    ("(+33) 6 50 37 56 47",     False, 11),
    ("louizahdd@gmail.com",     False, 11),
]

right_lines = [
    ("Madame la Directrice / Monsieur le Directeur", False, 11),
    ("ISFEC Bretagne – Site d'Arradon",              False, 11),
    ("3 allée des Fougères, BP 25",                  False, 11),
    ("56610 Arradon",                                False, 11),
]

add_two_column_header(doc, left_lines, right_lines)

# ── Espace ───────────────────────────────────────────────────────────────────
add_para(doc, space_before=8, space_after=0)

# ── Date ─────────────────────────────────────────────────────────────────────
add_para(doc, "Lorient, le 1er mars 2026",
         align=WD_ALIGN_PARAGRAPH.RIGHT, space_before=4, space_after=12)

# ── Objet ────────────────────────────────────────────────────────────────────
p_obj = doc.add_paragraph()
p_obj.paragraph_format.space_before = Pt(0)
p_obj.paragraph_format.space_after  = Pt(14)
run_label = p_obj.add_run("Objet : ")
set_font(run_label, bold=True, size=11)
run_val = p_obj.add_run("Candidature au Master M2E – Mention 1\u1d49\u02b3 degré")
set_font(run_val, bold=True, size=11)

# ── Salutation ───────────────────────────────────────────────────────────────
add_para(doc, "Madame, Monsieur,", space_before=0, space_after=10)

# ════════════════════════════════════════════════════════════════════════════
#  CORPS DE LA LETTRE
# ════════════════════════════════════════════════════════════════════════════

paragraphes = [
    # § 1 — Modifié : introduction spécifique à l'ISFEC d'Arradon (premier choix)
    (
        "Titulaire d'une licence en Sciences de l'éducation obtenue à l'Université Rennes 2, "
        "je me permets de vous adresser ma candidature pour intégrer le Master M2E, mention "
        "1\u1d49\u02b3 degré, proposé par l'ISFEC d'Arradon. Cet établissement représente mon "
        "premier choix de formation, en raison de la qualité de l'accompagnement qu'il offre "
        "et de son partenariat avec l'Université Catholique de l'Ouest, qui me semble "
        "particulièrement adapté à mon projet professionnel : devenir professeure des écoles."
    ),
    # § 2 — Parcours académique (corps conservé)
    (
        "Mon parcours en Sciences de l'éducation m'a permis de construire des bases solides "
        "en pédagogie, en psychologie du développement et en didactique. Ces apprentissages "
        "ont progressivement orienté ma réflexion vers l'enseignement du premier degré, un "
        "milieu dans lequel je me projette avec conviction. La préparation au CRPE intégrée "
        "à votre Master M2E constitue pour moi une étape décisive pour accéder à ce métier "
        "dans les meilleures conditions."
    ),
    # § 3 — Expériences (corps conservé)
    (
        "Sur le plan pratique, mes expériences au contact des enfants et des apprenants ont "
        "renforcé cette orientation. En tant qu'animatrice périscolaire dans le quartier "
        "prioritaire de Villejean à Rennes, j'ai appris à m'adapter à des publics diversifiés "
        "et à faire preuve de patience et de bienveillance dans des situations parfois "
        "complexes. J'ai également exercé comme tutrice de français auprès d'apprenants "
        "étrangers, une expérience qui m'a amenée à travailler la reformulation, la "
        "différenciation pédagogique et l'écoute active — des compétences directement "
        "transposables dans l'exercice du métier d'enseignante."
    ),
    # § 4 — Motivation et conclusion (corps conservé)
    (
        "Ces expériences m'ont confirmé que l'enseignement va bien au-delà de la simple "
        "transmission de savoirs : il s'agit d'accompagner des élèves dans leur "
        "développement, de les encourager et d'adapter sa pratique à leurs besoins. C'est "
        "dans cette perspective que je souhaite m'investir pleinement dans votre formation, "
        "convaincue que l'ISFEC d'Arradon m'apportera les ressources théoriques et "
        "pratiques nécessaires pour exercer ce métier avec rigueur et engagement."
    ),
]

for texte in paragraphes:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(10)
    p.paragraph_format.first_line_indent = Cm(1)
    run = p.add_run(texte)
    set_font(run, size=11)

# ── Formule de politesse ─────────────────────────────────────────────────────
add_para(
    doc,
    "Dans l'attente d'une réponse favorable, je me tiens à votre disposition "
    "pour tout entretien ou renseignement complémentaire.",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    space_before=4, space_after=10
)
add_para(
    doc,
    "Veuillez recevoir, Madame, Monsieur, l'expression de mes salutations distinguées.",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    space_before=0, space_after=30
)

# ── Signature ────────────────────────────────────────────────────────────────
add_para(doc, "Louiza Hadid",
         align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=11,
         space_before=0, space_after=0)

# ════════════════════════════════════════════════════════════════════════════
#  SAUVEGARDE
# ════════════════════════════════════════════════════════════════════════════
output_path = "/home/user/Lettre-de-motivation/Lettre_motivation_ISFEC_Arradon_M2E.docx"
doc.save(output_path)
print(f"Fichier généré : {output_path}")
