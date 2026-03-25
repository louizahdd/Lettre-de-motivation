"""
CV PDF — Design Nude/Beige — Louiza Hadid
Conseillère en Insertion Professionnelle — Mission Locale de Lorient
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor, white
from reportlab.pdfgen import canvas as CV
from reportlab.platypus import Paragraph, Spacer, HRFlowable, Frame, KeepTogether
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

OUTPUT = "/home/user/Lettre-de-motivation/CV_Louiza_Hadid_CIP.pdf"
W, H   = A4  # 595.27 x 841.89 pt

# ── Couleurs ──────────────────────────────────────────────────────────────────
HDR  = HexColor("#C4A882")
SIDE = HexColor("#EDE5D8")
BODY = HexColor("#FDFAF6")
BRN  = HexColor("#6E5232")
CRM  = HexColor("#FFF4E4")
DRK  = HexColor("#2D2A26")
GRY  = HexColor("#7A7265")
ACC  = HexColor("#C4A882")

# ── Fonts — Liberation Sans (équivalent Calibri) ──────────────────────────────
_base = "/usr/share/fonts/truetype/liberation/LiberationSans"
try:
    pdfmetrics.registerFont(TTFont("LS",    _base + "-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("LS-B",  _base + "-Bold.ttf"))
    pdfmetrics.registerFont(TTFont("LS-I",  _base + "-Italic.ttf"))
    pdfmetrics.registerFont(TTFont("LS-BI", _base + "-BoldItalic.ttf"))
    from reportlab.pdfbase.pdfmetrics import registerFontFamily
    registerFontFamily("LS", normal="LS", bold="LS-B",
                       italic="LS-I", boldItalic="LS-BI")
    F, FB, FI, FBI = "LS", "LS-B", "LS-I", "LS-BI"
except:
    F, FB, FI, FBI = ("Helvetica", "Helvetica-Bold",
                      "Helvetica-Oblique", "Helvetica-BoldOblique")

# ── Layout ────────────────────────────────────────────────────────────────────
HDR_H  = 3.3 * cm
SIDE_W = 6.6 * cm

# Frames
SL, SR         = 1.4*cm, 0.3*cm
MAIN_L, MAIN_R = 0.45*cm, 1.4*cm
BOT            = 1.2*cm

sf = dict(x1=SL, y1=BOT, width=SIDE_W-SL-SR,   height=H-HDR_H-BOT)
mf = dict(x1=SIDE_W+MAIN_L, y1=BOT,
          width=W-SIDE_W-MAIN_L-MAIN_R, height=H-HDR_H-BOT)

# ── Style helper ──────────────────────────────────────────────────────────────
def sty(name, font=F, size=9.5, color=DRK, sb=0, sa=2,
        align=TA_LEFT, li=0, leading=None):
    return ParagraphStyle(name, fontName=font, fontSize=size, textColor=color,
                          leading=leading or size*1.35,
                          spaceBefore=sb, spaceAfter=sa,
                          alignment=align, leftIndent=li)

# ── Flowables ─────────────────────────────────────────────────────────────────
def sec_title(title, side=False):
    sz = 8.5 if side else 9.5
    return [
        Spacer(1, 0.2*cm),
        Paragraph(title.upper(), sty("st", font=FB, size=sz, color=BRN, sb=0, sa=1)),
        HRFlowable(width="100%", thickness=0.5, color=ACC, spaceAfter=3),
    ]

def item(text, font=F, size=9, color=None, sb=0, sa=2, li=0):
    return Paragraph(text, sty("i", font=font, size=size,
                               color=color or DRK, sb=sb, sa=sa, li=li))

def bullet(text, size=9.5):
    txt = f'<font name="{F}">\u2013</font>\u00a0{text}'
    return Paragraph(txt, sty("b", font=F, size=size, sb=0, sa=1.5, li=12))

def exp_block(poste, structure, lieu, dates, bullets):
    elems = [
        Paragraph(
            f'<b>{poste}</b>\u00a0\u00a0'
            f'<font name="{FI}" size="9" color="{GRY.hexval()}">'
            f'\u2014\u00a0{dates}</font>',
            sty("ep", font=FB, size=9.5, sb=6, sa=0)),
        Paragraph(
            f'<i>{structure} \u2014 {lieu}</i>',
            sty("es", font=FI, size=9, color=GRY, sb=0, sa=1)),
    ] + [bullet(b) for b in bullets]
    return elems


# ══════════════════════════════════════════════════════════════════════════════
#  Canvas + backgrounds
# ══════════════════════════════════════════════════════════════════════════════
c = CV.Canvas(OUTPUT, pagesize=A4)

# Header
c.setFillColor(HDR)
c.rect(0, H - HDR_H, W, HDR_H, fill=1, stroke=0)

# Sidebar
c.setFillColor(SIDE)
c.rect(0, 0, SIDE_W, H - HDR_H, fill=1, stroke=0)

# Main
c.setFillColor(BODY)
c.rect(SIDE_W, 0, W - SIDE_W, H - HDR_H, fill=1, stroke=0)

# ── Header contenu ────────────────────────────────────────────────────────────
pad_l = SL
name_y  = H - 0.85*cm
title_y = H - 1.75*cm
info_y  = H - 2.55*cm

c.setFillColor(white)
c.setFont(FB, 25)
c.drawString(pad_l, name_y, "Louiza Hadid")

c.setFillColor(CRM)
c.setFont(FI, 12)
c.drawString(pad_l, title_y, "Conseill\u00e8re en insertion professionnelle")

c.setFillColor(HexColor("#F0E4D2"))
c.setFont(F, 8.5)
contact = ("06 50 37 56 47  \u2022  louizahdd@gmail.com  \u2022  "
           "3 rue de Ventspils, 56100 Lorient  \u2022  Permis B")
c.drawString(pad_l, info_y, contact)

# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR story
# ══════════════════════════════════════════════════════════════════════════════
side_story = []

# Formation
side_story += sec_title("Formation", side=True)
side_story += [
    item("<b>Licence Sciences de l\u2019\u00c9ducation</b>", size=9, sb=3, sa=0),
    item("Universit\u00e9 Rennes 2", font=FI, size=8.5, color=GRY, sa=0),
    item("2022 \u2013 2025", font=FI, size=8.5, color=GRY, sa=4),
    item("Option\u00a0: Action \u00c9ducation et Formation", size=8.5, sa=1, li=5),
    item("(2 semestres)", size=8.5, sa=3, li=5),
]
for ligne in [
    "\u00b7 Dispositifs d\u2019insertion (CEJ, CIVIS,",
    "  apprentissage, formation professionnelle)",
    "\u00b7 Accompagnement des publics",
    "\u00b7 Psychologie du d\u00e9veloppement",
    "\u00b7 Enqu\u00eates terrain\u00a0: observations classe,",
    "  entretiens enseignants & \u00e9l\u00e8ves",
]:
    side_story.append(item(ligne, size=8.5, sa=1, li=5))

side_story += [
    Spacer(1, 0.2*cm),
    item("<b>Baccalaur\u00e9at G\u00e9n\u00e9ral</b>", size=9, sb=0, sa=0),
    item("Lyc\u00e9e Dupuy de L\u00f4me \u2014 Lorient", font=FI, size=8.5, color=GRY, sa=0),
    item("2018 \u2013 2021", font=FI, size=8.5, color=GRY, sa=0),
    item("SES \u00b7 LLCER Anglais", font=FI, size=8.5, color=GRY, sa=2),
]

# Compétences
side_story += sec_title("Comp\u00e9tences", side=True)
for sk in [
    "\u00b7 \u00c9coute active et bienveillance",
    "\u00b7 Accompagnement individualis\u00e9",
    "\u00b7 Techniques d\u2019entretien",
    "\u00b7 Suivi administratif de dossiers",
    "\u00b7 Animation de groupe",
    "\u00b7 M\u00e9diation et relationnel",
    "\u00b7 Travail en \u00e9quipe",
    "\u00b7 Rigueur et organisation",
]:
    side_story.append(item(sk, size=9, sa=2))

# Langues & Outils
side_story += sec_title("Langues & Outils", side=True)
for it in ["\u00b7 Fran\u00e7ais \u2014 natif", "\u00b7 Anglais \u2014 C1"]:
    side_story.append(item(it, size=9, sa=2))
side_story.append(item("<b>Informatique</b>", size=9, sb=5, sa=1))
for it in ["\u00b7 Pack Office (Word, Excel)", "\u00b7 Outils collaboratifs en ligne"]:
    side_story.append(item(it, size=9, sa=2))

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN story
# ══════════════════════════════════════════════════════════════════════════════
main_story = []

# Profil
main_story += sec_title("Profil")
main_story.append(Paragraph(
    "Dipl\u00f4m\u00e9e d\u2019une Licence Sciences de l\u2019\u00c9ducation, je suis "
    "rigoureuse, organis\u00e9e et dot\u00e9e d\u2019un sens aigu de l\u2019\u00e9coute. "
    "Mes exp\u00e9riences dans l\u2019animation en milieu prioritaire, la relation "
    "client et le tutorat de publics allophones m\u2019ont appris \u00e0 adapter mon "
    "approche, \u00e0 g\u00e9rer des situations complexes et \u00e0 assurer un suivi "
    "individualis\u00e9. Familiarise\u0301e avec les dispositifs d\u2019insertion "
    "gr\u00e2ce \u00e0 ma formation, je suis motiv\u00e9e \u00e0 mettre mes "
    "comp\u00e9tences au service des jeunes de la Mission Locale du Pays de Lorient.",
    sty("pr", font=F, size=9.5, sb=2, sa=0, align=TA_JUSTIFY)))

# Expériences
main_story += sec_title("Exp\u00e9riences professionnelles")

main_story += exp_block(
    "T\u00e9l\u00e9op\u00e9ratrice Recouvrement & Auditrice",
    "T\u00e9l\u00e9maque \u2013 Groupe P\u00e9n\u00e9lope", "Caudan (56)",
    "Sept. 2025 \u2013 pr\u00e9sent",
    [
        "Conduite d\u2019appels\u00a0: \u00e9coute active, reformulation, "
        "recherche de solutions",
        "Appels myst\u00e8res\u00a0: \u00e9valuation qualit\u00e9 d\u2019accueil, "
        "notation, reporting",
        "Suivi administratif rigoureux des dossiers (saisie, mise \u00e0 jour, classement)",
    ]
)

main_story += exp_block(
    "Charg\u00e9e de Relation Client",
    "Foncia", "Rennes",
    "Mai \u2013 sept. 2025",
    [
        "Accueil physique et t\u00e9l\u00e9phonique, orientation et traitement "
        "des demandes",
        "Suivi individualis\u00e9 des dossiers, r\u00e9daction de comptes rendus, "
        "Pack Office quotidien",
        "Gestion de situations conflictuelles avec calme et professionnalisme",
    ]
)

main_story += exp_block(
    "B\u00e9n\u00e9vole \u2014 Tutorat de fran\u00e7ais pour allophones",
    "Association Sup\u00c9ducation", "Rennes",
    "Janv. \u2013 avr. 2024",
    [
        "Accompagnement individualis\u00e9 en petits groupes (4\u20136\u00a0pers.) "
        "\u00e0 distance",
        "Adaptation p\u00e9dagogique aux besoins individuels, suivi de la progression",
        "\u00c9coute active, patience, m\u00e9diation avec des publics vuln\u00e9rables",
    ]
)

main_story += exp_block(
    "Animatrice P\u00e9riscolaire",
    "Mairie de Rennes \u2013 \u00c9cole REP", "Rennes",
    "Janv. \u2013 mai 2023",
    [
        "Accueil et encadrement d\u2019enfants en r\u00e9seau d\u2019\u00e9ducation "
        "prioritaire",
        "Adaptation aux publics diversifi\u00e9s, maintien d\u2019un cadre "
        "bienveillant et structurant",
        "Collaboration avec l\u2019\u00e9quipe p\u00e9dagogique et lien avec "
        "les familles",
    ]
)

main_story += exp_block(
    "\u00c9quipi\u00e8re \u2192 Formatrice \u2192 Chef d\u2019\u00e9quipe",
    "McDonald\u2019s", "Rennes",
    "Sept. 2021 \u2013 nov. 2022",
    [
        "Formation et int\u00e9gration des nouveaux collaborateurs",
        "Management d\u2019\u00e9quipe, coordination et gestion sous pression",
        "Progression rapide en 14 mois, sens des responsabilit\u00e9s",
    ]
)

# ══════════════════════════════════════════════════════════════════════════════
#  Flow stories into frames
# ══════════════════════════════════════════════════════════════════════════════
Frame(**sf, showBoundary=0).addFromList(side_story, c)
Frame(**mf, showBoundary=0).addFromList(main_story, c)

c.save()
print(f"PDF généré : {OUTPUT}")
