"""
CV PDF v3 — Design Moderne Épuré — Louiza Hadid
Blanc · Accent bleu ardoise · Montserrat · ATS-optimisé
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor, white
from reportlab.pdfgen import canvas as CV
from reportlab.platypus import (Paragraph, Spacer, HRFlowable,
                                 Frame, Table, TableStyle)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY, TA_RIGHT

OUTPUT = "/home/user/Lettre-de-motivation/CV_Louiza_Hadid_CIP.pdf"
W, H   = A4

# ── Fonts (Montserrat) ────────────────────────────────────────────────────────
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily

_ub = "/usr/share/fonts/truetype/ubuntu/Ubuntu"
pdfmetrics.registerFont(TTFont("UB",   _ub + "-R.ttf"))
pdfmetrics.registerFont(TTFont("UB-M", _ub + "-M.ttf"))
pdfmetrics.registerFont(TTFont("UB-B", _ub + "-B.ttf"))
pdfmetrics.registerFont(TTFont("UB-I", _ub + "-RI.ttf"))
pdfmetrics.registerFont(TTFont("UB-MI",_ub + "-MI.ttf"))
registerFontFamily("UB", normal="UB", bold="UB-B", italic="UB-I", boldItalic="UB-B")
F, FM, FSB, FB, FI = "UB", "UB-M", "UB-M", "UB-B", "UB-I"

# ── Palette ───────────────────────────────────────────────────────────────────
NAVY = HexColor("#1D3461")
BLUE = HexColor("#2558A8")
DARK = HexColor("#1A1A2E")
GRY  = HexColor("#6B7280")
LGRY = HexColor("#D1D5DB")

# ── Layout ────────────────────────────────────────────────────────────────────
MX, MY = 1.8*cm, 1.5*cm
CW     = W - 2*MX
HDR_H  = 2.8*cm

# ── Style factory ─────────────────────────────────────────────────────────────
def S(name, font=F, size=9.5, color=DARK, sb=0, sa=2,
      align=TA_LEFT, li=0, leading=None):
    return ParagraphStyle(name,
        fontName=font, fontSize=size, textColor=color,
        leading=leading or size * 1.42,
        spaceBefore=sb, spaceAfter=sa,
        alignment=align, leftIndent=li)

# ── Helpers ───────────────────────────────────────────────────────────────────
def sec(title):
    return [
        Spacer(1, 0.28*cm),
        Paragraph(title.upper(),
            S("st", font=FSB, size=8.5, color=BLUE, sb=0, sa=1)),
        HRFlowable(width="100%", thickness=0.8, color=BLUE, spaceAfter=5),
    ]

def bullet_p(text):
    return Paragraph(
        f'<font name="{FSB}" color="{BLUE.hexval()}">▸</font>  {text}',
        S("b", font=FM, size=9.5, sb=0, sa=2, li=10))

def exp(poste, structure, lieu, dates, bullets):
    # Ligne poste | dates via table 2 colonnes
    hdr = Table([[
        Paragraph(f'<font name="{FB}" size="10" color="{NAVY.hexval()}">{poste}</font>',
                  S("et", font=FB, size=10, color=NAVY, sb=7, sa=0)),
        Paragraph(f'<font name="{FM}" size="8.5" color="{GRY.hexval()}">{dates}</font>',
                  S("ed", font=FM, size=8.5, color=GRY, sb=7, sa=0, align=TA_RIGHT)),
    ]], colWidths=[CW * 0.65, CW * 0.35])
    hdr.setStyle(TableStyle([
        ("VALIGN",       (0,0), (-1,-1), "BOTTOM"),
        ("LEFTPADDING",  (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING",   (0,0), (-1,-1), 0),
        ("BOTTOMPADDING",(0,0), (-1,-1), 0),
    ]))
    elems = [hdr,
        Paragraph(f'<font name="{FI}" size="9" color="{GRY.hexval()}">'
                  f'{structure} — {lieu}</font>',
                  S("es", font=FI, size=9, color=GRY, sb=1, sa=3))]
    elems += [bullet_p(b) for b in bullets]
    return elems

def form(diplome, ecole, dates, details=None):
    hdr = Table([[
        Paragraph(f'<font name="{FSB}" size="10" color="{NAVY.hexval()}">{diplome}</font>',
                  S("ft", font=FSB, size=10, color=NAVY, sb=6, sa=0)),
        Paragraph(f'<font name="{FM}" size="8.5" color="{GRY.hexval()}">{dates}</font>',
                  S("fd", font=FM, size=8.5, color=GRY, sb=6, sa=0, align=TA_RIGHT)),
    ]], colWidths=[CW * 0.65, CW * 0.35])
    hdr.setStyle(TableStyle([
        ("VALIGN",       (0,0), (-1,-1), "BOTTOM"),
        ("LEFTPADDING",  (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING",   (0,0), (-1,-1), 0),
        ("BOTTOMPADDING",(0,0), (-1,-1), 0),
    ]))
    elems = [hdr,
        Paragraph(ecole, S("fe", font=FI, size=9, color=GRY, sb=1, sa=2))]
    if details:
        for d in details:
            elems.append(bullet_p(d))
    return elems


# ══════════════════════════════════════════════════════════════════════════════
#  Canvas + header
# ══════════════════════════════════════════════════════════════════════════════
c = CV.Canvas(OUTPUT, pagesize=A4)

# Bande navy en haut
c.setFillColor(NAVY)
c.rect(0, H - HDR_H, W, HDR_H, fill=1, stroke=0)

# Trait bleu sous header
c.setStrokeColor(BLUE)
c.setLineWidth(2.5)
c.line(0, H - HDR_H - 1.5, W, H - HDR_H - 1.5)

# Nom
c.setFillColor(white)
c.setFont(FB, 26)
c.drawString(MX, H - 1.05*cm, "LOUIZA HADID")

# Titre
c.setFillColor(HexColor("#93B4DD"))
c.setFont(FM, 11)
c.drawString(MX, H - 1.85*cm, "Conseillère en insertion professionnelle")

# Contact
c.setFillColor(HexColor("#BFD0E8"))
c.setFont(F, 8.5)
c.drawString(MX, H - 2.52*cm,
    "06 50 37 56 47   ·   louizahdd@gmail.com   ·   Lorient (56)   ·   Permis B")


# ══════════════════════════════════════════════════════════════════════════════
#  Story
# ══════════════════════════════════════════════════════════════════════════════
story = []

# PROFIL ──────────────────────────────────────────────────────────────────────
story += sec("Profil")
story.append(Paragraph(
    "Diplômée en Sciences de l'Éducation (Rennes 2, 2025), spécialisée en "
    "<b>dispositifs d'insertion et accompagnement des publics</b>. "
    "Solide expérience en relation client, animation en milieu prioritaire "
    "et tutorat individualisé. Capable de mener un entretien, "
    "poser un diagnostic et assurer un suivi rigoureux. "
    "Motivée, organisée et prête à m'engager pour les jeunes 16–25 ans "
    "de la Mission Locale du Pays de Lorient.",
    S("pr", font=FM, size=9.5, sb=2, sa=0, align=TA_JUSTIFY)))

# EXPÉRIENCES ─────────────────────────────────────────────────────────────────
story += sec("Expériences professionnelles")

story += exp(
    "Téléopératrice Recouvrement &amp; Auditrice Qualité",
    "Télémaque – Groupe Pénélope", "Caudan (56)",
    "Sept. 2025 – présent",
    [
        "Conduite d'entretiens téléphoniques : analyse de situation, "
        "reformulation, proposition de solutions de paiement adaptées",
        "Évaluation qualité par <b>appels mystères</b> : grille "
        "multicritères, notation et restitution d'analyses",
        "Gestion et mise à jour de dossiers dans un environnement "
        "à fort volume",
    ]
)

story += exp(
    "Chargée de Relation Client",
    "Foncia", "Rennes",
    "Mai – sept. 2025",
    [
        "Accueil et orientation d'un flux quotidien de clients "
        "(téléphone + physique)",
        "Traitement de réclamations complexes : écoute, "
        "reformulation, résolution dans les délais",
        "Suivi et archivage de dossiers locatifs sous Pack Office",
    ]
)

story += exp(
    "Bénévole – Tutorat de français langue étrangère",
    "Association SupÉducation", "Rennes",
    "Janv. – avr. 2024",
    [
        "Accompagnement de <b>4 à 6 apprenants allophones</b> par "
        "session, en format individuel et semi-collectif, à distance",
        "Évaluation des besoins, adaptation pédagogique aux "
        "niveaux A1–B1, suivi individualisé de la progression",
        "Pratique de l'écoute active, de la bienveillance et "
        "de la médiation avec des publics vulnérables",
    ]
)

story += exp(
    "Animatrice Périscolaire",
    "Mairie de Rennes – École REP", "Rennes",
    "Janv. – mai 2023",
    [
        "Animation de groupes de 10 à 20 élèves en zone "
        "d'éducation prioritaire",
        "Médiation, gestion de conflits, maintien d'un cadre "
        "bienveillant et structurant",
        "Collaboration étroite avec l'équipe pédagogique, "
        "lien avec les familles",
    ]
)

story += exp(
    "Équipière  ▶  Formatrice  ▶  Chef d'équipe",
    "McDonald's", "Rennes",
    "Sept. 2021 – nov. 2022",
    [
        "Formation et intégration des nouveaux collaborateurs "
        "à leurs postes de travail",
        "Management d'une équipe de 6 à 8 personnes en service",
        "<b>Évolution interne en 14 mois</b> : équipière "
        "→ formatrice → chef d'équipe",
    ]
)

# FORMATION ───────────────────────────────────────────────────────────────────
story += sec("Formation")

story += form(
    "Licence Sciences de l'Éducation",
    "Université Rennes 2",
    "2022 – 2025",
    [
        "Option Action Éducation et Formation (2 sem.) — "
        "dispositifs d'insertion, CEJ, CIVIS, politiques de formation",
        "Enquêtes terrain : observations en classe, entretiens "
        "enseignants et élèves",
    ]
)

story += form(
    "Baccalauréat Général — SES · LLCER Anglais",
    "Lycée Dupuy de Lôme, Lorient",
    "2021",
)

# COMPÉTENCES + LANGUES + INFOS ───────────────────────────────────────────────
story += sec("Compétences & Informations")
story.append(Spacer(1, 0.1*cm))

cw3 = CW / 3

def mini_col(items, titre=None):
    lines = []
    if titre:
        lines.append(Paragraph(titre,
            S("ct", font=FSB, size=8.5, color=NAVY, sb=0, sa=3)))
    for it in items:
        lines.append(Paragraph(
            f'<font name="{FSB}" color="{BLUE.hexval()}">▸</font>  {it}',
            S("ci", font=FM, size=9, sb=0, sa=2)))
    return lines

col1 = mini_col([
    "Entretien individuel",
    "Diagnostic de situation",
    "Accompagnement de parcours",
    "Médiation & gestion de conflits",
    "Animation de groupe",
    "Suivi administratif",
    "Accueil & relation client",
], "Compétences clés")

col2 = (mini_col(["Français — natif", "Anglais — C1 (LLCER)"], "Langues")
      + [Spacer(1, 0.2*cm)]
      + mini_col([
            "Pack Office (Word, Excel)",
            "Outils collaboratifs",
            "Dispositifs : CEJ, CIVIS…",
        ], "Outils & Dispositifs"))

col3 = mini_col([
    "Permis B + véhiculée",
    "Disponible immédiatement",
    "Lorient (56)",
], "Informations")

tbl = Table([[col1, col2, col3]], colWidths=[cw3, cw3, cw3])
tbl.setStyle(TableStyle([
    ("VALIGN",       (0,0), (-1,-1), "TOP"),
    ("LEFTPADDING",  (0,0), (-1,-1), 0),
    ("RIGHTPADDING", (0,0), (-1,-1), 10),
    ("TOPPADDING",   (0,0), (-1,-1), 0),
    ("BOTTOMPADDING",(0,0), (-1,-1), 0),
]))
story.append(tbl)

# ── Frame ─────────────────────────────────────────────────────────────────────
Frame(MX, MY, CW, H - HDR_H - MY - 0.4*cm,
      leftPadding=0, rightPadding=0,
      topPadding=0, bottomPadding=0,
      showBoundary=0).addFromList(story, c)

c.save()
print(f"PDF généré : {OUTPUT}")
