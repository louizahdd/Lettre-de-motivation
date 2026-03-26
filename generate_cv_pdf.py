"""
CV PDF v4 — Esthétique Épurée Beige/Crème/Pastel
Louiza Hadid — Conseillère en Insertion Professionnelle
Police : Lato | Palette : crème, beige rosé, sauge pâle
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor, white, Color
from reportlab.pdfgen import canvas as CV
from reportlab.platypus import (Paragraph, Spacer, HRFlowable,
                                 Frame, Table, TableStyle)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY, TA_RIGHT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily

OUTPUT = "/home/user/Lettre-de-motivation/CV_Louiza_Hadid_CIP.pdf"
W, H   = A4   # 595.27 × 841.89 pt

# ── Lato ──────────────────────────────────────────────────────────────────────
_L = "/usr/share/fonts/truetype/lato/Lato"
pdfmetrics.registerFont(TTFont("LT",    _L + "-Regular.ttf"))
pdfmetrics.registerFont(TTFont("LT-L",  _L + "-Light.ttf"))
pdfmetrics.registerFont(TTFont("LT-M",  _L + "-Medium.ttf"))
pdfmetrics.registerFont(TTFont("LT-SB", _L + "-Semibold.ttf"))
pdfmetrics.registerFont(TTFont("LT-B",  _L + "-Bold.ttf"))
pdfmetrics.registerFont(TTFont("LT-I",  _L + "-Italic.ttf"))
pdfmetrics.registerFont(TTFont("LT-LI", _L + "-LightItalic.ttf"))
registerFontFamily("LT", normal="LT", bold="LT-B",
                   italic="LT-I", boldItalic="LT-B")

F   = "LT"        # Regular
FL  = "LT-L"      # Light
FM  = "LT-M"      # Medium
FSB = "LT-SB"     # SemiBold
FB  = "LT-B"      # Bold
FI  = "LT-I"      # Italic
FLI = "LT-LI"     # Light Italic

# ── Palette ───────────────────────────────────────────────────────────────────
# Fonds
PAGE_BG  = HexColor("#FAF8F4")   # crème très doux — fond page
HDR_BG   = HexColor("#F2EDE5")   # beige clair — fond header

# Accents pastels (discrets)
ROSE     = HexColor("#C9968E")   # rose poudré — ligne déco, bullets
ROSE_LT  = HexColor("#EDD9D6")   # rose pâle — très léger
SAGE     = HexColor("#9BB09E")   # sauge pâle — titres sections
SEP      = HexColor("#D8D2CC")   # gris beige — séparateurs fins

# Texte
INK      = HexColor("#1C1C1C")   # quasi-noir — texte principal
CHARCOAL = HexColor("#3A3A3A")   # titres postes
STONE    = HexColor("#787878")   # dates, entreprises, secondaire

# ── Dimensions ────────────────────────────────────────────────────────────────
MX, MY = 1.85*cm, 1.0*cm
CW     = W - 2*MX             # ~493 pt
HDR_H  = 2.85*cm

# ── Style factory ─────────────────────────────────────────────────────────────
def S(name, font=F, size=9.5, color=INK, sb=0, sa=2,
      align=TA_LEFT, li=0, leading=None):
    return ParagraphStyle(name,
        fontName=font, fontSize=size, textColor=color,
        leading=leading or size * 1.5,
        spaceBefore=sb, spaceAfter=sa,
        alignment=align, leftIndent=li)

# ── Helpers ───────────────────────────────────────────────────────────────────
def sec(title):
    """Titre de section : majuscules, sauge, filet fin."""
    safe = title.upper().replace("&", "&amp;")
    return [
        Spacer(1, 0.18*cm),
        Paragraph(safe,
            S("st", font=FSB, size=8, color=SAGE, sb=0, sa=1,
              leading=8*1.4)),
        HRFlowable(width="100%", thickness=0.5, color=SEP, spaceAfter=4),
    ]

def bullet_p(text):
    """Bullet avec tiret rosé."""
    return Paragraph(
        f'<font name="{LT_ROSE}" color="{ROSE.hexval()}">–</font>'
        f'<font name="{F}">  {text}</font>',
        S("b", font=F, size=9.5, sb=0, sa=1.5, li=12))

# Raccourci couleur pour le tiret
LT_ROSE = FSB   # just uses FSB weight for the dash

def bullet_p(text):
    return Paragraph(
        f'<font color="{ROSE.hexval()}">–</font>  {text}',
        S("b", font=F, size=9.5, sb=0, sa=1.5, li=12))

def hdr_table(left_para, right_para, left_w=0.63):
    """Tableau 2 colonnes sans bordure pour aligner titre | dates."""
    t = Table([[left_para, right_para]],
              colWidths=[CW * left_w, CW * (1 - left_w)])
    t.setStyle(TableStyle([
        ("VALIGN",       (0,0), (-1,-1), "BOTTOM"),
        ("LEFTPADDING",  (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING",   (0,0), (-1,-1), 0),
        ("BOTTOMPADDING",(0,0), (-1,-1), 0),
    ]))
    return t

def exp(poste, structure, lieu, dates, bullets):
    elems = [
        hdr_table(
            Paragraph(
                f'<font name="{FSB}" size="10" color="{CHARCOAL.hexval()}">'
                f'{poste}</font>',
                S("et", font=FSB, size=10, color=CHARCOAL, sb=5, sa=0)),
            Paragraph(
                f'<font name="{FL}" size="8.5" color="{STONE.hexval()}">'
                f'{dates}</font>',
                S("ed", font=FL, size=8.5, color=STONE,
                  sb=5, sa=0, align=TA_RIGHT)),
        ),
        Paragraph(
            f'<font name="{FLI}" size="9" color="{STONE.hexval()}">'
            f'{structure}  ·  {lieu}</font>',
            S("es", font=FLI, size=9, color=STONE, sb=1, sa=2)),
    ]
    elems += [bullet_p(b) for b in bullets]
    return elems

def form(diplome, ecole, dates, details=None):
    elems = [
        hdr_table(
            Paragraph(
                f'<font name="{FSB}" size="10" color="{CHARCOAL.hexval()}">'
                f'{diplome}</font>',
                S("ft", font=FSB, size=10, color=CHARCOAL, sb=8, sa=0)),
            Paragraph(
                f'<font name="{FL}" size="8.5" color="{STONE.hexval()}">'
                f'{dates}</font>',
                S("fd", font=FL, size=8.5, color=STONE,
                  sb=8, sa=0, align=TA_RIGHT)),
        ),
        Paragraph(ecole,
            S("fe", font=FLI, size=9, color=STONE, sb=1, sa=2)),
    ]
    if details:
        for d in details:
            elems.append(bullet_p(d))
    return elems


# ══════════════════════════════════════════════════════════════════════════════
#  Canvas — fonds et header
# ══════════════════════════════════════════════════════════════════════════════
c = CV.Canvas(OUTPUT, pagesize=A4)

# Fond crème pleine page
c.setFillColor(PAGE_BG)
c.rect(0, 0, W, H, fill=1, stroke=0)

# Fond beige header
c.setFillColor(HDR_BG)
c.rect(0, H - HDR_H, W, HDR_H, fill=1, stroke=0)

# Ligne décorative rose sous le nom (courte, 5cm, gauche)
c.setFillColor(ROSE)
c.rect(MX, H - 1.45*cm, 4.8*cm, 1.5, fill=1, stroke=0)

# Ligne de séparation bas header (très fine, beige-gris)
c.setFillColor(SEP)
c.rect(0, H - HDR_H, W, 0.7, fill=1, stroke=0)

# ── Nom ───────────────────────────────────────────────────────────────────────
c.setFillColor(INK)
c.setFont(FB, 25)
c.drawString(MX, H - 1.05*cm, "LOUIZA HADID")

# ── Titre professionnel ───────────────────────────────────────────────────────
c.setFillColor(STONE)
c.setFont(FLI, 11)
c.drawString(MX, H - 1.78*cm,
             "Conseillère en insertion professionnelle")

# ── Contacts ─────────────────────────────────────────────────────────────────
c.setFillColor(STONE)
c.setFont(FL, 8.5)
c.drawString(MX, H - 2.48*cm,
    "06 50 37 56 47   ·   louizahdd@gmail.com   ·   "
    "Lorient (56)   ·   Permis B")


# ══════════════════════════════════════════════════════════════════════════════
#  Story
# ══════════════════════════════════════════════════════════════════════════════
story = []

# PROFIL ──────────────────────────────────────────────────────────────────────
story += sec("Profil")
story.append(Paragraph(
    "Diplômée en Sciences de l'Éducation (Rennes 2, 2025), "
    "spécialisée en <b>dispositifs d'insertion et accompagnement "
    "des publics</b>. Solide expérience en relation client, "
    "animation en milieu prioritaire et tutorat individualisé. "
    "À l'aise en entretien individuel, diagnostic de situation "
    "et suivi rigoureux. Motivée à s'investir pour les jeunes "
    "16–25 ans de la Mission Locale du Pays de Lorient.",
    S("pr", font=F, size=9.5, sb=2, sa=0, align=TA_JUSTIFY,
      leading=9.5*1.42)))

# EXPÉRIENCES ─────────────────────────────────────────────────────────────────
story += sec("Expériences professionnelles")

story += exp(
    "Téléopératrice Recouvrement &amp; Auditrice Qualité",
    "Télémaque – Groupe Pénélope", "Caudan (56)",
    "Sept. 2025 – présent",
    [
        "Conduite d'entretiens téléphoniques : analyse de situation, "
        "reformulation, proposition de solutions adaptées",
        "Évaluation qualité par appels mystères : grille multicritères, "
        "notation et restitution d'analyses",
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
        "Traitement de réclamations : écoute, reformulation, "
        "résolution dans les délais",
        "Suivi et archivage de dossiers locatifs sous Pack Office",
    ]
)

story += exp(
    "Bénévole – Tutorat de français langue étrangère",
    "Association SupÉducation", "Rennes",
    "Janv. – avr. 2024",
    [
        "Accompagnement de 4 à 6 apprenants allophones par session, "
        "individuel et semi-collectif, à distance",
        "Adaptation pédagogique aux niveaux A1–B1, "
        "suivi individualisé de la progression",
        "Écoute active, bienveillance et médiation "
        "avec des publics vulnérables",
    ]
)

story += exp(
    "Animatrice Périscolaire",
    "Mairie de Rennes – École REP", "Rennes",
    "Janv. – mai 2023",
    [
        "Animation de groupes de 10 à 20 élèves "
        "en zone d'éducation prioritaire",
        "Médiation, gestion de conflits, maintien d'un cadre "
        "bienveillant et structurant",
        "Collaboration avec l'équipe pédagogique, lien avec les familles",
    ]
)

story += exp(
    "Équipière  ›  Formatrice  ›  Chef d'équipe",
    "McDonald's", "Rennes",
    "Sept. 2021 – nov. 2022",
    [
        "Formation et intégration des nouveaux collaborateurs",
        "Management d'une équipe de 6 à 8 personnes en service",
        "Évolution interne en 14 mois : équipière → formatrice "
        "→ chef d'équipe",
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
        "dispositifs CEJ, CIVIS, politiques de formation et d'insertion",
        "Enquêtes terrain : observations en classe, entretiens "
        "enseignants et élèves",
    ]
)

story += form(
    "Baccalauréat Général — SES  ·  LLCER Anglais",
    "Lycée Dupuy de Lôme, Lorient",
    "2021",
)

# COMPÉTENCES & INFORMATIONS ──────────────────────────────────────────────────
story += sec("Compétences & Informations")
story.append(Spacer(1, 0.05*cm))

cw3 = CW / 3

def col(items, titre=None):
    lines = []
    if titre:
        lines.append(Paragraph(titre,
            S("ct", font=FSB, size=8.5, color=CHARCOAL, sb=0, sa=2)))
    for it in items:
        lines.append(Paragraph(
            f'<font color="{ROSE.hexval()}">–</font>  {it}',
            S("ci", font=F, size=9, sb=0, sa=1.5)))
    return lines

col1 = col([
    "Entretien individuel",
    "Diagnostic de situation",
    "Accompagnement de parcours",
    "Médiation & gestion de conflits",
    "Animation de groupe",
], "Compétences clés")

col2 = (col(["Français — natif", "Anglais — C1 (LLCER)"], "Langues")
      + [Spacer(1, 0.15*cm)]
      + col([
            "Pack Office (Word, Excel)",
            "Dispositifs : CEJ, CIVIS…",
        ], "Outils"))

col3 = col([
    "Permis B — véhiculée",
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
    # Ligne verticale de séparation entre colonnes
    ("LINEAFTER",    (0,0), (1,-1),  0.3, SEP),
    ("LEFTPADDING",  (1,0), (1,-1),  10),
    ("LEFTPADDING",  (2,0), (2,-1),  10),
]))
story.append(tbl)

# ── Rendu ─────────────────────────────────────────────────────────────────────
Frame(MX, MY, CW, H - HDR_H - MY - 0.15*cm,
      leftPadding=0, rightPadding=0,
      topPadding=0, bottomPadding=0,
      showBoundary=0).addFromList(story, c)

c.save()
print(f"PDF généré : {OUTPUT}")
