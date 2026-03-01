"""
Recrée une approximation du logo ISFEC Bretagne avec Pillow.
Couleurs et proportions basées sur l'image originale.
"""
from PIL import Image, ImageDraw, ImageFont
import math

W, H = 500, 140
img = Image.new("RGBA", (W, H), (255, 255, 255, 255))
draw = ImageDraw.Draw(img)

# ── Couleurs ────────────────────────────────────────────────────────────────
TAUPE  = (168, 134, 92)    # "isfec" – sable chaud
BLUE   = (62,  95, 145)    # "BRETAGNE" + symbole °
RING   = (215, 220, 225)   # cercles concentriques
TEAL   = (60,  155, 160)   # sphère centrale

# ── Polices ─────────────────────────────────────────────────────────────────
try:
    font_big   = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",  74)
    font_small = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", 18)
    font_deg   = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",  22)
except Exception:
    font_big   = ImageFont.load_default()
    font_small = font_big
    font_deg   = font_big

# ── 1. Décoration : cercles concentriques (droite) ──────────────────────────
cx, cy = 428, 70          # centre de la sphère
n_rings = 14
for i in range(n_rings, 0, -1):
    r = 12 + i * 9
    alpha = max(60, 255 - i * 16)
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ov_draw = ImageDraw.Draw(overlay)
    ov_draw.ellipse([cx - r, cy - r, cx + r, cy + r],
                    outline=(*RING, alpha), width=1)
    img = Image.alpha_composite(img, overlay)

draw = ImageDraw.Draw(img)

# ── 2. Sphère centrale (dégradé simulé par cercles concentriques pleins) ────
for i in range(18, 0, -1):
    ratio = i / 18
    r_c = int(4 + (1 - ratio) * 2)        # léger reflet
    col = (
        int(TEAL[0] * ratio + 220 * (1 - ratio)),
        int(TEAL[1] * ratio + 240 * (1 - ratio)),
        int(TEAL[2] * ratio + 245 * (1 - ratio)),
    )
    draw.ellipse([cx - i, cy - i, cx + i, cy + i], fill=col)

# ── 3. Texte "BRETAGNE" en bleu (petit, au-dessus) ──────────────────────────
bretagne_text = "BRETAGNE"
bbox_b = draw.textbbox((0, 0), bretagne_text, font=font_small)
bw = bbox_b[2] - bbox_b[0]
# aligné à droite avec "isfec"
isfec_text = "isfec"
bbox_i = draw.textbbox((0, 0), isfec_text, font=font_big)
iw = bbox_i[2] - bbox_i[0]

deg_text = "°"
bbox_d = draw.textbbox((0, 0), deg_text, font=font_deg)
dw = bbox_d[2] - bbox_d[0]

x_isfec = 28          # marge gauche après le °
x_deg   = x_isfec - dw - 3

# "BRETAGNE" aligné bord droit sur bord droit de "isfec"
x_bretagne = x_isfec + iw - bw
y_bretagne = 14
draw.text((x_bretagne, y_bretagne), bretagne_text, font=font_small, fill=BLUE)

# ── 4. Symbole ° ─────────────────────────────────────────────────────────────
y_deg = 30
draw.text((x_deg, y_deg), deg_text, font=font_deg, fill=BLUE)

# ── 5. Texte principal "isfec" en taupe ──────────────────────────────────────
y_isfec = 36
draw.text((x_isfec, y_isfec), isfec_text, font=font_big, fill=TAUPE)

# ── Convertir en RGB et sauvegarder ─────────────────────────────────────────
out = img.convert("RGB")
out.save("/home/user/Lettre-de-motivation/isfec_logo.png", dpi=(150, 150))
print("Logo créé : isfec_logo.png")
