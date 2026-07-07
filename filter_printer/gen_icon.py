import os

from PIL import Image, ImageDraw

SIZE = 256
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "icon.ico")

img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# =============================
# Shadow
# =============================
draw.rounded_rectangle(
    (34, 58, 226, 220),
    radius=24,
    fill=(0, 0, 0, 35),
)

# =============================
# Folder
# =============================
draw.rounded_rectangle(
    (26, 70, 230, 208),
    radius=20,
    fill=(241, 196, 15),
)

draw.rounded_rectangle(
    (26, 48, 118, 92),
    radius=16,
    fill=(245, 203, 66),
)

# Folder highlight
draw.rectangle(
    (30, 76, 226, 86),
    fill=(255, 230, 120),
)

# =============================
# Paper
# =============================
draw.rounded_rectangle(
    (88, 64, 188, 176),
    radius=10,
    fill="white",
)

# Fold corner
draw.polygon(
    [(164,64),(188,64),(188,88)],
    fill=(230,230,230)
)

# =============================
# Text lines
# =============================
for y in [92,108,124]:
    draw.line(
        (106,y,170,y),
        fill=(180,180,180),
        width=4
    )

# =============================
# Check mark
# =============================
draw.line(
    (110,144,128,160),
    fill=(46,204,113),
    width=8
)

draw.line(
    (128,160,164,124),
    fill=(46,204,113),
    width=8
)

# =============================
# Printer
# =============================
draw.rounded_rectangle(
    (70,158,182,216),
    radius=10,
    fill=(52,73,94)
)

draw.rectangle(
    (84,140,168,170),
    fill=(90,120,160)
)

draw.rectangle(
    (92,182,160,196),
    fill="white"
)

# =============================
# Speed lines
# =============================
speed_color = (52,152,219)

draw.line(
    (186,120,228,120),
    fill=speed_color,
    width=8
)

draw.line(
    (194,142,232,142),
    fill=speed_color,
    width=6
)

draw.line(
    (202,162,228,162),
    fill=speed_color,
    width=5
)

# =============================
# Border
# =============================
draw.rounded_rectangle(
    (26,70,230,208),
    radius=20,
    outline=(210,170,0),
    width=2
)

# =============================
# Save ICO
# =============================
img.save(
    "app.ico",
    format="ICO",
    sizes=[
        (16,16),
        (24,24),
        (32,32),
        (48,48),
        (64,64),
        (128,128),
        (256,256),
    ],
)

print("Generated app.ico")