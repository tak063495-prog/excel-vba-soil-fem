"""Create guided copies of the water-input screenshots.

The source screenshots are never modified.  Red frames and small numbered
markers follow the same visual convention as the existing manual figures.
"""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1] / "docs/manuals/images/water_display"
FONT = ImageFont.load_default()

FRAMES = {
    "input_mode.png": [((228, 56, 390, 175), "1")],
    "input.png": [((225, 55, 392, 176), "1"), ((2, 318, 680, 429), "2"), ((680, 318, 1250, 380), "3")],
    "boundary.png": [((0, 190, 890, 671), "1")],
    "stage_plan.png": [((0, 318, 367, 398), "1"), ((367, 318, 930, 398), "2")],
    "stages.png": [((0, 55, 750, 351), "1")],
    "joint_input.png": [((538,104,804,158),"")],
    "polyline_input.png": [((231,57,363,172),"1"),((24,335,616,418),"2"),((740,335,1126,356),"3")],
}


def marker(draw, x, y, text):
    draw.ellipse((x - 11, y - 11, x + 11, y + 11), fill="#d00000", outline="white", width=2)
    draw.text((x - 4, y - 7), text, fill="white", font=FONT)


def main():
    for name, frames in FRAMES.items():
        src = ROOT / name
        out = ROOT / name.replace(".png", "_guided.png")
        image = Image.open(src).convert("RGB")
        draw = ImageDraw.Draw(image)
        for index, ((x1, y1, x2, y2), label) in enumerate(frames):
            draw.rectangle((x1, y1, x2, y2), outline="#d00000", width=4)
            # Keep markers in the blank margin immediately above each input area.
            if label:
                marker(draw, x1 + 13, max(12, y1 - 28), label)
        image.save(out, format="PNG", optimize=True)
        print(out)


if __name__ == "__main__":
    main()
