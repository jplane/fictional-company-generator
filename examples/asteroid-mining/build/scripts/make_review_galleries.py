import json
from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[2]
BUILD = ROOT / "build"
plan = json.loads((BUILD / "plan.json").read_text())
folder = BUILD / "renders" / "review-galleries"
folder.mkdir(exist_ok=True)
index = []
for fmt in ("docx", "xlsx"):
    docs = [d for d in plan["documents"] if d["format"] == fmt]
    for start in range(0, len(docs), 10):
        batch = docs[start:start + 10]
        canvas = Image.new("RGB", (2300, 2320), "#cbd1d4")
        draw = ImageDraw.Draw(canvas)
        for offset, d in enumerate(batch):
            image = Image.open(BUILD / "renders" / d["id"] / "contact-sheet.png")
            image.thumbnail((440, 1090))
            x, y = (offset % 5) * 460 + 10, (offset // 5) * 1160 + 10
            draw.text((x, y), f'{d["id"]} {d["type"]}\n{d["publication_date"]}', fill="black")
            canvas.paste(image, (x, y + 42))
        filename = f"{fmt}-{start // 10 + 1:02}.png"
        canvas.save(folder / filename)
        index.append({"path": str(folder / filename), "document_ids": [d["id"] for d in batch]})
(BUILD / "narrative-workbook-gallery-index.json").write_text(json.dumps(index, indent=2) + "\n")
print(json.dumps(index, indent=2))
