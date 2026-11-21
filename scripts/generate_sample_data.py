"""Generate the fictional sample inputs for this demo.

Creates, for 5 made-up stores:
  input/count_sheet_photos/store-XX.jpg  photo-style images of handwritten count sheets
  input/system_export.xlsx               what the stock system says should be on hand
  input/entered_counts.csv               the counts as entered from the photos

All stores, items and numbers are invented. Run from the project root:
    python3 scripts/generate_sample_data.py
"""

import csv
import math
import random
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = Path(__file__).resolve().parent.parent
PHOTOS = ROOT / "input" / "count_sheet_photos"
HAND_FONT = ROOT / "assets" / "fonts" / "Caveat-Regular.woff"
PRINT_FONT = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
PRINT_BOLD = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"

rng = random.Random(42)
COUNT_DATE = date(2026, 11, 2)

STORES = [
    ("S01", "Northside"),
    ("S02", "Riverside"),
    ("S03", "Downtown"),
    ("S04", "Westfield"),
    ("S05", "Lakeview"),
]

ITEMS = [
    ("10021", "AA Batteries 4-pack"),
    ("10022", "AAA Batteries 4-pack"),
    ("10345", "LED Bulb 60W equiv."),
    ("10346", "LED Bulb 100W equiv."),
    ("11002", "Extension Cord 3m"),
    ("11003", "Extension Cord 5m"),
    ("12510", "Duct Tape 50m"),
    ("12511", "Masking Tape 25mm"),
    ("13001", "Screwdriver Set 6pc"),
    ("13020", "Claw Hammer 16oz"),
    ("13044", "Tape Measure 5m"),
    ("14100", "Work Gloves (L)"),
    ("14101", "Work Gloves (M)"),
    ("15200", "Paint Brush 50mm"),
    ("15201", "Paint Roller Kit"),
    ("16010", "Wood Glue 250ml"),
    ("16055", "Super Glue 3g"),
    ("17300", "Cable Ties 100pk"),
]
ITEM_NAMES = dict(ITEMS)

# Extra item a store wrote in by hand that is not in its system export.
WRITE_IN = {"S03": ("18800", "Flashlight LED (write-in)", 7)}

COUNTERS = ["J.M.", "A.R.", "K.T.", "D.L.", "S.P."]


def plan_store(code):
    """Return list of rows: dict(sku, desc, system, counted, kind, seen, note)."""
    rows = []
    diff_skus = set(rng.sample([s for s, _ in ITEMS], 3))
    skipped = rng.choice([s for s, _ in ITEMS]) if code in ("S02", "S05") else None
    problem_skus = rng.sample(
        [s for s, _ in ITEMS if s not in diff_skus and s != skipped], 1
    )
    for sku, desc in ITEMS:
        system = (
            rng.choice([0, 2, 4, 6]) if rng.random() < 0.08 else rng.randint(5, 120)
        )
        counted = system
        if sku in diff_skus:
            delta = rng.choice([-12, -6, -3, -2, -1, 1, 2, 4, 10])
            counted = max(0, system + delta)
        row = dict(
            sku=sku,
            desc=desc,
            system=system,
            counted=counted,
            kind="clear",
            seen="",
            note="",
        )
        if sku == skipped:
            row.update(counted=None, kind="blank", note="Line left blank on sheet")
        elif sku in problem_skus:
            kind = "crossed" if code in ("S01", "S03", "S05") else "smudged"
            row["kind"] = kind
            if kind == "crossed":
                old = counted + rng.choice([-10, 10, 3])
                row["seen"] = f"{old} crossed out, {counted} written next to it"
                row["note"] = "Used the corrected number - confirm with store"
                row["old"] = old
            else:
                row["seen"] = f"Smudged - looks like {counted} (last digit unclear)"
                row["note"] = "Best reading entered - confirm with store"
        rows.append(row)
    if code in WRITE_IN:
        sku, desc, qty = WRITE_IN[code]
        rows.append(
            dict(
                sku=sku,
                desc=desc,
                system=None,
                counted=qty,
                kind="writein",
                seen="Handwritten extra line",
                note="Item not on printed sheet or in system export",
            )
        )
    return rows


def hand(draw, xy, text, size=44, fill=(28, 45, 120)):
    font = ImageFont.truetype(str(HAND_FONT), size + rng.randint(-3, 3))
    x, y = xy
    for ch in text:
        draw.text((x, y + rng.randint(-2, 2)), ch, font=font, fill=fill)
        x += font.getlength(ch) + rng.uniform(-0.5, 1.5)
    return x


def render_sheet(code, name, counter, rows, path):
    W, H = 1240, 1754  # A4 at 150 dpi
    page = Image.new("RGB", (W, H), (250, 248, 240))
    d = ImageDraw.Draw(page)
    f_title = ImageFont.truetype(PRINT_BOLD, 40)
    f_label = ImageFont.truetype(PRINT_FONT, 24)
    f_head = ImageFont.truetype(PRINT_BOLD, 22)
    f_cell = ImageFont.truetype(PRINT_FONT, 22)
    ink = (35, 35, 35)

    d.text((80, 70), "STOCK COUNT SHEET", font=f_title, fill=ink)
    d.text((80, 140), "Store:", font=f_label, fill=ink)
    d.line((170, 168, 600, 168), fill=ink, width=1)
    hand(d, (185, 118), f"{code} {name}")
    d.text((660, 140), "Date:", font=f_label, fill=ink)
    d.line((735, 168, 1000, 168), fill=ink, width=1)
    hand(d, (750, 118), COUNT_DATE.strftime("%d/%m/%Y"))
    d.text((80, 200), "Counted by:", font=f_label, fill=ink)
    d.line((235, 228, 600, 228), fill=ink, width=1)
    hand(d, (250, 178), counter)

    cols = [80, 160, 300, 880, 1160]
    top, rh = 280, 64
    headers = ["#", "SKU", "Description", "Qty counted"]
    n_rows = len(ITEMS) + 3
    for i, h in enumerate(headers):
        d.text((cols[i] + 12, top + 20), h, font=f_head, fill=ink)
    for r in range(n_rows + 2):
        y = top + r * rh
        d.line((cols[0], y, cols[-1], y), fill=ink, width=2 if r < 2 else 1)
    for x in cols:
        d.line((x, top, x, top + (n_rows + 1) * rh), fill=ink, width=1)

    for i, row in enumerate(rows):
        y = top + (i + 1) * rh
        is_printed = row["kind"] != "writein"
        d.text((cols[0] + 12, y + 20), str(i + 1), font=f_cell, fill=ink)
        if is_printed:
            d.text((cols[1] + 12, y + 20), row["sku"], font=f_cell, fill=ink)
            d.text((cols[2] + 12, y + 20), row["desc"], font=f_cell, fill=ink)
        else:
            hand(d, (cols[1] + 10, y + 6), row["sku"], size=40)
            hand(d, (cols[2] + 10, y + 6), "Flashlight LED", size=40)
        qx, qy = cols[3] + 40 + rng.randint(-10, 20), y + 4
        if row["kind"] == "blank":
            continue
        if row["kind"] == "crossed":
            end = hand(d, (qx, qy), str(row["old"]))
            d.line((qx - 6, qy + 34, end + 6, qy + 24), fill=(28, 45, 120), width=4)
            hand(d, (end + 30, qy), str(row["counted"]))
        else:
            end = hand(d, (qx, qy), str(row["counted"]))
            if row["kind"] == "smudged":
                box = (int(qx) - 10, qy + 5, int(end) + 15, qy + rh - 5)
                region = page.crop(box).filter(ImageFilter.GaussianBlur(5))
                page.paste(region, box)

    d.text((80, top + (n_rows + 1) * rh + 30), "Signature:", font=f_label, fill=ink)
    hand(d, (220, top + (n_rows + 1) * rh + 8), counter.replace(".", ""), size=52)

    # Make it look like a phone photo: on a table, slightly rotated, uneven light.
    bg = Image.new("RGB", (W + 240, H + 240), (92, 74, 58))
    angle = rng.uniform(-2.2, 2.2)
    rotated = page.rotate(
        angle, resample=Image.BICUBIC, expand=True, fillcolor=(92, 74, 58)
    )
    bg.paste(rotated, (120 - (rotated.width - W) // 2, 120 - (rotated.height - H) // 2))
    shade = Image.new("L", bg.size)
    sd = ImageDraw.Draw(shade)
    cx, cy = rng.randint(200, bg.width - 200), rng.randint(200, bg.height - 200)
    maxd = math.hypot(bg.width, bg.height)
    for r in range(int(maxd), 0, -40):
        v = int(255 - 70 * (r / maxd))
        sd.ellipse((cx - r, cy - r, cx + r, cy + r), fill=v)
    shade = shade.filter(ImageFilter.GaussianBlur(60))
    black = Image.new("RGB", bg.size, (0, 0, 0))
    photo = Image.composite(bg, black, shade).filter(ImageFilter.GaussianBlur(0.6))
    photo = photo.resize((photo.width * 3 // 4, photo.height * 3 // 4), Image.LANCZOS)
    photo.save(path, quality=82)


def main():
    PHOTOS.mkdir(parents=True, exist_ok=True)
    entered, system_rows = [], []
    for (code, name), counter in zip(STORES, COUNTERS):
        rows = plan_store(code)
        photo = f"store-{code[1:]}.jpg"
        render_sheet(code, name, counter, rows, PHOTOS / photo)
        for line, row in enumerate(rows, start=1):
            if row["system"] is not None:
                system_rows.append((code, name, row["sku"], row["desc"], row["system"]))
            if row["counted"] is not None:
                entered.append(
                    dict(
                        store=code,
                        sku=row["sku"],
                        description=row["desc"],
                        counted_qty=row["counted"],
                        photo=photo,
                        sheet_line=line,
                        reading="Check"
                        if row["kind"] in ("crossed", "smudged", "writein")
                        else "Clear",
                        what_was_seen=row["seen"],
                        note=row["note"],
                    )
                )
            elif row["kind"] == "blank":
                entered.append(
                    dict(
                        store=code,
                        sku=row["sku"],
                        description=row["desc"],
                        counted_qty="",
                        photo=photo,
                        sheet_line=line,
                        reading="Check",
                        what_was_seen="No quantity written",
                        note=row["note"],
                    )
                )

    with open(ROOT / "input" / "entered_counts.csv", "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(entered[0].keys()))
        w.writeheader()
        w.writerows(entered)

    wb = Workbook()
    ws = wb.active
    ws.title = "Stock On Hand"
    ws.append(["Store", "Store Name", "SKU", "Description", "On Hand Qty"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="555555")
    for r in system_rows:
        ws.append(list(r))
    for col, width in zip("ABCDE", (8, 14, 10, 28, 12)):
        ws.column_dimensions[col].width = width
    wb.save(ROOT / "input" / "system_export.xlsx")
    print(
        f"{len(STORES)} photos, {len(entered)} entered lines, {len(system_rows)} system lines"
    )


if __name__ == "__main__":
    main()
