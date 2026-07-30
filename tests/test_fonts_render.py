"""Every font we write with must actually draw every character we write.

A font can claim coverage it does not have. An embedded subset lifted out of a
PDF keeps the full cmap while the glyphs the original document never used have
no outlines, so `Font.has_glyph` answers yes and the page comes out with holes
in it — a company name silently missing its W, a policy number missing its 9.
Checking the cmap is not enough; the only reliable test is to draw the glyph and
look for ink.
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import fitz
import numpy as np
import generate as g

failures = []

# Everything the generators can put on a page: company names and addresses are
# upper-cased in most templates but mixed-case on the Nganga one, plus digits and
# the punctuation that appears in VINs, policy numbers, dates and money.
REQUIRED = ("ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            "abcdefghijklmnopqrstuvwxyz"
            "0123456789"
            " -/.,#:$&'()")

FONTS = {
    "FONT_REG": g.FONT_REG,
    "FONT_BOLD": g.FONT_BOLD,
    "ARIAL_REG": g.ARIAL_REG,
    "ARIAL_BOLD": g.ARIAL_BOLD,
    "CALIBRI_REG": g.CALIBRI_REG,
    "CALIBRI_BOLD": g.CALIBRI_BOLD,
    "COC_FONT_BODY": g.COC_FONT_BODY,
    "COC_FONT_TITLE": g.COC_FONT_TITLE,
    "COC_FONT_DATE1": g.COC_FONT_DATE1,
}


def draws_ink(fontfile, ch):
    """Render one character large and report whether any dark pixel appears."""
    doc = fitz.open()
    page = doc.new_page(width=48, height=48)
    page.insert_text((8, 34), ch, fontfile=str(fontfile), fontname="probe", fontsize=26)
    samples = np.frombuffer(page.get_pixmap(dpi=72).samples, dtype=np.uint8)
    doc.close()
    return samples.min() <= 250


for label, fontfile in FONTS.items():
    if not os.path.exists(str(fontfile)):
        failures.append(f"{label}: font file missing: {fontfile}")
        continue

    blank = [ch for ch in REQUIRED if ch != " " and not draws_ink(fontfile, ch)]
    if blank:
        failures.append(
            f"{label} ({os.path.basename(str(fontfile))}): "
            f"{len(blank)} character(s) render blank: {''.join(blank)!r}"
        )

# The Nganga page is set in Calibri. Falling back to Arial is allowed, but
# silently falling back to DejaVu would be visibly wrong on that template.
calibri_face = os.path.basename(str(g.CALIBRI_REG)).lower()
if "dejavu" in calibri_face:
    failures.append(f"CALIBRI_REG resolved to a DejaVu face ({calibri_face}); "
                    f"expected Calibri, Carlito or Arial")

if failures:
    for f in failures[:10]:
        print("FAIL:", f)
    sys.exit(1)
print(f"PASS: all {len(FONTS)} fonts render every required character")
