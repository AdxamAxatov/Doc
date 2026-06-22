import os
import sys
import fitz

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import generate as g

failures = []

coc = g.generate_coc("TEST SCAN CO LLC", "1 A ST, DALLAS, TX 75001",
                     dates={"start_slash": "10/16/2025", "end_slash": "10/16/2026",
                            "start_dash": "10-16-2025", "end_dash": "10-16-2026", "due": "11/6/2025"})

out = g.scannify_to_pdf(coc)
if not out.exists():
    print("FAIL: scanned PDF not created")
    sys.exit(1)
if out.suffix.lower() != ".pdf":
    failures.append(f"output is not a .pdf: {out.name}")
if "coc" in out.name.lower():
    failures.append(f"scanned filename should not contain 'COC': {out.name}")

sdoc = fitz.open(out)
if len(sdoc) != 6:
    failures.append(f"scanned PDF should have 6 pages, got {len(sdoc)}")
# Each page should carry a (scanned) image.
for i in range(len(sdoc)):
    if not sdoc[i].get_images():
        failures.append(f"scanned page {i+1} has no image")

if failures:
    for f in failures:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: scannify_to_pdf produces a 6-page scanned PDF")
