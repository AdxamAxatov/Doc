import os
import sys
import fitz

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import generate as g

failures = []

COMPANY = "ACME TEST CARRIER LLC"
ADDRESS = "500 TEST BLVD, MIAMI, FL 33101"
dates = {"start_slash": "10/16/2025", "end_slash": "10/16/2026",
         "start_dash": "10-16-2025", "end_dash": "10-16-2026", "due": "11/6/2025"}

path = g.generate_coc(COMPANY, ADDRESS, dates=dates)
if not path.exists():
    print("FAIL: output PDF not created")
    sys.exit(1)

doc = fitz.open(path)
if len(doc) != 6:
    failures.append(f"expected 6 pages, got {len(doc)}")

texts = [doc[i].get_text() for i in range(len(doc))]
# Pages 2 & 5 use DejaVuSans (get_text-extractable). Page 1 uses Arial via
# insert_text, which isn't reliably extractable as plain text — page 1 is
# verified visually, not here.
for pidx in (1, 4):
    if "ACME TEST CARRIER LLC" not in texts[pidx]:
        failures.append(f"company not on page {pidx+1}")
if "10/16/2025" not in texts[1]:
    failures.append("start date not on page 2")
if "500 TEST BLVD" not in texts[1]:
    failures.append("new address street not on page 2")
joined = "\n".join(texts)
if "Crum & Forster" in joined:
    failures.append("pages 7-12 (Crum & Forster binder) were not dropped")

if failures:
    for f in failures:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: generate_coc — 6 pages, company/date/address filled, binder dropped")
