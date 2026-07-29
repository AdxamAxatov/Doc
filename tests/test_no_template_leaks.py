"""Generated documents must not carry the template's sample data.

Covering text with a filled rectangle only hides it — the characters stay in
the text layer and come back out of copy/paste or any text extractor. The
replacement helpers redact the original span first; this test is what stops
that quietly regressing back to cover-only.
"""

import os
import random
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import fitz
import generate as g

failures = []

COMPANY = "SAINT LOUIS FAM INC"
USDOT = "3213915"
ADDRESS = "2404 KARBA WAY, KISSIMMEE, FL 34746"

# Sample values baked into each template. None may survive into the output.
UTILITY_SAMPLES = [g.UT_NAME, "10318 CHEEVES", "HOUSTON, TX 77016"]
COC_SAMPLES = [g.COC_NAME, g.COC_NAME_P1, g.COC_ADDR1, g.COC_ADDR2,
               "10/16/2025", "10/16/2026", "10-16-2025", "10-16-2026", "11/6/2025"]
CW_SAMPLES = [g.T_COMPANY, g.T_POLICY, g.T_ADDR1, g.T_ADDR2,
              g.ALT_COMPANY_P4, g.ALT_COMPANY_REST, g.ALT_POLICY,
              g.CW_OLD_GARAGE] + list(g.CW_OLD_VINS)
NGANGA_SAMPLES = [g.NG_COMPANY, g.NG_ADDR1, g.NG_ADDR2, g.NG_VIN, g.NG_YEAR,
                  g.NG_DRIVER, g.NG_DOB, g.NG_LICENSE, g.NG_PREPARED,
                  g.NG_PERIOD, g.NGANGA_POLICY]

NGANGA_DRIVER = "Marcus Delacroix"


def text_of(path):
    doc = fitz.open(path)
    # The writer emits non-breaking spaces between words, and maps an ASCII
    # hyphen to U+2010 HYPHEN. Both render identically but break a plain
    # substring match, so normalise before the checks below.
    text = "\n".join(doc[i].get_text() for i in range(len(doc)))
    return text.replace("\xa0", " ").replace("‐", "-")


with tempfile.TemporaryDirectory() as tmp:
    out = Path(tmp)
    dates = g._coc_dates(rng=random.Random(7))

    docs = {
        "utility": (g.generate_utility(COMPANY, ADDRESS, out), UTILITY_SAMPLES),
        "coc": (g.generate_coc(COMPANY, ADDRESS, out, dates=dates), COC_SAMPLES),
        "cw": (g.generate_coverwhale(COMPANY, USDOT, ADDRESS, "CUS09116674", out,
                                     rng=random.Random(3)), CW_SAMPLES),
        "nganga": (g.generate_nganga(COMPANY, ADDRESS, NGANGA_DRIVER,
                                     "PT-26042619-01", out,
                                     rng=random.Random(5))[0], NGANGA_SAMPLES),
    }

    for name, (path, samples) in docs.items():
        text = text_of(path)

        for sample in samples:
            if sample and sample in text:
                failures.append(f"{name}: template sample still in text layer: {sample!r}")

        # Guard against the opposite failure — a redaction that strips
        # everything would otherwise pass the checks above.
        if COMPANY not in text:
            failures.append(f"{name}: company {COMPANY!r} missing from output")

    coc_text = text_of(docs["coc"][0])
    if dates["start_slash"] not in coc_text:
        failures.append(f"coc: start date {dates['start_slash']!r} missing from output")

    nganga_text = text_of(docs["nganga"][0])
    if NGANGA_DRIVER not in nganga_text:
        failures.append(f"nganga: driver {NGANGA_DRIVER!r} missing from output")
    if "PT-26042619-01" not in nganga_text:
        failures.append("nganga: policy number missing from output")

    util_text = text_of(docs["utility"][0])
    if "2404 KARBA WAY" not in util_text:
        failures.append("utility: street line missing from output")
    if "KISSIMMEE, FL 34746" not in util_text:
        failures.append("utility: city/state/zip line missing from output")

if failures:
    for f in failures[:10]:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: no template sample data leaks into generated documents")
