import os
import random
import re
import sys
import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import generate as g

failures = []

# Deterministic with a seeded RNG; sweep many seeds to cover all October days.
for seed in range(200):
    d = g._coc_dates(rng=random.Random(seed))
    s = datetime.datetime.strptime(d["start_slash"], "%m/%d/%Y").date()
    e = datetime.datetime.strptime(d["end_slash"], "%m/%d/%Y").date()
    due = datetime.datetime.strptime(d["due"], "%m/%d/%Y").date()
    if not (s.year == 2025 and s.month == 10 and 1 <= s.day <= 31):
        failures.append(f"seed {seed}: start not in Oct 2025: {d['start_slash']}")
    if not (e.year == 2026 and e.month == 10 and e.day == s.day):
        failures.append(f"seed {seed}: end not same day Oct 2026: {d['end_slash']}")
    if due != s + datetime.timedelta(days=21):
        failures.append(f"seed {seed}: due != start+21d: {d['due']}")
    if not re.match(r"^10/\d{2}/2025$", d["start_slash"]):
        failures.append(f"seed {seed}: start_slash format {d['start_slash']}")
    if not re.match(r"^10-\d{2}-2025$", d["start_dash"]):
        failures.append(f"seed {seed}: start_dash format {d['start_dash']}")
    if not re.match(r"^10-\d{2}-2026$", d["end_dash"]):
        failures.append(f"seed {seed}: end_dash format {d['end_dash']}")
    if re.search(r"/0\d/", d["due"]):
        failures.append(f"seed {seed}: due has leading-zero day: {d['due']}")

# Sample reproduction: a seed that yields the 16th must match the template exactly.
hit16 = next((g._coc_dates(rng=random.Random(s)) for s in range(500)
              if g._coc_dates(rng=random.Random(s))["start_slash"] == "10/16/2025"), None)
if hit16:
    if hit16["end_dash"] != "10-16-2026" or hit16["due"] != "11/6/2025":
        failures.append(f"16th sample mismatch: {hit16}")

if failures:
    for f in failures[:10]:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: _coc_dates (Oct 2025 start, same-day +1yr term, +21d due, formats)")
