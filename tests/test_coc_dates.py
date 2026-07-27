import os
import random
import re
import sys
import calendar
import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import generate as g

failures = []

# Read the term from the implementation rather than repeating it. This test
# previously hardcoded October 2025 and went red the moment the term shifted to
# April 2026 — the values live in one place now so that cannot recur.
YEAR = g.COC_TERM_YEAR
MONTH = g.COC_TERM_MONTH
DAYS_IN_MONTH = calendar.monthrange(YEAR, MONTH)[1]
DAYS_IN_END_MONTH = calendar.monthrange(YEAR + 1, MONTH)[1]

# Deterministic with a seeded RNG; sweep many seeds to cover every day of the month.
seen_days = set()
for seed in range(200):
    d = g._coc_dates(rng=random.Random(seed))
    s = datetime.datetime.strptime(d["start_slash"], "%m/%d/%Y").date()
    e = datetime.datetime.strptime(d["end_slash"], "%m/%d/%Y").date()
    due = datetime.datetime.strptime(d["due"], "%m/%d/%Y").date()
    seen_days.add(s.day)

    if not (s.year == YEAR and s.month == MONTH and 1 <= s.day <= DAYS_IN_MONTH):
        failures.append(f"seed {seed}: start not in {MONTH:02d}/{YEAR}: {d['start_slash']}")
    # Same day one year on, except a Feb 29 start which clamps to Feb 28.
    if not (e.year == YEAR + 1 and e.month == MONTH
            and e.day == min(s.day, DAYS_IN_END_MONTH)):
        failures.append(f"seed {seed}: end not same day {MONTH:02d}/{YEAR + 1}: {d['end_slash']}")
    if due != s + datetime.timedelta(days=21):
        failures.append(f"seed {seed}: due != start+21d: {d['due']}")
    if not re.match(rf"^{MONTH:02d}/\d{{2}}/{YEAR}$", d["start_slash"]):
        failures.append(f"seed {seed}: start_slash format {d['start_slash']}")
    if not re.match(rf"^{MONTH:02d}-\d{{2}}-{YEAR}$", d["start_dash"]):
        failures.append(f"seed {seed}: start_dash format {d['start_dash']}")
    if not re.match(rf"^{MONTH:02d}-\d{{2}}-{YEAR + 1}$", d["end_dash"]):
        failures.append(f"seed {seed}: end_dash format {d['end_dash']}")
    # The template prints the due date without a zero-padded day.
    if re.search(r"/0\d/", d["due"]):
        failures.append(f"seed {seed}: due has leading-zero day: {d['due']}")

# The sweep is only meaningful if it actually exercised the whole month.
if seen_days != set(range(1, DAYS_IN_MONTH + 1)):
    missing = sorted(set(range(1, DAYS_IN_MONTH + 1)) - seen_days)
    failures.append(f"200 seeds did not cover every start day; missing {missing}")

# Sample reproduction: a start on the 16th must produce the exact strings the
# template expects, derived from the same term constants.
sample_start = datetime.date(YEAR, MONTH, 16)
sample_end = datetime.date(YEAR + 1, MONTH, 16)
sample_due = sample_start + datetime.timedelta(days=21)
want_start = sample_start.strftime("%m/%d/%Y")

hit16 = next((g._coc_dates(rng=random.Random(s)) for s in range(500)
              if g._coc_dates(rng=random.Random(s))["start_slash"] == want_start), None)
if hit16 is None:
    failures.append(f"no seed in range(500) produced a start of {want_start}")
else:
    want_end_dash = sample_end.strftime("%m-%d-%Y")
    want_due = f"{sample_due.month}/{sample_due.day}/{sample_due.year}"
    if hit16["end_dash"] != want_end_dash or hit16["due"] != want_due:
        failures.append(
            f"16th sample mismatch: got end_dash={hit16['end_dash']} due={hit16['due']}, "
            f"want end_dash={want_end_dash} due={want_due}"
        )

if failures:
    for f in failures[:10]:
        print("FAIL:", f)
    sys.exit(1)
print(f"PASS: _coc_dates ({MONTH:02d}/{YEAR} start, same-day +1yr term, +21d due, formats)")
