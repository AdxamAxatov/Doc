import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
from validation import looks_like_company, looks_like_address

failures = []

# Real company names sampled from assets/All Companies.csv, plus the incident value.
GOOD_COMPANIES = [
    "SAINT LOUIS FAM INC",
    "LALO'S TRUCKING INC",
    "1 GUY 1 GIRL 1 TRUCK LLC",
    "LEO EMPIRE SERVICES LLC",
    "BEE LOGISTICS SOLUTIONS LLC",
    "D-TAG INCORPORATION",
    "GKH",
    # Regression: real companies with more digits than letters. An earlier
    # draft of looks_like_company rejected both.
    "7573 LLC",
    "1524 INC",
]

BAD_COMPANIES = [
    "daylenis@leoempireservicesllc.com",   # the reported incident
    "someone@example.org",
    "https://leoempireservicesllc.com",
    "www.leoempire.com",
    "8005551234",
    "3213915",
    "",
    "X",
]

# Manual entries and the newline form stored in the CSV.
GOOD_ADDRESSES = [
    "265 Faulkner Dr, Niantic, IL 62551",
    "8514 Fenway Dr Houston TX 77036",
    "10318 CHEEVES, HOUSTON, TX 77016",
    "2404 KARBA WAY\n \n KISSIMMEE, FL 34746",
    "721 N MIRASOL AVE, MESA, AZ 85207",
]

BAD_ADDRESSES = [
    "LEO EMPIRE SERVICES LLC",             # the reported incident
    "SAINT LOUIS FAM INC",
    "daylenis@leoempireservicesllc.com",
    "no digits here at all",
    "",
    "12",
]

for s in GOOD_COMPANIES:
    err = looks_like_company(s)
    if err is not None:
        failures.append(f"company wrongly rejected: {s!r} -> {err!r}")

for s in BAD_COMPANIES:
    if looks_like_company(s) is None:
        failures.append(f"company wrongly accepted: {s!r}")

for s in GOOD_ADDRESSES:
    err = looks_like_address(s)
    if err is not None:
        failures.append(f"address wrongly rejected: {s!r} -> {err!r}")

for s in BAD_ADDRESSES:
    if looks_like_address(s) is None:
        failures.append(f"address wrongly accepted: {s!r}")

# Error messages must be non-empty strings a user can act on.
for bad, fn in [(BAD_COMPANIES, looks_like_company), (BAD_ADDRESSES, looks_like_address)]:
    for s in bad:
        err = fn(s)
        if not isinstance(err, str) or not err.strip():
            failures.append(f"{fn.__name__}({s!r}) returned unusable error {err!r}")

if failures:
    for f in failures[:10]:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: input validation (company/address shape checks)")
