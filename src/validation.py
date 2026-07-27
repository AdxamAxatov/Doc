"""
Input shape validators for the Telegram bot.

Pure functions with no third-party imports, so they can be exercised without a
Telegram token or the python-telegram-bot package installed.

Each validator answers one question: "is this value shaped like the thing we
asked for?"  It returns a user-facing error message when the answer is no, and
None when the value is acceptable.  It does not check that a company or address
actually exists.
"""

import re

EMAIL_RE = re.compile(r"[^@\s]+@[^@\s]+\.[A-Za-z]{2,}")
URL_RE = re.compile(r"(?:https?://|www\.)\S+", re.I)
ZIP_RE = re.compile(r"\b\d{5}(?:-\d{4})?\b")
STATE_RE = re.compile(
    r"\b(?:A[KLRZ]|C[AOT]|D[CE]|FL|GA|HI|I[ADLN]|K[SY]|LA|M[ADEINOST]|"
    r"N[CDEHJMVY]|O[HKR]|P[AR]|RI|S[CD]|T[NX]|UT|V[AT]|W[AIVY])\b",
    re.I,
)

ADDR_EXAMPLE = "Example: 1234 Main St, Houston, TX 77016"


def looks_like_company(s: str) -> str | None:
    """Error message if `s` is not shaped like a company name, else None."""
    s = (s or "").strip()
    if len(s) < 2:
        return "That's too short for a company name. Please enter the company name."
    if EMAIL_RE.search(s):
        return ("That looks like an email address, not a company name.\n"
                "Please enter the company name (e.g. LEO EMPIRE SERVICES LLC).")
    if URL_RE.search(s):
        return ("That looks like a website, not a company name.\n"
                "Please enter the company name.")
    # Only a total absence of letters is disqualifying. A "more digits than
    # letters" rule was tried and rejected: it wrongly blocked the real
    # companies 7573 LLC and 1524 INC, and caught nothing this check misses.
    if not any(c.isalpha() for c in s):
        return ("That looks like a number, not a company name.\n"
                "Please enter the company name.")
    return None


def looks_like_address(s: str) -> str | None:
    """Error message if `s` is not shaped like a mailing address, else None."""
    s = (s or "").strip()
    if len(s) < 6:
        return f"That's too short for an address. {ADDR_EXAMPLE}"
    if EMAIL_RE.search(s):
        return f"That looks like an email address, not an address.\n{ADDR_EXAMPLE}"
    if not any(c.isdigit() for c in s):
        return ("That doesn't look like an address — no street number or ZIP code.\n"
                f"{ADDR_EXAMPLE}")
    # Any one of these is enough: a street line alone is what we are rejecting.
    if not ("," in s or "\n" in s or ZIP_RE.search(s) or STATE_RE.search(s)):
        return ("That doesn't look like a full address — it needs a city and state.\n"
                f"{ADDR_EXAMPLE}")
    return None
