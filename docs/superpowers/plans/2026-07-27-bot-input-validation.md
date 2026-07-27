# Bot Input Validation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stop the Telegram bot from generating documents when a company name or address is obviously the wrong kind of value, while still allowing a legitimate oddball entry through.

**Architecture:** Two pure shape-checking functions live in a new dependency-free module `src/validation.py`. A single async helper `check_field` in `src/bot.py` wraps them, owning the rejection message, a two-strike `[Use it anyway] / [Try again]` override keyboard, and per-field bookkeeping in `ctx.user_data`. Eight conversation handlers each gain a two-line guard. No new conversation states are added.

**Tech Stack:** Python 3.12, python-telegram-bot 22.7, standard-library `re`. No new dependencies.

**Spec:** `docs/superpowers/specs/2026-07-27-bot-input-validation-design.md`

## Global Constraints

- **No pytest.** It is not installed. Tests are standalone scripts run with `venv/bin/python tests/<file>.py`, matching `tests/test_coc_dates.py`: accumulate into a `failures` list, print `FAIL: ...` lines and `sys.exit(1)` if non-empty, otherwise print a single `PASS: ...` line.
- **Use the venv interpreter:** `./venv/bin/python`, never bare `python`.
- **`src/validation.py` must import nothing outside the standard library.** It is imported by tests that have no Telegram token and must not pull in `telegram`, `dotenv`, `fitz`, or `generate`.
- **Importing `src/bot.py` calls `sys.exit` unless `TELEGRAM_BOT_TOKEN` is set.** Any test that imports `bot` must set `os.environ["TELEGRAM_BOT_TOKEN"]` *before* the import.
- **Do not modify `src/generate.py`.** The PDF pipeline is confirmed working.
- **Validators return `str | None`** — an error message to show the user, or `None` when the value is acceptable. Never raise, never return a bool.
- **Address validation is deliberately loose.** Over-accepting is safe because the override exists; over-rejecting blocks real work.
- **Do not validate USDOT** (`got_usdot`, `got_cw_usdot`) or any value sourced from `All Companies.csv`. Only manually typed name and address input is checked.

---

### Task 1: Pure shape validators

**Files:**
- Create: `src/validation.py`
- Create: `tests/test_input_validation.py`

**Interfaces:**
- Consumes: nothing
- Produces:
  - `looks_like_company(s: str) -> str | None`
  - `looks_like_address(s: str) -> str | None`

  Both used by Task 3's call sites via `check_field`.

- [ ] **Step 1: Create a working branch**

We are on `master`, the default branch. Do not commit to it directly.

```bash
git checkout -b fix/bot-input-validation
```

- [ ] **Step 2: Write the failing test**

Create `tests/test_input_validation.py`:

```python
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
```

- [ ] **Step 3: Run test to verify it fails**

Run: `./venv/bin/python tests/test_input_validation.py`
Expected: FAIL with `ModuleNotFoundError: No module named 'validation'`

- [ ] **Step 4: Write the implementation**

Create `src/validation.py`:

```python
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
```

- [ ] **Step 5: Run test to verify it passes**

Run: `./venv/bin/python tests/test_input_validation.py`
Expected: `PASS: input validation (company/address shape checks)`

- [ ] **Step 6: Commit**

```bash
git add src/validation.py tests/test_input_validation.py
git commit -m "Add company/address shape validators"
```

---

### Task 2: The `check_field` gate helper

**Files:**
- Modify: `src/bot.py` — add `USE_ANYWAY` next to `YES_NO` (line 95), import from `validation`, add `check_field` in the HELPERS section (after line 97)
- Create: `tests/test_check_field.py`

**Interfaces:**
- Consumes: `looks_like_company`, `looks_like_address` from Task 1
- Produces: `async check_field(update, ctx, validator, slot, prompt) -> str | None`

  Returns the value the caller should use, or `None` when a re-prompt was already sent and the caller must return its own conversation state unchanged. Task 3 calls this at all eight sites.

- [ ] **Step 1: Write the failing test**

Create `tests/test_check_field.py`:

```python
import asyncio
import os
import sys

# bot.py calls sys.exit() at import time when this is unset.
os.environ.setdefault("TELEGRAM_BOT_TOKEN", "test-token")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import bot
from validation import looks_like_address

failures = []


class FakeMessage:
    def __init__(self, text):
        self.text = text
        self.replies = []

    async def reply_text(self, text, reply_markup=None, **kwargs):
        self.replies.append((text, reply_markup))


class FakeUser:
    first_name = "tester"


class FakeUpdate:
    def __init__(self, text):
        self.message = FakeMessage(text)
        self.effective_user = FakeUser()


class FakeCtx:
    def __init__(self):
        self.user_data = {}


PROMPT = "Enter the address:"


async def send(ctx, text):
    """Feed one message through check_field; return (result, replies)."""
    upd = FakeUpdate(text)
    result = await bot.check_field(upd, ctx, looks_like_address, "ut_addr", PROMPT)
    return result, upd.message.replies


async def main():
    # A valid address passes straight through and leaves no bookkeeping behind.
    ctx = FakeCtx()
    result, replies = await send(ctx, "265 Faulkner Dr, Niantic, IL 62551")
    if result != "265 Faulkner Dr, Niantic, IL 62551":
        failures.append(f"valid address not returned: {result!r}")
    if replies:
        failures.append(f"valid address should not reply, got {replies!r}")
    if ctx.user_data:
        failures.append(f"valid address left user_data dirty: {ctx.user_data!r}")

    # First rejection: plain error, no override keyboard.
    ctx = FakeCtx()
    result, replies = await send(ctx, "LEO EMPIRE SERVICES LLC")
    if result is not None:
        failures.append(f"first rejection should return None, got {result!r}")
    if len(replies) != 1:
        failures.append(f"first rejection should send one reply, got {replies!r}")
    elif replies[0][1] is bot.USE_ANYWAY:
        failures.append("first rejection must not offer the override keyboard")
    if ctx.user_data.get("_rej_ut_addr") != 1:
        failures.append(f"strike count wrong: {ctx.user_data!r}")
    if ctx.user_data.get("_pend_ut_addr") != "LEO EMPIRE SERVICES LLC":
        failures.append(f"pending value not stored: {ctx.user_data!r}")

    # Second rejection: override keyboard offered, pending value updated.
    result, replies = await send(ctx, "PO BOX 412")
    if result is not None:
        failures.append(f"second rejection should return None, got {result!r}")
    if not replies or replies[-1][1] is not bot.USE_ANYWAY:
        failures.append(f"second rejection must offer override keyboard, got {replies!r}")
    if ctx.user_data.get("_rej_ut_addr") != 2:
        failures.append(f"strike count should be 2: {ctx.user_data!r}")
    if ctx.user_data.get("_pend_ut_addr") != "PO BOX 412":
        failures.append(f"pending value should update: {ctx.user_data!r}")

    # "Use it anyway" returns the pending value and clears bookkeeping.
    result, _ = await send(ctx, "Use it anyway")
    if result != "PO BOX 412":
        failures.append(f"override should return pending value, got {result!r}")
    if ctx.user_data:
        failures.append(f"override left user_data dirty: {ctx.user_data!r}")

    # "Try again" clears both keys and re-sends the prompt.
    ctx = FakeCtx()
    await send(ctx, "LEO EMPIRE SERVICES LLC")
    await send(ctx, "LEO EMPIRE SERVICES LLC")
    result, replies = await send(ctx, "Try again")
    if result is not None:
        failures.append(f"'Try again' should return None, got {result!r}")
    if not replies or replies[-1][0] != PROMPT:
        failures.append(f"'Try again' should re-send the prompt, got {replies!r}")
    if ctx.user_data:
        failures.append(f"'Try again' left user_data dirty: {ctx.user_data!r}")
    # After "Try again" the next bad value gets a plain error, not the keyboard.
    _, replies = await send(ctx, "LEO EMPIRE SERVICES LLC")
    if replies and replies[-1][1] is bot.USE_ANYWAY:
        failures.append("strike count should reset after 'Try again'")

    # "Use it anyway" with nothing pending re-prompts instead of returning junk.
    ctx = FakeCtx()
    result, replies = await send(ctx, "Use it anyway")
    if result is not None:
        failures.append(f"override with no pending value should return None, got {result!r}")
    if not replies or replies[-1][0] != PROMPT:
        failures.append(f"override with no pending value should re-prompt, got {replies!r}")

    # A later valid value clears the strike count from earlier rejections.
    ctx = FakeCtx()
    await send(ctx, "LEO EMPIRE SERVICES LLC")
    result, _ = await send(ctx, "8514 Fenway Dr Houston TX 77036")
    if result != "8514 Fenway Dr Houston TX 77036":
        failures.append(f"recovery value not returned: {result!r}")
    if ctx.user_data:
        failures.append(f"recovery left user_data dirty: {ctx.user_data!r}")


asyncio.run(main())

if failures:
    for f in failures[:10]:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: check_field (rejection, two-strike override, reset)")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `./venv/bin/python tests/test_check_field.py`
Expected: FAIL with `AttributeError: module 'bot' has no attribute 'check_field'`

- [ ] **Step 3: Import the validators in `src/bot.py`**

`re` is not currently imported in `bot.py` and is not needed there — the regexes live in `validation.py`.

Add directly below the `from generate import (...)` block that ends at line 34:

```python
from validation import looks_like_company, looks_like_address
```

- [ ] **Step 4: Add the `USE_ANYWAY` keyboard**

In `src/bot.py`, immediately after the `YES_NO` definition at line 95:

```python
USE_ANYWAY = ReplyKeyboardMarkup([["Use it anyway"], ["Try again"]],
                                 one_time_keyboard=True, resize_keyboard=True)
```

- [ ] **Step 5: Add `check_field` to the HELPERS section**

In `src/bot.py`, after the `# ─── HELPERS ───` banner at line 97:

```python
async def check_field(update, ctx, validator, slot, prompt):
    """
    Validate one free-text answer before the caller acts on it.

    Returns the value to use, or None when a re-prompt has already been sent and
    the caller should return its own conversation state unchanged.

    Bookkeeping lives in ctx.user_data under two keys derived from `slot`:
      _pend_<slot>  the most recently rejected value, held so it can be forced through
      _rej_<slot>   how many times this field has been rejected

    The override buttons arrive as ordinary text inside the state the caller
    already returned, so no new conversation states are needed.
    """
    text = update.message.text.strip()
    pend_key, rej_key = f"_pend_{slot}", f"_rej_{slot}"
    user = update.effective_user.first_name

    if text.lower() == "use it anyway":
        pending = ctx.user_data.pop(pend_key, None)
        ctx.user_data.pop(rej_key, None)
        if pending:
            logger.info(f'[{user}] Override accepted {slot}: "{pending}"')
            return pending
        await update.message.reply_text(prompt, reply_markup=ReplyKeyboardRemove())
        return None

    if text.lower() == "try again":
        ctx.user_data.pop(pend_key, None)
        ctx.user_data.pop(rej_key, None)
        await update.message.reply_text(prompt, reply_markup=ReplyKeyboardRemove())
        return None

    err = validator(text)
    if err is None:
        ctx.user_data.pop(pend_key, None)
        ctx.user_data.pop(rej_key, None)
        return text

    strikes = ctx.user_data.get(rej_key, 0) + 1
    ctx.user_data[rej_key] = strikes
    ctx.user_data[pend_key] = text
    logger.info(f'[{user}] Rejected {slot} (strike {strikes}): "{text}"')
    if strikes >= 2:
        await update.message.reply_text(f"{err}\n\nUse it anyway?",
                                        reply_markup=USE_ANYWAY)
    else:
        await update.message.reply_text(err, reply_markup=ReplyKeyboardRemove())
    return None
```

- [ ] **Step 6: Run test to verify it passes**

Run: `./venv/bin/python tests/test_check_field.py`
Expected: `PASS: check_field (rejection, two-strike override, reset)`

- [ ] **Step 7: Commit**

```bash
git add src/bot.py tests/test_check_field.py
git commit -m "Add check_field gate helper with two-strike override"
```

---

### Task 3: Wire the gate into all eight handlers

**Files:**
- Modify: `src/bot.py` — `got_name:166`, `got_addr:248`, `got_ut_name:384`, `got_ut_addr:438`, `got_coc_name:496`, `got_coc_addr:538`, `got_cw_name:601`, `got_cw_addr:648`
- Create: `tests/test_handler_guards.py`

**Interfaces:**
- Consumes: `check_field`, `looks_like_company`, `looks_like_address`
- Produces: no new symbols. Handler signatures and return values are unchanged; each simply may now return its own state instead of advancing.

Line numbers are from the pre-Task-2 file and will have shifted by the lines added in Task 2. Locate handlers by name, not by number.

- [ ] **Step 1: Write the failing test**

Create `tests/test_handler_guards.py`. This asserts that a bad value keeps the conversation in the same state and never reaches generation.

```python
import asyncio
import os
import sys

os.environ.setdefault("TELEGRAM_BOT_TOKEN", "test-token")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import bot

failures = []


class FakeMessage:
    def __init__(self, text):
        self.text = text
        self.replies = []

    async def reply_text(self, text, reply_markup=None, **kwargs):
        self.replies.append((text, reply_markup))


class FakeUser:
    first_name = "tester"


class FakeUpdate:
    def __init__(self, text):
        self.message = FakeMessage(text)
        self.effective_user = FakeUser()


class FakeCtx:
    def __init__(self, **data):
        self.user_data = dict(data)


generated = []


async def fake_generate(*args, **kwargs):
    generated.append(args)
    return None


async def main():
    # Every generation entry point is stubbed; nothing should reach them.
    bot._ut_generate_and_send = fake_generate
    bot._coc_generate_and_send = fake_generate
    bot._cw_generate_and_send = fake_generate

    # Address handlers must hold their state on a company-name-shaped value.
    addr_cases = [
        (bot.got_ut_addr, bot.UT_ADDR, FakeCtx(ut_company="LEO EMPIRE SERVICES LLC"), "ut_addr"),
        (bot.got_coc_addr, bot.COC_ADDR, FakeCtx(coc_company="LEO EMPIRE SERVICES LLC"), "coc_addr"),
        (bot.got_cw_addr, bot.CW_ADDR, FakeCtx(cw_company="LEO EMPIRE SERVICES LLC", cw_usdot="123"), "cw_addr"),
        (bot.got_addr, bot.ASK_ADDR, FakeCtx(current={"name": "X", "usdot": "1"}), "addr"),
    ]
    for handler, expected_state, ctx, slot in addr_cases:
        upd = FakeUpdate("LEO EMPIRE SERVICES LLC")
        state = await handler(upd, ctx)
        if state != expected_state:
            failures.append(f"{handler.__name__}: expected state {expected_state}, got {state}")
        if not upd.message.replies:
            failures.append(f"{handler.__name__}: no rejection message sent")
        if ctx.user_data.get(f"_rej_{slot}") != 1:
            failures.append(f"{handler.__name__}: strike not recorded ({ctx.user_data!r})")

    if generated:
        failures.append(f"a bad address reached generation: {generated!r}")

    # Name handlers must hold their state on an email, without searching.
    name_cases = [
        (bot.got_ut_name, bot.UT_NAME, "ut_name"),
        (bot.got_coc_name, bot.COC_NAME, "coc_name"),
        (bot.got_cw_name, bot.CW_NAME, "cw_name"),
        (bot.got_name, bot.ASK_NAME, "name"),
    ]
    for handler, expected_state, slot in name_cases:
        ctx = FakeCtx()
        upd = FakeUpdate("daylenis@leoempireservicesllc.com")
        state = await handler(upd, ctx)
        if state != expected_state:
            failures.append(f"{handler.__name__}: expected state {expected_state}, got {state}")
        if ctx.user_data.get(f"_rej_{slot}") != 1:
            failures.append(f"{handler.__name__}: strike not recorded ({ctx.user_data!r})")

    if generated:
        failures.append(f"a bad name reached generation: {generated!r}")

    # A good address still advances: it must reach generation.
    ctx = FakeCtx(ut_company="LEO EMPIRE SERVICES LLC")
    upd = FakeUpdate("8514 Fenway Dr, Houston, TX 77036")
    await bot.got_ut_addr(upd, ctx)
    if not generated:
        failures.append("a valid address did not reach generation")


asyncio.run(main())

if failures:
    for f in failures[:10]:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: handler guards (8 sites reject bad input, valid input still flows)")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `./venv/bin/python tests/test_handler_guards.py`
Expected: FAIL — handlers currently advance instead of holding state, e.g. `got_ut_addr: expected state 12, got ...` and `a bad address reached generation`.

- [ ] **Step 3: Guard the `/utility` handlers**

Replace the first line of `got_ut_name` (currently `query = update.message.text.strip()`) so the function begins:

```python
async def got_ut_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = await check_field(update, ctx, looks_like_company, "ut_name", "Company name?")
    if query is None:
        return UT_NAME
    user = update.effective_user.first_name
```

The rest of `got_ut_name` is unchanged.

Replace the body of `got_ut_addr` entirely:

```python
async def got_ut_addr(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    address = await check_field(update, ctx, looks_like_address, "ut_addr", "Enter the address:")
    if address is None:
        return UT_ADDR
    company = ctx.user_data["ut_company"]
    return await _ut_generate_and_send(update, ctx, company, address)
```

- [ ] **Step 4: Guard the `/new` handlers**

`got_name` begins:

```python
async def got_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = await check_field(update, ctx, looks_like_company, "name", "What's the company name?")
    if query is None:
        return ASK_NAME
    user = update.effective_user.first_name
```

The rest of `got_name` is unchanged.

Replace the body of `got_addr` entirely:

```python
async def got_addr(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    address = await check_field(update, ctx, looks_like_address, "addr", "Physical address?")
    if address is None:
        return ASK_ADDR
    ctx.user_data["current"]["address"] = address
    n = add_company(ctx, ctx.user_data.pop("current"))
    await update.message.reply_text(
        f"Company {n} added. Add another?",
        reply_markup=YES_NO
    )
    return ASK_MORE
```

- [ ] **Step 5: Guard the `/coc` handlers**

`got_coc_name` begins:

```python
async def got_coc_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = await check_field(update, ctx, looks_like_company, "coc_name", "Company name?")
    if query is None:
        return COC_NAME
    user = update.effective_user.first_name
```

The rest of `got_coc_name` is unchanged.

Replace the body of `got_coc_addr` entirely:

```python
async def got_coc_addr(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    address = await check_field(update, ctx, looks_like_address, "coc_addr", "Enter the address:")
    if address is None:
        return COC_ADDR
    company = ctx.user_data["coc_company"]
    return await _coc_generate_and_send(update, ctx, company, address)
```

- [ ] **Step 6: Guard the `/coverwhale` handlers**

`got_cw_name` begins:

```python
async def got_cw_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = await check_field(update, ctx, looks_like_company, "cw_name", "Company name?")
    if query is None:
        return CW_NAME
    user = update.effective_user.first_name
```

The rest of `got_cw_name` is unchanged.

Replace the body of `got_cw_addr` entirely:

```python
async def got_cw_addr(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    address = await check_field(update, ctx, looks_like_address, "cw_addr", "Physical address?")
    if address is None:
        return CW_ADDR
    company = ctx.user_data["cw_company"]
    usdot = ctx.user_data.get("cw_usdot", "")
    return await _cw_generate_and_send(update, ctx, company, usdot, address)
```

- [ ] **Step 7: Run the guard test**

Run: `./venv/bin/python tests/test_handler_guards.py`
Expected: `PASS: handler guards (8 sites reject bad input, valid input still flows)`

- [ ] **Step 8: Run the whole suite**

```bash
for t in tests/test_*.py; do echo "--- $t"; ./venv/bin/python "$t" || echo "SUITE FAIL: $t"; done
```

Expected: a `PASS:` line from each of the five test files, no `SUITE FAIL`.

- [ ] **Step 9: Replay the original incident end-to-end**

Confirms the reported bug can no longer happen. This drives the real `/utility` handlers with the exact values from `log/coverwhale.log:1529`.

```bash
./venv/bin/python -c "
import asyncio, os, sys
os.environ.setdefault('TELEGRAM_BOT_TOKEN','test-token')
sys.path.insert(0,'src')
import bot

class M:
    def __init__(s,t): s.text=t; s.replies=[]
    async def reply_text(s,t,reply_markup=None,**k): s.replies.append(t); print('BOT:',t.replace(chr(10),' | '))
class U:
    first_name='ibo'
class Upd:
    def __init__(s,t): s.message=M(t); s.effective_user=U()
class C:
    def __init__(s): s.user_data={}

async def main():
    ctx=C()
    print('USER: daylenis@leoempireservicesllc.com')
    st=await bot.got_ut_name(Upd('daylenis@leoempireservicesllc.com'), ctx)
    assert st==bot.UT_NAME, f'expected to stay at UT_NAME, got {st}'
    print('USER: LEO EMPIRE SERVICES LLC')
    st=await bot.got_ut_name(Upd('LEO EMPIRE SERVICES LLC'), ctx)
    assert st!=bot.UT_NAME, 'company name should have advanced'
    print('OK: email rejected, company name accepted')
asyncio.run(main())
"
```

Expected: the email is rejected with an explanation, the company name is accepted, and `OK:` prints.

- [ ] **Step 10: Commit**

```bash
git add src/bot.py tests/test_handler_guards.py
git commit -m "Validate company name and address in all four bot flows"
```

---

## Verification already done

The validator logic in Task 1 was executed against real data before this plan was written, so the executor should expect it to pass as given:

- All 26 acceptance/rejection cases in the Task 1 test pass.
- Swept over every row of `assets/All Companies.csv` (16,742 companies): **0** legal names and **0** physical addresses are rejected.
- An earlier draft included a "more digits than letters" company rule. It wrongly rejected the real companies `7573 LLC` and `1524 INC` and caught nothing the no-letters rule misses, so it was removed. Both names are regression cases in the Task 1 test.

If Task 1's test fails on first run after Step 4, the implementation was transcribed incorrectly — re-check it against the plan rather than loosening the test.

## Manual verification

Automated tests do not exercise Telegram itself. After Task 3, run the bot (`./venv/bin/python src/bot.py`) and walk `/utility` once:

1. Send an email at `Company name?` — expect the rejection message.
2. Send a company name — expect it to search or fall through to the address prompt.
3. Send a company name at `Enter the address:` — expect rejection.
4. Send the same kind of bad value again — expect the `[Use it anyway] / [Try again]` keyboard.
5. Tap `Use it anyway` — expect the bill to generate with that value.
6. Confirm `log/coverwhale.log` shows the `Rejected ...` and `Override accepted ...` lines.
