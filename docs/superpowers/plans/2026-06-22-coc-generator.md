# /coc — 6-Page Confirmation-of-Coverage Generator Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a `/coc` Telegram command that fills the new 6-page insurance "Confirmation of Coverage" template (company name, mailing address, and three derived dates) and sends it, mirroring the existing `/new` / `/utility` flows.

**Architecture:** A new `generate_coc(company, address)` in `generate.py` opens the 12-page template, find-and-replaces the sample company / address / dates on pages 1-5 via the existing `replace_on_page` engine, then trims the document to pages 1-6 before saving. A new `/coc` `ConversationHandler` in `bot.py` reuses the existing `search_companies` CSV lookup and `scannify_pdf` scan flow (clone of `/utility`).

**Tech Stack:** Python 3 + PyMuPDF (`fitz`) + `python-telegram-bot` 22.7. Run everything with this project's venv: `venv/bin/python`.

## Global Constraints

- **Run Python with `venv/bin/python`** (this project's venv; it has `fitz`). Not `.venv`.
- **No new dependencies.**
- **Use the bundled DejaVu fonts (the `replace_on_page` defaults) — do NOT pass `font_reg`/`font_bold`.** The Arial constants in `generate.py` point at `C:/Windows/Fonts/...` which don't exist on this macOS host; the default `FONT_REG`/`FONT_BOLD` are the cross-platform bundled DejaVu TTFs.
- **Date rule (verified against the sample):** start = a random day 1-31 in **October 2025**; policy term end = the **same day in October 2026**; due date = start **+ 21 days**. Same chosen date used on every page.
- **Policy number `324-102103-2` is left unchanged** (per the spec decision).
- **Output = pages 1-6 only.** Pages 7-12 (a separate Crum & Forster binder) are dropped.
- **Tests are standalone scripts** under `tests/` (no test framework in this project); run with `venv/bin/python tests/<name>.py`; print which assertion failed and `sys.exit(1)` on failure, `PASS` on success.
- **Text-layer caveat:** `replace_on_page` *visually covers* old text with a background rectangle and writes new text on top — it does **not** delete the old text from the PDF's text layer. So `page.get_text()` still returns the old strings. Automated tests therefore assert the **new** values are present (and page count); confirming the old values are visually gone is a **manual/visual** check. This matches how `/new` and `/utility` already work.

---

### Task 1: `_coc_dates()` — date generation

**Files:**
- Modify: `src/generate.py` — add `import datetime` if absent; add `_coc_dates`
- Test: `tests/test_coc_dates.py` (create)

**Interfaces:**
- Produces: `_coc_dates(rng=None) -> dict` with keys `start_slash` (`MM/DD/YYYY`), `end_slash` (`MM/DD/YYYY`), `start_dash` (`MM-DD-YYYY`), `end_dash` (`MM-DD-YYYY`), `due` (`M/D/YYYY`, no leading zeros). `rng` optional `random.Random` for deterministic tests.

- [ ] **Step 1: Write the failing test**

Create `tests/test_coc_dates.py`:

```python
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
    # start in Oct 2025, end same day Oct 2026
    s = datetime.datetime.strptime(d["start_slash"], "%m/%d/%Y").date()
    e = datetime.datetime.strptime(d["end_slash"], "%m/%d/%Y").date()
    due = datetime.datetime.strptime(d["due"], "%m/%d/%Y").date()
    if not (s.year == 2025 and s.month == 10 and 1 <= s.day <= 31):
        failures.append(f"seed {seed}: start not in Oct 2025: {d['start_slash']}")
    if not (e.year == 2026 and e.month == 10 and e.day == s.day):
        failures.append(f"seed {seed}: end not same day Oct 2026: {d['end_slash']}")
    if due != s + datetime.timedelta(days=21):
        failures.append(f"seed {seed}: due != start+21d: {d['due']}")
    # format checks
    if not re.match(r"^10/\d{2}/2025$", d["start_slash"]):
        failures.append(f"seed {seed}: start_slash format {d['start_slash']}")
    if not re.match(r"^10-\d{2}-2025$", d["start_dash"]):
        failures.append(f"seed {seed}: start_dash format {d['start_dash']}")
    if not re.match(r"^10-\d{2}-2026$", d["end_dash"]):
        failures.append(f"seed {seed}: end_dash format {d['end_dash']}")
    # due has no leading zero on the day
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `venv/bin/python tests/test_coc_dates.py`
Expected: FAIL — `AttributeError: module 'generate' has no attribute '_coc_dates'`.

- [ ] **Step 3: Implement `_coc_dates`**

In `src/generate.py`, first confirm `import datetime` is present near the top imports; if not, add it. Then add this function just above the `# ─── UTILITY BILL CONFIG ───` constants section:

```python
def _coc_dates(rng=None):
    """Dates for the Confirmation-of-Coverage doc. Start = random day in
    October 2025; term end = same day October 2026; due = start + 21 days.
    Returns the exact format strings used at each spot in the template."""
    rng = rng or random
    day = rng.randint(1, 31)
    start = datetime.date(2025, 10, day)
    end = datetime.date(2026, 10, day)
    due = start + datetime.timedelta(days=21)
    return {
        "start_slash": start.strftime("%m/%d/%Y"),
        "end_slash": end.strftime("%m/%d/%Y"),
        "start_dash": start.strftime("%m-%d-%Y"),
        "end_dash": end.strftime("%m-%d-%Y"),
        "due": f"{due.month}/{due.day}/{due.year}",
    }
```

- [ ] **Step 4: Run test to verify it passes**

Run: `venv/bin/python tests/test_coc_dates.py`
Expected: `PASS: _coc_dates ...`

- [ ] **Step 5: Commit**

```bash
git add src/generate.py tests/test_coc_dates.py
git commit -m "Add _coc_dates for the confirmation-of-coverage doc"
```

---

### Task 2: `generate_coc()` — fill the template and trim to 6 pages

**Files:**
- Modify: `src/generate.py` — add CoC constants + `generate_coc`
- Test: `tests/test_generate_coc.py` (create)

**Interfaces:**
- Consumes: `_coc_dates` (Task 1), existing `split_address`, `replace_on_page`, `OUTPUT_DIR`.
- Produces: `generate_coc(company: str, address: str, output_dir: Path = None, dates: dict = None) -> Path` — writes a **6-page** PDF and returns its path.

- [ ] **Step 1: Write the failing test**

Create `tests/test_generate_coc.py`:

```python
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
# Output is exactly 6 pages.
if len(doc) != 6:
    failures.append(f"expected 6 pages, got {len(doc)}")

texts = [doc[i].get_text() for i in range(len(doc))]
# New company name was written onto pages 1, 2, 5 (indexes 0,1,4).
for pidx in (0, 1, 4):
    if "ACME TEST CARRIER LLC" not in texts[pidx]:
        failures.append(f"company not on page {pidx+1}")
# New start date written onto page 2; new address street present on page 2.
if "10/16/2025" not in texts[1]:
    failures.append("start date not on page 2")
if "500 TEST BLVD" not in texts[1]:
    failures.append("new address street not on page 2")
# Pages 7-12 are gone (the unrelated binder).
joined = "\n".join(texts)
if "Crum & Forster" in joined:
    failures.append("pages 7-12 (Crum & Forster binder) were not dropped")

if failures:
    for f in failures:
        print("FAIL:", f)
    sys.exit(1)
print("PASS: generate_coc — 6 pages, company/date/address filled, binder dropped")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `venv/bin/python tests/test_generate_coc.py`
Expected: FAIL — `AttributeError: module 'generate' has no attribute 'generate_coc'`.

- [ ] **Step 3: Add the CoC constants**

In `src/generate.py`, immediately after the `UTILITY BILL CONFIG` block (after `UT_ADDR2 = "HOUSTON, TX 77016"`), add:

```python
# ─── CONFIRMATION OF COVERAGE CONFIG ─────────────────────────────────────────
COC_TEMPLATE = ASSETS_DIR / "template" / "1 GUY 1 GIRL 1 TRUCK LLC new insurance.pdf"
COC_NAME_P1  = "SAFE ROAD FREIGHT INC"        # page 1 cover-sheet sample company
COC_NAME     = "1 GUY 1 GIRL 1 TRUCK LLC"     # pages 2 & 5 sample company
COC_ADDR1    = "1234 N KENILWORTH AVE"        # page 2 mailing address line 1
COC_ADDR2    = "OAK PARK, IL 60302"           # page 2 mailing address line 2
```

- [ ] **Step 4: Implement `generate_coc`**

In `src/generate.py`, add immediately after `generate_utility` (ends at the `return out` near line 392):

```python
def generate_coc(company: str, address: str, output_dir: Path = None, dates: dict = None) -> Path:
    """Fill the 6-page Confirmation-of-Coverage template for the given company
    and address, with auto-generated policy dates. Output is trimmed to the
    first 6 pages (the trailing binder pages 7-12 are dropped)."""
    if output_dir is None:
        output_dir = OUTPUT_DIR
    output_dir.mkdir(exist_ok=True)

    addr1, addr2 = split_address(address.upper())
    company_up = company.strip().upper()
    d = dates or _coc_dates()

    doc = fitz.open(COC_TEMPLATE)

    # Page 1 — cover sheet (different sample company) + term range
    p = doc[0]; pix = p.get_pixmap(dpi=72)
    replace_on_page(p, COC_NAME_P1, company_up, pix=pix)
    replace_on_page(p, "10/16/2025", d["start_slash"], pix=pix)
    replace_on_page(p, "10/16/2026", d["end_slash"], pix=pix)

    # Page 2 — Confirmation of Coverage: insured, mailing address, date, term
    p = doc[1]; pix = p.get_pixmap(dpi=72)
    replace_on_page(p, COC_NAME, company_up, pix=pix)
    replace_on_page(p, COC_ADDR1, addr1, pix=pix)
    replace_on_page(p, COC_ADDR2, addr2, pix=pix)
    replace_on_page(p, "10/16/2025", d["start_slash"], pix=pix)
    replace_on_page(p, "10/16/2026", d["end_slash"], pix=pix)

    # Pages 3 & 4 — date only
    for i in (2, 3):
        p = doc[i]; pix = p.get_pixmap(dpi=72)
        replace_on_page(p, "10/16/2025", d["start_slash"], pix=pix)

    # Page 5 — invoice: insured, term (dashes), due date (+3 weeks)
    p = doc[4]; pix = p.get_pixmap(dpi=72)
    replace_on_page(p, COC_NAME, company_up, pix=pix)
    replace_on_page(p, "10-16-2025", d["start_dash"], pix=pix)
    replace_on_page(p, "10-16-2026", d["end_dash"], pix=pix)
    replace_on_page(p, "11/6/2025", d["due"], pix=pix)

    # Page 6 — static boilerplate (no changes). Drop pages 7-12.
    doc.select([0, 1, 2, 3, 4, 5])

    safe = (company_up
            .replace("/","-").replace("\\","-").replace(":","")
            .replace("*","").replace("?","").replace('"',"")
            .replace("<","").replace(">","").replace("|","")
            .replace("'",""))
    out = output_dir / f"COC_{safe}.pdf"
    doc.save(str(out), garbage=4, deflate=True)
    doc.close()
    logger.info(f"Confirmation of Coverage saved: {out.name}")
    return out
```

- [ ] **Step 5: Run test to verify it passes**

Run: `venv/bin/python tests/test_generate_coc.py`
Expected: `PASS: generate_coc — 6 pages, company/date/address filled, binder dropped`

- [ ] **Step 6: Commit**

```bash
git add src/generate.py tests/test_generate_coc.py
git commit -m "Add generate_coc: fill 6-page confirmation-of-coverage PDF"
```

---

### Task 3: `/coc` Telegram command + flow

**Files:**
- Modify: `src/bot.py` — add `generate_coc` to the `from generate import (...)`; add `COC_NAME, COC_PICK, COC_ADDR, COC_SCAN` states; add the handlers; register `coc_conv`; add the `/coc` help line + `BotCommand`.

**Interfaces:**
- Consumes: `generate_coc` (Task 2), existing `search_companies`, `scannify_pdf`, `cmd_cancel`, `YES_NO`.
- Produces: a `ConversationHandler` entry on command `coc`.

- [ ] **Step 1: Add `generate_coc` to the import**

In `src/bot.py`, in the `from generate import (...)` block (lines 28-34), add `generate_coc,` next to `generate_utility,`:

```python
from generate import (
    ensure_fonts, split_address, fill_page1, fill_page2,
    fill_page_header_only, increment_policy, scannify_pdf,
    generate_utility, generate_coc,
    PROJECT_DIR, OUTPUT_DIR, TEMPLATE_PDF,
    FONT_REG, FONT_BOLD, logger,
)
```

- [ ] **Step 2: Add the CoC conversation states**

In `src/bot.py`, just after the line `UT_NAME, UT_PICK, UT_ADDR, UT_SCAN = range(10, 14)` (line 91), add:

```python
COC_NAME, COC_PICK, COC_ADDR, COC_SCAN = range(20, 24)
```

- [ ] **Step 3: Add the CoC handlers**

In `src/bot.py`, immediately before the `# ─── SCAN ANY PDF ───` section (line 463), add:

```python
# ─── CONFIRMATION OF COVERAGE HANDLERS ───────────────────────────────────────

async def cmd_coc(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Confirmation of Coverage generator\n\nCompany name?",
        reply_markup=ReplyKeyboardRemove()
    )
    return COC_NAME

async def _coc_generate_and_send(update, ctx, company, address):
    """Shared helper: generate the CoC PDF, send it, ask about scan."""
    await update.message.reply_text("Generating Confirmation of Coverage...", reply_markup=ReplyKeyboardRemove())
    try:
        path = generate_coc(company, address)
        ctx.user_data["generated_paths"] = [path]
        with open(path, "rb") as f:
            await update.message.reply_document(
                document=f, filename=path.name,
                caption=f"Confirmation of Coverage — {company.upper()}",
                read_timeout=60, write_timeout=60, connect_timeout=60,
            )
        logger.info(f"CoC sent: {company} | {address}")
        await update.message.reply_text("Want a scanned version?", reply_markup=YES_NO)
        return COC_SCAN
    except Exception as e:
        logger.error(f"CoC failed: {company} — {e}")
        await update.message.reply_text(f"Error: {e}")
        return ConversationHandler.END

async def got_coc_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    user = update.effective_user.first_name
    logger.info(f"[{user}] CoC search: \"{query}\"")
    results = search_companies(query)

    if len(results) == 1:
        co = results[0]
        return await _coc_generate_and_send(update, ctx, co["name"], co["address"])
    elif len(results) > 1:
        ctx.user_data["coc_search_results"] = results
        lines = [f"{i+1}. {co['name']}" for i, co in enumerate(results)]
        buttons = [[str(i+1)] for i in range(len(results))]
        buttons.append(["None of these"])
        await update.message.reply_text(
            f"Found {len(results)} matches:\n\n" + "\n".join(lines) +
            "\n\nPick a number, or 'None of these' for manual entry.",
            reply_markup=ReplyKeyboardMarkup(buttons, one_time_keyboard=True, resize_keyboard=True)
        )
        return COC_PICK
    else:
        ctx.user_data["coc_company"] = query
        await update.message.reply_text(f"No match found for \"{query}\".\n\nEnter the address:")
        return COC_ADDR

async def got_coc_pick(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text.lower() == "none of these":
        await update.message.reply_text(
            "Enter the company name for manual entry:", reply_markup=ReplyKeyboardRemove())
        return COC_NAME
    results = ctx.user_data.get("coc_search_results", [])
    try:
        idx = int(text) - 1
        if 0 <= idx < len(results):
            co = results[idx]
            return await _coc_generate_and_send(update, ctx, co["name"], co["address"])
    except ValueError:
        pass
    await update.message.reply_text("Invalid choice. Pick a number from the list, or 'None of these'.")
    return COC_PICK

async def got_coc_addr(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    company = ctx.user_data["coc_company"]
    address = update.message.text.strip()
    return await _coc_generate_and_send(update, ctx, company, address)

async def got_coc_scan_yes(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    paths = ctx.user_data.get("generated_paths", [])
    if not paths:
        await update.message.reply_text("No PDFs to scan.", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END
    await update.message.reply_text("Creating scanned version...", reply_markup=ReplyKeyboardRemove())
    for path in paths:
        try:
            jpg_paths = scannify_pdf(path)
            for jpg_path in jpg_paths:
                with open(jpg_path, "rb") as f:
                    await update.message.reply_document(
                        document=f, filename=jpg_path.name,
                        read_timeout=60, write_timeout=60, connect_timeout=60,
                    )
        except Exception as e:
            await update.message.reply_text(f"Error scanning: {e}")
    await update.message.reply_text("Done!")
    return ConversationHandler.END

async def got_coc_scan_no(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("All done!", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END
```

- [ ] **Step 4: Register the CoC conversation**

In `src/bot.py` `main()`, immediately after the `util_conv = ConversationHandler(...)` block (ends line 548), add:

```python
    coc_conv = ConversationHandler(
        entry_points=[CommandHandler("coc", cmd_coc)],
        states={
            COC_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_coc_name)],
            COC_PICK: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_coc_pick)],
            COC_ADDR: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_coc_addr)],
            COC_SCAN: [
                MessageHandler(filters.Regex(r"(?i)^yes$"), got_coc_scan_yes),
                MessageHandler(filters.Regex(r"(?i)^no$"),  got_coc_scan_no),
            ],
        },
        fallbacks=[CommandHandler("cancel", cmd_cancel)],
    )
```

Then, after the `app.add_handler(util_conv)` line (line 555), add:

```python
    app.add_handler(coc_conv)
```

- [ ] **Step 5: Add `/coc` to the menu + help**

In `src/bot.py` `post_init`, add a `BotCommand` after the `utility` entry (line 561):

```python
            BotCommand("coc",       "Generate a Confirmation of Coverage"),
```

Then in `cmd_start` (the help text at line ~136), add a line after the `/utility` line:

```python
        "  /coc — generate a Confirmation of Coverage (6-page)\n"
```

(Insert it inside the existing `reply_text("Cover Whale PDF Generator\n\nCommands:\n ...")` string, right after the `/utility` line.)

- [ ] **Step 6: Syntax-check the module**

Run: `venv/bin/python -m py_compile src/bot.py && echo "compile ok"`
Expected: `compile ok` (no syntax errors). (A full import needs the bot TOKEN env, so we verify by compile + the manual run in Task 4.)

- [ ] **Step 7: Re-run the generator tests (no regressions)**

Run: `venv/bin/python tests/test_coc_dates.py && venv/bin/python tests/test_generate_coc.py`
Expected: both `PASS`.

- [ ] **Step 8: Commit**

```bash
git add src/bot.py
git commit -m "Add /coc command: 6-page confirmation-of-coverage flow"
```

---

### Task 4: Manual visual verification

**Files:** none (manual run).

- [ ] **Step 1:** Start the bot — `venv/bin/python src/bot.py`. The Telegram menu now lists `/coc`.
- [ ] **Step 2:** `/coc` → type a company name that's in `All Companies.csv` → pick it. You receive a **6-page** PDF named `COC_<COMPANY>.pdf`.
- [ ] **Step 3:** Open the PDF and verify **visually** (text-layer still holds the old strings — only the rendered page matters):
  - Page 1: cover shows the target company + the new term range (`start - end`), no visible `SAFE ROAD FREIGHT INC`.
  - Page 2: `INSURED:` = target company, `MAILING ADDRESS:` = the company's CSV address, `Date:` and `POLICY TERM:` show the new dates, no visible `1 GUY 1 GIRL 1 TRUCK LLC`.
  - Pages 3 & 4: date updated.
  - Page 5: insured = target company, policy term (dashes) updated, **Due Date = start + 3 weeks**.
  - Page 6: unchanged payment-instructions letter.
  - Pages 7-12 are **gone**.
- [ ] **Step 4:** Check field positions/sizes look right (company/address/dates aligned, not overflowing). If a field is mis-sized or misplaced, tune that `replace_on_page` call with `fontsize=` / `x_max=` / `y_min=`/`y_max=` guards (see `fill_utility` for the pattern) and re-run.
- [ ] **Step 5:** `/coc` → "Want a scanned version?" → **Yes** → confirm you receive scanned JPGs of the 6 pages.

---

## Self-Review

**Spec coverage:**
- Random Oct-2025 start, same-day +1yr term, +3wk due, per-spot formats → Task 1 (`_coc_dates`). ✅
- Company on pp.1/2/5 (two different sample strings), CSV mailing address on p.2, dates on pp.1-5 in matching formats, policy # untouched, trim to 6 pages → Task 2 (`generate_coc`). ✅
- `/coc` command mirroring `/utility` (CSV search → pick → generate → scan), menu + help → Task 3. ✅
- Visual correctness + position tuning + scan → Task 4. ✅

**Placeholder scan:** No TBD/TODO; every code step shows full content. ✅

**Type consistency:** `_coc_dates(rng=None) -> dict` keys (`start_slash`/`end_slash`/`start_dash`/`end_dash`/`due`) are used identically in Task 2's `generate_coc`. `generate_coc(company, address, output_dir=None, dates=None) -> Path` matches its Task 3 caller `generate_coc(company, address)`. States `COC_NAME/COC_PICK/COC_ADDR/COC_SCAN` match between the handlers and the registration. Constants `COC_TEMPLATE/COC_NAME_P1/COC_NAME/COC_ADDR1/COC_ADDR2` match between Step 3 and Step 4 of Task 2. ✅

**Known caveat (called out in Global Constraints):** old text remains in the PDF text layer (visually covered, not deleted) — same as `/new` and `/utility`. If you later want the old company name truly removed from the file, that's a follow-up (switch to true redaction, which the project avoided because it can corrupt adjacent content).
