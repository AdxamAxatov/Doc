#!/usr/bin/env python3
"""
Cover Whale PDF Generator  v3
─────────────────────────────
Reads company data from Excel → fills Cover Whale insurance policy template
→ outputs one ready PDF per company.

Run:  py generate.py
"""

import fitz
import openpyxl
import os, sys, urllib.request, zipfile, io, logging, random, calendar
from datetime import datetime, date, timedelta
from pathlib import Path
from PIL import Image, ImageFilter, ImageEnhance
import numpy as np

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ─── LOGGING ─────────────────────────────────────────────────────────────────

LOG_DIR = Path(__file__).parent.parent / "log"
LOG_DIR.mkdir(exist_ok=True)

logger = logging.getLogger("coverwhale")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    fh = logging.FileHandler(LOG_DIR / "coverwhale.log", encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
    logger.addHandler(fh)

# ─── CONFIG ──────────────────────────────────────────────────────────────────

SCRIPT_DIR      = Path(__file__).parent
PROJECT_DIR     = SCRIPT_DIR.parent
ASSETS_DIR      = PROJECT_DIR / "assets"
TEMPLATE_PDF    = ASSETS_DIR / "template" / "Cover Whale -   VIATIC LLC.pdf"
EXCEL_FILE      = ASSETS_DIR / "companies.xlsx"
OUTPUT_DIR      = PROJECT_DIR / "output"
STARTING_POLICY = "CUS09116674"

FONT_REG  = ASSETS_DIR / "DejaVuSansCondensed.ttf"
FONT_BOLD = ASSETS_DIR / "DejaVuSansCondensed-Bold.ttf"
FONT_ZIP  = "https://sourceforge.net/projects/dejavu/files/dejavu/2.37/dejavu-fonts-ttf-2.37.zip/download"

# Arial fonts — macOS system path, falling back to the Windows system path.
_MAC_ARIAL   = Path("/System/Library/Fonts/Supplemental/Arial.ttf")
_MAC_ARIAL_B = Path("/System/Library/Fonts/Supplemental/Arial Bold.ttf")
ARIAL_REG  = _MAC_ARIAL   if _MAC_ARIAL.exists()   else Path("C:/Windows/Fonts/arial.ttf")
ARIAL_BOLD = _MAC_ARIAL_B if _MAC_ARIAL_B.exists() else Path("C:/Windows/Fonts/arialbd.ttf")

# Confirmation-of-Coverage policy term. The start day is randomised within this
# month and the term runs to the same day one year later. Shifting the term means
# changing these two values and nothing else — tests/test_coc_dates.py reads them
# rather than repeating the month and year.
COC_TERM_YEAR  = 2026
COC_TERM_MONTH = 4                      # April


def _coc_dates(rng=None):
    """Dates for the Confirmation-of-Coverage doc. Start = random day in
    COC_TERM_MONTH of COC_TERM_YEAR; term end = same day one year later;
    due = start + 21 days.
    Returns the exact format strings used at each spot in the template."""
    rng = rng or random
    day = rng.randint(1, calendar.monthrange(COC_TERM_YEAR, COC_TERM_MONTH)[1])
    # A Feb 29 start has no counterpart in a non-leap year; fall back to the
    # last day of the month, as policy terms conventionally do.
    end_day = min(day, calendar.monthrange(COC_TERM_YEAR + 1, COC_TERM_MONTH)[1])
    start = date(COC_TERM_YEAR,     COC_TERM_MONTH, day)
    end   = date(COC_TERM_YEAR + 1, COC_TERM_MONTH, end_day)
    due = start + timedelta(days=21)
    return {
        "start_slash": start.strftime("%m/%d/%Y"),
        "end_slash": end.strftime("%m/%d/%Y"),
        "start_dash": start.strftime("%m-%d-%Y"),
        "end_dash": end.strftime("%m-%d-%Y"),
        "due": f"{due.month}/{due.day}/{due.year}",
    }


# ─── UTILITY BILL CONFIG ─────────────────────────────────────────────────────
UTILITY_TEMPLATE = ASSETS_DIR / "template" / "Utility_IVORY JULIUS CHRISTOPHER.pdf"
UT_NAME    = "IVORY JULIUS CHRISTOPHER"
UT_ADDR1   = "10318 CHEEVES,"
UT_ADDR2   = "HOUSTON, TX 77016"

# ─── CONFIRMATION OF COVERAGE CONFIG ─────────────────────────────────────────
COC_TEMPLATE = ASSETS_DIR / "template" / "1 GUY 1 GIRL 1 TRUCK LLC new insurance.pdf"
COC_NAME_P1  = "SAFE ROAD FREIGHT INC"        # page 1 cover-sheet sample company
COC_NAME     = "1 GUY 1 GIRL 1 TRUCK LLC"     # pages 2 & 5 sample company
COC_ADDR1    = "1234 N KENILWORTH AVE"        # page 2 mailing address line 1
COC_ADDR2    = "OAK PARK, IL 60302"           # page 2 mailing address line 2

# Exact fonts/sizes/colors read from the template's own text spans:
#   page 1  company = Arial-BoldMT 22pt #171717 ; date = Arial-Black 16pt #171717 (centered)
#   pages 2-4 values = DejaVuSans 10pt black ; page 5 values = DejaVuSans 9pt black
def _coc_font(*candidates, fallback):
    for c in candidates:
        if Path(c).exists():
            return Path(c)
    return fallback

COC_FONT_BODY  = ASSETS_DIR / "DejaVuSans.ttf"  # pages 2-5 values
COC_FONT_TITLE = _coc_font("/System/Library/Fonts/Supplemental/Arial Bold.ttf",
                           "C:/Windows/Fonts/arialbd.ttf", fallback=COC_FONT_BODY)
COC_FONT_DATE1 = _coc_font("/System/Library/Fonts/Supplemental/Arial Black.ttf",
                           "C:/Windows/Fonts/ariblk.ttf", fallback=COC_FONT_BODY)
COC_CLR_DARK  = (0x17 / 255.0, 0x17 / 255.0, 0x17 / 255.0)  # #171717
COC_CLR_BLACK = (0.0, 0.0, 0.0)


def _coc_set(page, old, new, *, size, color, font, center=False):
    """Cover the template's sample text with white and write the new value at
    the same baseline using an exact font/size/color — no heuristics, no
    background sampling (these fields all sit on white)."""
    fp = str(font)
    fobj = fitz.Font(fontfile=fp)
    fname = "F" + str(abs(hash(fp)) % 100000)
    for rect in page.search_for(old):
        page.draw_rect(fitz.Rect(rect.x0 - 1, rect.y0 - 1, rect.x1 + 1, rect.y1 + 1),
                       color=(1, 1, 1), fill=(1, 1, 1), width=0)
        tw = fobj.text_length(new, fontsize=size)
        x = (page.rect.width - tw) / 2 if center else rect.x0
        y = rect.y1 - 1.0
        page.insert_text((x, y), new, fontfile=fp, fontname=fname, fontsize=size, color=color)

# ─── TEMPLATE VALUES (VIATIC LLC base PDF) ───────────────────────────────────
T_COMPANY = "VIATIC LLC"
T_USDOT   = "USDOT # 3846659"
T_ADDR1   = "3975 NW 176TH ST"
T_ADDR2   = "MIAMI GARDENS, FL 33055"
T_POLICY  = "CUS09114581"

# ─── STATIC TEMPLATE VALUES (Tr=3 invisible in template → must be re-rendered) ─
T_FROM      = "From:"
T_TO        = "To:"
T_TERM_FROM = "October 14, 2025"
T_TERM_TO   = "October 14, 2026"
T_BROKER1   = "Empire State Brokerage Services LLC - DAVID SCHEPSMAN"
T_BROKER2   = "(DAVID@ESBSLLC.COM )"
T_ISSUED_TM = "11:31:12 EST (Eastern Standard Time)"

# ─── PAGE 2 COLUMN CENTERS (Confirmation of Coverage table) ──────────────────
# Center of Mailing Address column (~x400 to x560) → used for centering the value
P2_MAILING_ADDR_CX  = 480.0

# ─── TOP-RIGHT FIXED RIGHT EDGE ──────────────────────────────────────────────
# Both policy # and company name right-align to the same edge across all pages
TOP_RIGHT_X = 552.0

# ─── TEXT COLORS ─────────────────────────────────────────────────────────────
CLR_HEADER  = (54/255, 54/255, 54/255)    # #363636 — top-right header text
CLR_TITLE   = (51/255, 51/255, 51/255)    # #333333 — center bold title / USDOT

# ─── FONTS ───────────────────────────────────────────────────────────────────

def ensure_fonts():
    if FONT_REG.exists() and FONT_BOLD.exists():
        return
    print("  Downloading DejaVu fonts (one-time) ...")
    with urllib.request.urlopen(FONT_ZIP) as r:
        data = r.read()
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for m in zf.namelist():
            if m.endswith("DejaVuSansCondensed.ttf") and not FONT_REG.exists():
                FONT_REG.write_bytes(zf.read(m));  print(f"  + {FONT_REG.name}")
            elif m.endswith("DejaVuSansCondensed-Bold.ttf") and not FONT_BOLD.exists():
                FONT_BOLD.write_bytes(zf.read(m)); print(f"  + {FONT_BOLD.name}")
    if not FONT_REG.exists() or not FONT_BOLD.exists():
        print("  Font extraction failed — place TTF files next to generate.py")
        sys.exit(1)

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def increment_policy(p: str) -> str:
    digits = "".join(c for c in p if c.isdigit())
    prefix = "".join(c for c in p if not c.isdigit())
    return f"{prefix}{int(digits)+1:0{len(digits)}d}"


def split_address(addr: str):
    """
    Split address into (street, city_state_zip).
    Handles Excel cells where lines are separated by \n or \n\n.
    Falls back to comma split.
    """
    addr = addr.strip()
    if "\n" in addr:
        parts = [p.strip() for p in addr.split("\n") if p.strip()]
        return (parts[0], parts[-1]) if len(parts) >= 2 else (parts[0], "")
    parts = addr.split(",", 1)
    return (parts[0].strip(), parts[1].strip()) if len(parts) == 2 else (addr, "")


def height_to_params(h: float):
    """
    (fontsize, use_bold, center_on_page) from text bbox height.
    Calibrated from the Cover Whale template PDF.
    """
    if h > 14:    return 16.50, True,  True    # Page 1 large bold company title
    elif h > 11:  return 10.62, False, False   # Page 1 box name / address
    elif h > 7.5: return  8.56, False, False   # Page 2 Named Insured value
    else:         return  7.36, False, False   # Top-right tiny header labels


def sample_bg(pix, rect, pw, ph):
    """
    Sample the PDF background color just LEFT of the text bbox.
    Returns (r, g, b) floats 0-1.  Falls back to white if anything fails.
    """
    if pix is None:
        return (1.0, 1.0, 1.0)
    sx = pix.width  / pw
    sy = pix.height / ph
    px = int((rect.x0 - 4) * sx)
    py = int(((rect.y0 + rect.y1) / 2) * sy)
    if px < 0:
        px = int((rect.x1 + 4) * sx)
    px = max(0, min(pix.width  - 1, px))
    py = max(0, min(pix.height - 1, py))
    try:
        return tuple(c / 255.0 for c in pix.pixel(px, py)[:3])
    except Exception:
        return (1.0, 1.0, 1.0)


def replace_on_page(page, old_text, new_text, pix=None,
                    fontsize=None, bold=False, center=False,
                    cell_center_x=None, cell_right_x=None, cell_left_x=None,
                    cell_bounds=None,
                    top_right_x=None,
                    x_min=None, x_max=None,
                    y_min=None, y_max=None,
                    color=None,
                    font_reg=None, font_bold=None):
    """
    Find every occurrence of old_text on page that passes the x/y filters,
    cover it with a filled rectangle matching the background, then write
    new_text at the correct position.
    Uses draw_rect instead of redaction annotations to avoid corrupting
    adjacent content in the PDF stream.
    """
    hits = page.search_for(old_text)
    if not hits:
        return

    pw, ph = page.rect.width, page.rect.height

    for rect in hits:
        # ── positional guards ────────────────────────────────────────────────
        if x_min is not None and rect.x0 < x_min: continue
        if x_max is not None and rect.x1 > x_max: continue
        if y_min is not None and rect.y0 < y_min: continue
        if y_max is not None and rect.y1 > y_max: continue

        # ── font / size ──────────────────────────────────────────────────────
        h = rect.y1 - rect.y0
        if fontsize is None:
            sz, use_bold, use_center = height_to_params(h)
        else:
            sz, use_bold, use_center = fontsize, bold, center

        # ── x alignment ──────────────────────────────────────────────────────
        f_reg  = font_reg  or FONT_REG
        f_bold = font_bold or FONT_BOLD
        fp       = str(f_bold if use_bold else f_reg)
        font_obj = fitz.Font(fontfile=fp)
        tw       = font_obj.text_length(new_text, fontsize=sz)

        if use_center:
            # Large bold title — center across full page width
            x = (pw - tw) / 2
        elif cell_bounds is not None:
            # Smart center: center within cell, left-align with padding if too wide
            cb_left, cb_right = cell_bounds
            cell_w = cb_right - cb_left
            if tw < cell_w:
                x = cb_left + (cell_w - tw) / 2
            else:
                x = cb_left + 2.0
        elif cell_right_x is not None:
            # Right-align within a table column
            x = cell_right_x - tw
        elif cell_center_x is not None:
            # Center within a table column
            x = cell_center_x - tw / 2
        elif cell_left_x is not None:
            # Fixed left edge within a table column
            x = cell_left_x
        elif rect.x0 > 400 and rect.y0 < 80:
            # Top-right corner header → right-align to fixed common edge
            trx = top_right_x if top_right_x is not None else TOP_RIGHT_X
            x = min(trx, pw - 4) - tw
        else:
            # Body text — left-align at original position
            x = rect.x0

        y = rect.y1 - 1.0          # baseline just inside the bottom of the bbox

        # ── cover old text with a rectangle matching the actual background ──
        bg = sample_bg(pix, rect, pw, ph)
        cover = fitz.Rect(rect.x0 - 1.0, rect.y0 - 1.0,
                          rect.x1 + 1.0, rect.y1 + 1.0)
        page.draw_rect(cover, color=bg, fill=bg, width=0)

        # ── write new text on top ─────────────────────────────────────────────
        if font_reg is not None:
            fnm = "ArialBd" if use_bold else "Arial"
        else:
            fnm = "DejaVuSCBd" if use_bold else "DejaVuSC"
        text_color = color if color is not None else (0, 0, 0)
        page.insert_text((x, y), new_text, fontfile=fp, fontname=fnm, fontsize=sz, color=text_color)

# ─── PAGE FILL FUNCTIONS ─────────────────────────────────────────────────────

def fill_page1(page, company, usdot, addr1, addr2, policy, pix):
    """Page 1 — cover page."""

    # Top-right:  TGL Policy #:  CUS...  — right edge aligned to truck picture (x≈569)
    replace_on_page(page,
                    f"TGL Policy #:  {T_POLICY}",
                    f"TGL Policy #:  {policy}",
                    fontsize=7.36, top_right_x=569.0, pix=pix, color=CLR_HEADER)

    # Top-right company name
    replace_on_page(page, T_COMPANY, company,
                    fontsize=7.36, top_right_x=569.0, pix=pix, x_min=400, y_max=80, color=CLR_HEADER)

    # Centre bold company name (large title)
    replace_on_page(page, T_COMPANY, company,
                    fontsize=16.50, bold=True, center=True, pix=pix, x_max=400, y_min=80, y_max=400, color=CLR_TITLE)

    # Box name (below title)
    replace_on_page(page, T_COMPANY, company,
                    fontsize=10.62, pix=pix, x_max=400, y_min=400, y_max=600)

    # USDOT line (bold, centred)
    replace_on_page(page, T_USDOT, f"USDOT # {usdot}",
                    fontsize=11.0, bold=True, center=True, pix=pix, color=CLR_TITLE)

    # Address box — restrict to box area only (y 560–640)
    replace_on_page(page, T_ADDR1, addr1,
                    fontsize=10.62, pix=pix, y_min=560, y_max=610)
    replace_on_page(page, T_ADDR2, addr2,
                    fontsize=10.62, pix=pix, y_min=600, y_max=640)

    # Re-render static values (template stores them as Tr=3 invisible text)
    replace_on_page(page, T_FROM,      T_FROM,      pix=pix, fontsize=10.62, y_min=625, y_max=650)
    replace_on_page(page, T_TERM_FROM, T_TERM_FROM, pix=pix, fontsize=10.62, y_min=625, y_max=650)
    replace_on_page(page, T_TO,        T_TO,        pix=pix, fontsize=10.62, y_min=625, y_max=650)
    replace_on_page(page, T_TERM_TO,   T_TERM_TO,   pix=pix, fontsize=10.62, y_min=625, y_max=650)
    replace_on_page(page, T_BROKER1,   T_BROKER1,   pix=pix, fontsize=10.62, y_min=645, y_max=670)
    replace_on_page(page, T_BROKER2,   T_BROKER2,   pix=pix, fontsize=10.62, y_min=665, y_max=690)
    replace_on_page(page, T_TERM_FROM, T_TERM_FROM, pix=pix, fontsize=10.62, y_min=685, y_max=710)
    replace_on_page(page, T_ISSUED_TM, T_ISSUED_TM, pix=pix, fontsize=10.62, y_min=685, y_max=710)


def fill_page2(page, company, addr1, addr2, policy, pix):
    """Page 2 — Confirmation of Coverage."""

    # Top-right header
    replace_on_page(page,
                    f"TGL Policy #:  {T_POLICY}",
                    f"TGL Policy #:  {policy}",
                    fontsize=7.36, top_right_x=569.0, pix=pix, color=CLR_HEADER)

    # Top-right company name (right-aligned, top-right corner only)
    replace_on_page(page, T_COMPANY, company,
                    fontsize=7.36, top_right_x=569.0, pix=pix, x_min=400, y_max=80, color=CLR_HEADER)

    # Named Insured — smart centered within cell (42.8 – 217.9)
    replace_on_page(page, T_COMPANY, company,
                    fontsize=8.56,
                    cell_bounds=(42.8, 217.9),
                    pix=pix, x_max=300, y_min=200, y_max=240)

    # Policy Number — smart centered within cell (305.8 – 437.1)
    replace_on_page(page, T_POLICY, policy,
                    fontsize=9.63,
                    cell_bounds=(305.8, 437.1),
                    pix=pix, x_min=300, y_min=155, y_max=195)

    # Mailing Address — smart centered within cell (393.7 – 568.8)
    replace_on_page(page, T_ADDR1, addr1,
                    fontsize=8.56,
                    cell_bounds=(393.7, 568.8),
                    pix=pix, y_min=200, y_max=225)
    replace_on_page(page, T_ADDR2, addr2,
                    fontsize=8.56,
                    cell_bounds=(393.7, 568.8),
                    pix=pix, y_min=215, y_max=245)


def fill_page_header_only(page, company, policy, pix):
    """Pages 3+ — only top-right corner needs updating."""
    replace_on_page(page,
                    f"TGL Policy #:  {T_POLICY}",
                    f"TGL Policy #:  {policy}",
                    fontsize=7.36, top_right_x=569.0, pix=pix, color=CLR_HEADER)
    replace_on_page(page, T_COMPANY, company,
                    fontsize=7.36, top_right_x=569.0, pix=pix, x_min=400, y_max=80, color=CLR_HEADER)

# ─── UTILITY BILL FILL ────────────────────────────────────────────────────────

def fill_utility(page, company, addr1, addr2, pix):
    """
    Replace company name and address on page 1 of the Comcast utility template.
    Two locations: upper-left header area and bottom payment slip.
    """
    ar = dict(font_reg=ARIAL_REG, font_bold=ARIAL_BOLD)

    # Upper-left: bold company name (size 12)
    replace_on_page(page, UT_NAME, company,
                    fontsize=12.0, bold=True, pix=pix,
                    y_min=110, y_max=145, x_max=300, **ar)

    # Upper-left: address lines (size 9)
    replace_on_page(page, UT_ADDR1, addr1,
                    fontsize=9.0, pix=pix,
                    y_min=155, y_max=180, x_max=200, **ar)
    replace_on_page(page, UT_ADDR2, addr2,
                    fontsize=9.0, pix=pix,
                    y_min=170, y_max=195, x_max=200, **ar)

    # Bottom payment slip: company name (size 9, regular)
    replace_on_page(page, UT_NAME, company,
                    fontsize=9.0, pix=pix,
                    y_min=625, y_max=650, x_max=200, **ar)

    # Bottom payment slip: address lines (size 9)
    replace_on_page(page, UT_ADDR1, addr1,
                    fontsize=9.0, pix=pix,
                    y_min=640, y_max=660, x_max=200, **ar)
    replace_on_page(page, UT_ADDR2, addr2,
                    fontsize=9.0, pix=pix,
                    y_min=650, y_max=670, x_max=200, **ar)


def generate_utility(company: str, address: str, output_dir: Path = None) -> Path:
    """Generate a Comcast utility bill PDF with the given company name and address."""
    if output_dir is None:
        output_dir = OUTPUT_DIR
    output_dir.mkdir(exist_ok=True)

    addr1, addr2 = split_address(address.upper())
    company_up = company.strip().upper()

    doc = fitz.open(UTILITY_TEMPLATE)
    page = doc[0]
    pix = page.get_pixmap(dpi=72)
    fill_utility(page, company_up, addr1, addr2, pix)

    safe = (company_up
            .replace("/","-").replace("\\","-").replace(":","")
            .replace("*","").replace("?","").replace('"',"")
            .replace("<","").replace(">","").replace("|","")
            .replace("'",""))
    out = output_dir / f"Utility_{safe}.pdf"
    doc.save(str(out), garbage=4, deflate=True)
    doc.close()
    logger.info(f"Utility bill saved: {out.name}")
    return out


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

    # Page 1 — Arial cover, centered, #171717: company (Bold 22pt), term (Black 16pt).
    p = doc[0]
    _coc_set(p, COC_NAME_P1, company_up, size=22, color=COC_CLR_DARK,
             font=COC_FONT_TITLE, center=True)
    _coc_set(p, "10/16/2025 - 10/16/2026", f"{d['start_slash']} - {d['end_slash']}",
             size=16, color=COC_CLR_DARK, font=COC_FONT_DATE1, center=True)

    # Page 2 — Confirmation of Coverage: DejaVuSans 10pt black.
    p = doc[1]
    _coc_set(p, COC_NAME, company_up, size=10, color=COC_CLR_BLACK, font=COC_FONT_BODY)
    _coc_set(p, COC_ADDR1, addr1, size=10, color=COC_CLR_BLACK, font=COC_FONT_BODY)
    _coc_set(p, COC_ADDR2, addr2, size=10, color=COC_CLR_BLACK, font=COC_FONT_BODY)
    _coc_set(p, "10/16/2025", d["start_slash"], size=10, color=COC_CLR_BLACK, font=COC_FONT_BODY)
    _coc_set(p, "10/16/2026", d["end_slash"], size=10, color=COC_CLR_BLACK, font=COC_FONT_BODY)

    # Pages 3 & 4 — date only, DejaVuSans 10pt black.
    for i in (2, 3):
        _coc_set(doc[i], "10/16/2025", d["start_slash"], size=10, color=COC_CLR_BLACK, font=COC_FONT_BODY)

    # Page 5 — invoice: DejaVuSans 9pt black. insured, term (dashes), due date (+3 weeks).
    p = doc[4]
    _coc_set(p, COC_NAME, company_up, size=9, color=COC_CLR_BLACK, font=COC_FONT_BODY)
    _coc_set(p, "10-16-2025", d["start_dash"], size=9, color=COC_CLR_BLACK, font=COC_FONT_BODY)
    _coc_set(p, "10-16-2026", d["end_dash"], size=9, color=COC_CLR_BLACK, font=COC_FONT_BODY)
    _coc_set(p, "11/6/2025", d["due"], size=9, color=COC_CLR_BLACK, font=COC_FONT_BODY)

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


# ─── COVER WHALE FULL DOCUMENT (all 9 pages) ─────────────────────────────────
# The VIATIC template is a Frankenstein: pages 1-3 are VIATIC LLC (policy
# CUS09114581) but pages 4-9 were spliced in from OTHER quotes — page 4 header
# says CHARLIE HAULING LLC, pages 5-9 say DEKS TRANSPORT LLC, all under policy
# CUS09116580, plus page 6 carries a vehicle/driver schedule and page 7 an
# address — all belonging to those other carriers. generate_coverwhale() makes
# the whole document consistent with the target company.

ALT_POLICY      = "CUS09116580"            # baked-in policy on pages 4-9
ALT_COMPANY_P4  = "CHARLIE HAULING LLC"    # baked-in company on page 4 header
ALT_COMPANY_REST = "DEKS TRANSPORT LLC"    # baked-in company on pages 5-9 headers

# Original schedule values on page 6 / 7 (replaced every run)
CW_OLD_VINS = ["3AKJHHFG3PSNL7399", "1FUJHHDR4NLMZ0512",
               "3AKJHHFG5NSNE9904", "3AKJHHDRXNSNF2060"]
CW_OLD_GARAGE = "17250 DALLAS PKWY, DALLAS, TX 75248"

_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_FIRST_NAMES = ["JAMES", "ROBERT", "MICHAEL", "DAVID", "JOSE", "CARLOS", "JOHN",
                "LUIS", "DANIEL", "ANTHONY", "KEVIN", "BRIAN", "JASON", "ERIC",
                "JUAN", "MARK", "STEVEN", "ANDREW", "RAYMOND", "GREGORY",
                "MIGUEL", "DENNIS", "JERRY", "TYLER", "AARON", "HENRY"]
_LAST_NAMES = ["SMITH", "JOHNSON", "GARCIA", "MARTINEZ", "BROWN", "DAVIS",
               "RODRIGUEZ", "WILSON", "ANDERSON", "THOMAS", "HERNANDEZ", "MOORE",
               "JACKSON", "WHITE", "HARRIS", "CLARK", "LEWIS", "WALKER", "HALL",
               "ALLEN", "YOUNG", "KING", "WRIGHT", "TORRES", "NGUYEN", "REED"]


def _split_full_address(address: str):
    """(street, city, state, zip) from a free-form address string.
    Reuses split_address for the street / city-line split, then peels the zip
    and 2-letter state off the city line."""
    street, rest = split_address(address.upper())
    toks = rest.replace(",", " ").split()
    zipc = state = ""
    if toks and any(c.isdigit() for c in toks[-1]):
        zipc = toks.pop()
    if toks and len(toks[-1]) == 2 and toks[-1].isalpha():
        state = toks.pop()
    city = " ".join(toks)
    return street, city, state, zipc


def _rand_vin(rng):
    chars = "ABCDEFGHJKLMNPRSTUVWXYZ0123456789"   # real VINs omit I, O, Q
    return "".join(rng.choice(chars) for _ in range(17))


def _rand_vehicle_years(rng, n):
    """n years from {2022,2023,2024}, never all identical."""
    pool = [2022, 2023, 2024]
    while True:
        ys = [rng.choice(pool) for _ in range(n)]
        if len(set(ys)) > 1:
            return ys


def _rand_dob(rng):
    """('Mon, DD', 'YYYY') — birth year below 2000."""
    return f"{rng.choice(_MONTHS)}, {rng.randint(1, 28):02d}", str(rng.randint(1965, 1999))


def _rand_hire(rng):
    """('Mon, DD', '2025') — random hire date in 2025."""
    return f"{rng.choice(_MONTHS)}, {rng.randint(1, 28):02d}", "2025"


def _page_spans(page):
    """Flat list of every text span on a page (snapshot of the original layout)."""
    return [s for b in page.get_text("dict")["blocks"]
            for l in b.get("lines", []) for s in l["spans"] if s["text"].strip()]


def _find_span(spans, x0, y0, tol=2.5):
    """Rect of the span whose top-left corner is ~(x0, y0). None if not found."""
    for s in spans:
        bx0, by0, bx1, by1 = s["bbox"]
        if abs(bx0 - x0) <= tol and abs(by0 - y0) <= tol:
            return fitz.Rect(bx0, by0, bx1, by1)
    return None


def _bg_at(pix, x, y, pw, ph):
    """Background color (0-1 floats) of one rendered point. White if no pixmap."""
    if pix is None:
        return (1.0, 1.0, 1.0)
    sx, sy = pix.width / pw, pix.height / ph
    px = max(0, min(pix.width - 1, int(x * sx)))
    py = max(0, min(pix.height - 1, int(y * sy)))
    try:
        return tuple(c / 255.0 for c in pix.pixel(px, py)[:3])
    except Exception:
        return (1.0, 1.0, 1.0)


def _set_rect(page, rect, new_text, *, size, font=FONT_REG, color=(0.0, 0.0, 0.0),
              bg=(1.0, 1.0, 1.0), max_width=None):
    """Cover a span's rect with the row background color and write new_text at
    the same left baseline. Pass bg for shaded (zebra-striped) rows so the
    cover patch doesn't show. Pass max_width to shrink the font when the new
    value would overflow its column (e.g. a wide random VIN crashing into the
    next column)."""
    if rect is None:
        return
    fp = str(font)
    fobj = fitz.Font(fontfile=fp)
    fname = "F" + str(abs(hash(fp)) % 100000)
    if max_width:
        while size > 5.0 and fobj.text_length(new_text, fontsize=size) > max_width:
            size -= 0.2
    page.draw_rect(fitz.Rect(rect.x0 - 1, rect.y0 - 1, rect.x1 + 1, rect.y1 + 1),
                   color=bg, fill=bg, width=0)
    page.insert_text((rect.x0, rect.y1 - 1.0), new_text,
                     fontfile=fp, fontname=fname, fontsize=size, color=color)


def fill_header_alt(page, old_company, old_policy, company, policy, pix):
    """Top-right header on pages 4-9 (different baked-in company/policy than
    pages 1-3). Right-aligns to the same x=569 edge, DejaVu 7.36 #363636."""
    replace_on_page(page, f"TGL Policy #:  {old_policy}", f"TGL Policy #:  {policy}",
                    fontsize=7.36, top_right_x=569.0, pix=pix, color=CLR_HEADER)
    replace_on_page(page, old_company, company,
                    fontsize=7.36, top_right_x=569.0, pix=pix,
                    x_min=400, y_max=80, color=CLR_HEADER)


def fill_page6_schedule(page, street, city, state, zipc, rng, pix=None):
    """Page 6 — vehicle schedule (VIN + year), garage location, driver schedule
    (names, DOB, hire date). Coordinate-targeted because years/dates repeat.
    The tables are zebra-striped, so each cell is covered with its row's actual
    background color (sampled from pix at a reliably-blank probe point)."""
    spans = _page_spans(page)
    pw, ph = page.rect.width, page.rect.height

    # Vehicle VIN (x0≈48.4) + Year (x0≈156.5) per row. Coordinate-targeted: '2022'
    # repeats 6× on the page. Row bg sampled at x=152 (blank gap before Year).
    # VINs are width-fitted (max_width=103) so a wide random VIN never crashes
    # into the Year column at x=156.5.
    veh_y = [184.7, 206.8, 228.9, 251.0]
    for y, yr in zip(veh_y, _rand_vehicle_years(rng, len(veh_y))):
        row_bg = _bg_at(pix, 152, y, pw, ph)
        _set_rect(page, _find_span(spans, 48.4, y), _rand_vin(rng), size=9.6,
                  bg=row_bg, max_width=103)
        _set_rect(page, _find_span(spans, 156.5, y), str(yr), size=9.6, bg=row_bg)

    # Garage location → company address. Coordinate-targeted (not replace_on_page)
    # because the bold "Garage Location:" label sits flush to the value's left, so
    # left-sampling the bg catches a dark label glyph. Probe a blank spot at x=450.
    _set_rect(page, _find_span(spans, 134.7, 273.1),
              f"{street}, {city}, {state} {zipc}", size=9.6,
              bg=_bg_at(pix, 450, 273.1, pw, ph), max_width=420)

    # Driver schedule rows (size 6.3). Coords from the template's own spans.
    # Row bg probed at x=295 (blank gap between Years-Exp and Date-of-Hire) — the
    # driver table's far-left margin stays white even on grey rows, so we must
    # NOT sample at the left edge.
    rows = [
        dict(fn=(46.4, 494.0), ln=(89.2, 494.0),
             dob_md=(225.3, 490.3), dob_yr=(225.3, 497.8),
             hire_md=(311.0, 490.3), hire_yr=(311.0, 497.8)),
        dict(fn=(46.4, 516.0), ln=(89.2, 516.0),
             dob_md=(225.3, 512.2), dob_yr=(225.3, 519.8),
             hire_md=(311.0, 512.2), hire_yr=(311.0, 519.8)),
        dict(fn=(46.4, 537.9), ln=(89.2, 537.9),
             dob_md=(225.3, 534.2), dob_yr=(225.3, 541.7),
             hire_md=(311.0, 534.2), hire_yr=(311.0, 541.7)),
    ]
    for row in rows:
        row_bg = _bg_at(pix, 295, row["fn"][1], pw, ph)
        fn, ln = rng.choice(_FIRST_NAMES), rng.choice(_LAST_NAMES)
        dob_md, dob_yr = _rand_dob(rng)
        hire_md, hire_yr = _rand_hire(rng)
        vals = {"fn": fn, "ln": ln, "dob_md": dob_md, "dob_yr": dob_yr,
                "hire_md": hire_md, "hire_yr": hire_yr}
        for key, val in vals.items():
            x0, y0 = row[key]
            _set_rect(page, _find_span(spans, x0, y0), val, size=6.3, bg=row_bg)


def fill_page7_address(page, street, city, state, zipc, pix=None):
    """Page 7 — the Address / City / State / Zip row → company address. Each cell
    is width-fitted so a long street/city can't spill into the next column."""
    spans = _page_spans(page)
    pw, ph = page.rect.width, page.rect.height
    row_bg = _bg_at(pix, 250, 115.6, pw, ph)   # blank gap between Address & City
    # (x0, value, max column width before the next column starts)
    cells = [(48.8, street, 273), (327.8, city, 67), (399.5, state, 65), (469.2, zipc, 95)]
    for x0, val, mw in cells:
        _set_rect(page, _find_span(spans, x0, 115.6), val.upper(), size=10.2,
                  bg=row_bg, max_width=mw)


def generate_coverwhale(company: str, usdot: str, address: str, policy: str,
                        output_dir: Path = None, rng=None) -> Path:
    """Fill the WHOLE 9-page Cover Whale policy for one company — pages 1-3 like
    /new, plus pages 4-9 (headers + page-6 schedule + page-7 address) so the
    entire document is consistent with the target company."""
    if output_dir is None:
        output_dir = OUTPUT_DIR
    output_dir.mkdir(exist_ok=True)
    rng = rng or random

    company_up = company.strip().upper()
    addr1, addr2 = split_address(address.upper())
    street, city, state, zipc = _split_full_address(address)

    doc = fitz.open(TEMPLATE_PDF)

    p = doc[0]; pix = p.get_pixmap(dpi=72)
    fill_page1(p, company_up, usdot, addr1, addr2, policy, pix)

    p = doc[1]; pix = p.get_pixmap(dpi=72)
    fill_page2(p, company_up, addr1, addr2, policy, pix)

    p = doc[2]; pix = p.get_pixmap(dpi=72)          # page 3 — VIATIC header
    fill_page_header_only(p, company_up, policy, pix)

    p = doc[3]; pix = p.get_pixmap(dpi=72)          # page 4 — CHARLIE HAULING header
    fill_header_alt(p, ALT_COMPANY_P4, ALT_POLICY, company_up, policy, pix)

    p = doc[4]; pix = p.get_pixmap(dpi=72)          # page 5 — DEKS header
    fill_header_alt(p, ALT_COMPANY_REST, ALT_POLICY, company_up, policy, pix)

    p = doc[5]; pix = p.get_pixmap(dpi=72)          # page 6 — DEKS header + schedule
    fill_header_alt(p, ALT_COMPANY_REST, ALT_POLICY, company_up, policy, pix)
    fill_page6_schedule(p, street, city, state, zipc, rng, pix)

    p = doc[6]; pix = p.get_pixmap(dpi=72)          # page 7 — DEKS header + address
    fill_header_alt(p, ALT_COMPANY_REST, ALT_POLICY, company_up, policy, pix)
    fill_page7_address(p, street, city, state, zipc, pix)

    for i in (7, 8):                                # pages 8 & 9 — DEKS header only
        p = doc[i]; pix = p.get_pixmap(dpi=72)
        fill_header_alt(p, ALT_COMPANY_REST, ALT_POLICY, company_up, policy, pix)

    safe = (company_up
            .replace("/", "-").replace("\\", "-").replace(":", "")
            .replace("*", "").replace("?", "").replace('"', "")
            .replace("<", "").replace(">", "").replace("|", "")
            .replace("'", ""))
    out = output_dir / f"Cover Whale - {safe}.pdf"
    doc.save(str(out), garbage=4, deflate=True)
    doc.close()
    logger.info(f"Cover Whale full policy saved: {out.name}")
    return out


# ─── SCAN EFFECT ──────────────────────────────────────────────────────────────

def _scan_one(page, dpi: int):
    """Render one PDF page and apply the photo/scan effect. Returns a PIL Image."""
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    w, h = img.size

    # 1. Gray/warm paper tint — blend toward scanner-gray
    paper = Image.new("RGB", img.size, (235, 232, 225))
    img = Image.blend(img, paper, alpha=0.12)

    # 2. Reduce contrast & brightness (washed out / printed look)
    img = ImageEnhance.Contrast(img).enhance(0.82)
    img = ImageEnhance.Brightness(img).enhance(0.93)
    img = ImageEnhance.Sharpness(img).enhance(0.7)

    # 3. Gaussian noise (scanner grain)
    arr = np.array(img, dtype=np.int16)
    noise = np.random.normal(0, 4.5, arr.shape).astype(np.int16)
    arr = np.clip(arr + noise, 0, 255).astype(np.uint8)
    img = Image.fromarray(arr)

    # 4. Blur (scanner/camera softness)
    img = img.filter(ImageFilter.GaussianBlur(radius=0.7))

    # 5. Slight rotation (paper not aligned perfectly)
    angle = random.uniform(-0.7, 0.7)
    img = img.rotate(angle, resample=Image.BICUBIC, expand=False,
                     fillcolor=(230, 228, 222))

    # 6. Subtle edge shadow (very light, not a frame)
    shadow = np.ones((h, w), dtype=np.float32)
    margin_x = int(w * 0.03)
    margin_y = int(h * 0.025)

    for i in range(margin_x):
        f = (i / margin_x) ** 0.8
        shadow[:, i] *= (0.88 + 0.12 * f)
        shadow[:, w - 1 - i] *= (0.90 + 0.10 * f)
    for i in range(margin_y):
        f = (i / margin_y) ** 0.8
        shadow[i, :] *= (0.92 + 0.08 * f)
        shadow[h - 1 - i, :] *= (0.88 + 0.12 * f)

    img_arr = np.array(img, dtype=np.float32)
    for c in range(3):
        img_arr[:, :, c] *= shadow
    img = Image.fromarray(np.clip(img_arr, 0, 255).astype(np.uint8))

    # 7. Slight color temperature shift (warm/yellowish like old scanner)
    final_arr = np.array(img, dtype=np.int16)
    final_arr[:, :, 0] = np.clip(final_arr[:, :, 0] + 3, 0, 255)   # slight red boost
    final_arr[:, :, 2] = np.clip(final_arr[:, :, 2] - 4, 0, 255)   # slight blue drop
    return Image.fromarray(final_arr.astype(np.uint8))


def scannify_pdf(input_path: Path, output_dir: Path = None, dpi: int = 250) -> list[Path]:
    """
    Take a clean PDF and produce JPG images (first 3 pages only)
    that look like photos/scans of a printed document.
    Files are named with the local-time timestamp MMDDYYYYHHMMSS captured
    per page at save time. Returns list of JPG paths.
    """
    if output_dir is None:
        output_dir = input_path.parent

    doc = fitz.open(input_path)
    jpg_paths = []
    num_pages = min(3, len(doc))

    for page_num in range(num_pages):
        img = _scan_one(doc[page_num], dpi)

        # Save as JPG — filename is local-time MMDDYYYYHHMMSS, captured per page.
        # If two pages land in the same second, append _2, _3, ... to avoid overwrite.
        stamp = datetime.now().strftime("%m%d%Y%H%M%S")
        jpg_path = output_dir / f"{stamp}.jpg"
        dup = 2
        while jpg_path.exists():
            jpg_path = output_dir / f"{stamp}_{dup}.jpg"
            dup += 1
        img.save(str(jpg_path), "JPEG", quality=88)
        jpg_paths.append(jpg_path)

    doc.close()
    logger.info(f"Scanned JPGs saved: {[p.name for p in jpg_paths]}")
    return jpg_paths


def scannify_to_pdf(input_path: Path, output_dir: Path = None, dpi: int = 200) -> Path:
    """Scan EVERY page of a PDF and combine them into one scanned-look PDF.
    Each page is JPEG-compressed so the output stays small. Returns the path."""
    if output_dir is None:
        output_dir = input_path.parent

    src = fitz.open(input_path)
    out_doc = fitz.open()
    for i in range(len(src)):
        img = _scan_one(src[i], dpi)
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=85)
        w, h = img.size
        page = out_doc.new_page(width=w * 72.0 / dpi, height=h * 72.0 / dpi)
        page.insert_image(page.rect, stream=buf.getvalue())
    src.close()

    stamp = datetime.now().strftime("%m%d%Y%H%M%S")
    base = input_path.stem
    if base.startswith("COC_"):           # drop the doc-type prefix from the scan name
        base = base[len("COC_"):]
    out = output_dir / f"{base}_scanned_{stamp}.pdf"
    out_doc.save(str(out), garbage=4, deflate=True)
    out_doc.close()
    logger.info(f"Scanned PDF saved: {out.name}")
    return out


# ─── MAIN ────────────────────────────────────────────────────────────────────

def generate():
    print("\n" + "=" * 57)
    print("  Cover Whale PDF Generator  v3")
    print("=" * 57)
    logger.info("=" * 40)
    logger.info("Batch generation started")

    print("\n[1/3] Fonts ...")
    ensure_fonts()
    print("      DejaVuSansCondensed      OK")
    print("      DejaVuSansCondensed-Bold OK")

    print(f"\n[2/3] Reading: {EXCEL_FILE}")
    if not os.path.exists(EXCEL_FILE):
        print("  File not found — check EXCEL_FILE path in script."); sys.exit(1)

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    hdrs = [str(c.value or "").strip() for c in ws[2]]
    print(f"      Columns : {hdrs}")
    print(f"      Companies: {ws.max_row - 2}")

    def col(*names):
        for n in names:
            for i, h in enumerate(hdrs):
                if h.replace(" ", "").lower() == n.replace(" ", "").lower():
                    return i
        return None

    C_NAME  = col("Legal Name")
    C_USDOT = col("U SDOT Number", "USDOT Number", "USDOT")
    C_ADDR  = col("Physical Address")

    if any(c is None for c in [C_NAME, C_USDOT, C_ADDR]):
        print(f"  Cannot find required columns in: {hdrs}"); sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"\n[3/3] Generating -> {OUTPUT_DIR}")

    policy  = STARTING_POLICY
    count   = 0
    errors  = []

    for row in ws.iter_rows(min_row=3, values_only=True):
        if not any(row):
            continue

        company = str(row[C_NAME]  or "").strip().upper()
        usdot   = str(row[C_USDOT] or "").strip()
        address = str(row[C_ADDR]  or "").strip().upper()

        if not company:
            continue

        addr1, addr2 = split_address(address)

        print(f"\n  [{count+1:02d}] {company}")
        print(f"        Policy : {policy}")
        print(f"        USDOT  : {usdot}")
        print(f"        Street : {addr1}")
        print(f"        City   : {addr2}")
        logger.info(f"Generating [{count+1:02d}] {company} | Policy: {policy} | USDOT: {usdot}")

        try:
            doc = fitz.open(TEMPLATE_PDF)

            # ── Page 1 ──────────────────────────────────────────────────────
            p   = doc[0]
            pix = p.get_pixmap(dpi=72)   # snapshot BEFORE any changes
            fill_page1(p, company, usdot, addr1, addr2, policy, pix)

            # ── Page 2 ──────────────────────────────────────────────────────
            p   = doc[1]
            pix = p.get_pixmap(dpi=72)
            fill_page2(p, company, addr1, addr2, policy, pix)

            # ── Pages 3+ ────────────────────────────────────────────────────
            for i in range(2, len(doc)):
                p   = doc[i]
                pix = p.get_pixmap(dpi=72)
                fill_page_header_only(p, company, policy, pix)

            # ── Save ────────────────────────────────────────────────────────
            safe = (company
                    .replace("/","-").replace("\\","-").replace(":",  "")
                    .replace("*","").replace("?", "").replace('"', "")
                    .replace("<","").replace(">", "").replace("|",  "")
                    .replace("'",""))
            out = OUTPUT_DIR / f"Cover Whale - {safe}.pdf"
            doc.save(str(out), garbage=4, deflate=True)
            doc.close()
            print(f"        Saved  -> {out.name}")
            logger.info(f"Saved: {out.name}")

            policy = increment_policy(policy)
            count += 1

        except Exception as e:
            import traceback
            errors.append((company, str(e)))
            print(f"        ERROR: {e}")
            logger.error(f"Failed: {company} — {e}")
            traceback.print_exc()

    print("\n" + "=" * 57)
    print(f"  Done!  {count} PDFs  ->  {OUTPUT_DIR}")
    if errors:
        print(f"  Errors ({len(errors)}):")
        for n, e in errors:
            print(f"    * {n}: {e}")
    print("=" * 57 + "\n")
    logger.info(f"Batch complete: {count} PDFs, {len(errors)} errors")


if __name__ == "__main__":
    generate()
