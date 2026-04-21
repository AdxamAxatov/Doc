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
import os, sys, urllib.request, zipfile, io, logging, random
from datetime import datetime
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

# Arial fonts (Windows system)
ARIAL_REG  = Path("C:/Windows/Fonts/arial.ttf")
ARIAL_BOLD = Path("C:/Windows/Fonts/arialbd.ttf")

# ─── UTILITY BILL CONFIG ─────────────────────────────────────────────────────
UTILITY_TEMPLATE = ASSETS_DIR / "template" / "Utility_IVORY JULIUS CHRISTOPHER.pdf"
UT_NAME    = "IVORY JULIUS CHRISTOPHER"
UT_ADDR1   = "10318 CHEEVES,"
UT_ADDR2   = "HOUSTON, TX 77016"

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


# ─── SCAN EFFECT ──────────────────────────────────────────────────────────────

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
        page = doc[page_num]
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
        r, g, b = img.split()
        r = ImageEnhance.Brightness(r.convert("L").convert("RGB")).enhance(1.02)
        # simpler: just adjust channels via numpy
        final_arr = np.array(img, dtype=np.int16)
        final_arr[:, :, 0] = np.clip(final_arr[:, :, 0] + 3, 0, 255)   # slight red boost
        final_arr[:, :, 2] = np.clip(final_arr[:, :, 2] - 4, 0, 255)   # slight blue drop
        img = Image.fromarray(final_arr.astype(np.uint8))

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
