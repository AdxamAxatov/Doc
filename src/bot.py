#!/usr/bin/env python3
"""
Cover Whale Telegram Bot
─────────────────────────
Wizard-style bot: collects company info field-by-field,
generates PDF(s), sends them back.

Run:  py bot.py
"""

import os, json, sys, csv, logging
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
if not TOKEN:
    sys.exit("TELEGRAM_BOT_TOKEN not set in .env")

from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove, BotCommand
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    ConversationHandler, ContextTypes, filters,
)

# Import core generation logic from generate.py
import fitz
from generate import (
    ensure_fonts, split_address, fill_page1, fill_page2,
    fill_page_header_only, increment_policy, scannify_pdf,
    generate_utility,
    PROJECT_DIR, OUTPUT_DIR, TEMPLATE_PDF,
    FONT_REG, FONT_BOLD, logger,
)

# ─── COMPANY DATABASE ────────────────────────────────────────────────────────

ASSETS_DIR = PROJECT_DIR / "assets"
ALL_COMPANIES_FILE = ASSETS_DIR / "All Companies.csv"
COMPANIES_DB = []

def load_companies_db():
    """Load all companies from CSV into memory on startup."""
    global COMPANIES_DB
    if not ALL_COMPANIES_FILE.exists():
        print(f"  Warning: {ALL_COMPANIES_FILE} not found — lookup disabled")
        return
    with open(ALL_COMPANIES_FILE, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            COMPANIES_DB.append({
                "name":    row.get("Legal Name", "").strip(),
                "usdot":   row.get("U SDOT Number", "").strip(),
                "address": row.get("Physical Address", "").strip(),
            })
    print(f"  Loaded {len(COMPANIES_DB)} companies from {ALL_COMPANIES_FILE.name}")
    logger.info(f"Loaded {len(COMPANIES_DB)} companies from {ALL_COMPANIES_FILE.name}")

def search_companies(query: str, max_results: int = 10):
    """Case-insensitive partial match on company name."""
    q = query.strip().upper()
    if not q:
        return []
    results = []
    for co in COMPANIES_DB:
        if q in co["name"].upper():
            results.append(co)
            if len(results) >= max_results:
                break
    return results

# ─── POLICY STATE ─────────────────────────────────────────────────────────────

STATE_FILE = ASSETS_DIR / "policy_state.json"
DEFAULT_POLICY = "CUS09116674"

def load_policy() -> str:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text())["policy"]
        except Exception:
            pass
    return DEFAULT_POLICY

def save_policy(policy: str):
    STATE_FILE.write_text(json.dumps({"policy": policy}))

# ─── CONVERSATION STATES ───────────────────────────────────────────────────────

ASK_NAME, ASK_PICK, ASK_USDOT, ASK_ADDR, ASK_MORE, ASK_SCAN = range(6)
UT_NAME, UT_PICK, UT_ADDR, UT_SCAN = range(10, 14)

YES_NO = ReplyKeyboardMarkup([["Yes", "No"]], one_time_keyboard=True, resize_keyboard=True)

# ─── HELPERS ──────────────────────────────────────────────────────────────────

def make_pdf(company: str, usdot: str, address: str, policy: str) -> Path:
    """Generate one PDF and return its path."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    addr1, addr2 = split_address(address.upper())
    company_up = company.strip().upper()

    doc = fitz.open(TEMPLATE_PDF)

    p = doc[0]; pix = p.get_pixmap(dpi=72)
    fill_page1(p, company_up, usdot, addr1, addr2, policy, pix)

    p = doc[1]; pix = p.get_pixmap(dpi=72)
    fill_page2(p, company_up, addr1, addr2, policy, pix)

    for i in range(2, len(doc)):
        p = doc[i]; pix = p.get_pixmap(dpi=72)
        fill_page_header_only(p, company_up, policy, pix)

    safe = (company_up
            .replace("/","-").replace("\\","-").replace(":","")
            .replace("*","").replace("?","").replace('"',"")
            .replace("<","").replace(">","").replace("|","")
            .replace("'",""))
    out = OUTPUT_DIR / f"Cover Whale - {safe}.pdf"
    doc.save(str(out), garbage=4, deflate=True)
    doc.close()
    return out

def add_company(ctx, company_data):
    """Add a company dict to the session list."""
    ctx.user_data["companies"].append(company_data)
    return len(ctx.user_data["companies"])

# ─── HANDLERS ─────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Cover Whale PDF Generator\n\n"
        "Commands:\n"
        "  /new — generate a policy PDF\n"
        "  /utility — generate a utility bill\n"
        "  /scan — scan any PDF (or just send a PDF)\n"
        "  /policy — view current policy number\n"
        "  /setpolicy CUS09116674 — override policy number"
    )

async def cmd_policy(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"Current policy: {load_policy()}")

async def cmd_setpolicy(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    args = ctx.args
    if not args:
        await update.message.reply_text("Usage: /setpolicy CUS09116674")
        return
    save_policy(args[0].strip())
    await update.message.reply_text(f"Policy set to {args[0].strip()}")

async def cmd_new(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["companies"] = []
    await update.message.reply_text(
        "What's the company name?",
        reply_markup=ReplyKeyboardRemove()
    )
    return ASK_NAME

async def got_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    user = update.effective_user.first_name
    logger.info(f"[{user}] Search: \"{query}\"")
    results = search_companies(query)

    if len(results) == 1:
        # Exact single match — auto-fill
        co = results[0]
        n = add_company(ctx, co)
        logger.info(f"[{user}] Match: {co['name']} (USDOT: {co['usdot']})")
        await update.message.reply_text(
            f"Found: {co['name']}\n"
            f"USDOT: {co['usdot']}\n"
            f"Address: {co['address']}\n\n"
            f"Company {n} added. Add another?",
            reply_markup=YES_NO
        )
        return ASK_MORE

    elif len(results) > 1:
        # Multiple matches — let user pick
        logger.info(f"[{user}] Multiple matches: {len(results)} results")
        ctx.user_data["search_results"] = results
        lines = [f"{i+1}. {co['name']}" for i, co in enumerate(results)]
        buttons = [[str(i+1)] for i in range(len(results))]
        buttons.append(["None of these"])
        await update.message.reply_text(
            f"Found {len(results)} matches:\n\n" + "\n".join(lines) +
            "\n\nPick a number, or 'None of these' for manual entry.",
            reply_markup=ReplyKeyboardMarkup(buttons, one_time_keyboard=True, resize_keyboard=True)
        )
        return ASK_PICK

    else:
        # No match — manual entry
        logger.info(f"[{user}] No match — manual entry")
        ctx.user_data["current"] = {"name": query}
        await update.message.reply_text(
            f"No match found for \"{query}\".\n"
            "Let's enter manually.\n\n"
            "USDOT number?"
        )
        return ASK_USDOT

async def got_pick(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()

    if text.lower() == "none of these":
        # Fall back to manual — ask for the name again fresh
        await update.message.reply_text(
            "Enter the company name for manual entry:",
            reply_markup=ReplyKeyboardRemove()
        )
        ctx.user_data["manual_mode"] = True
        return ASK_NAME

    results = ctx.user_data.get("search_results", [])
    try:
        idx = int(text) - 1
        if 0 <= idx < len(results):
            co = results[idx]
            n = add_company(ctx, co)
            await update.message.reply_text(
                f"Selected: {co['name']}\n"
                f"USDOT: {co['usdot']}\n"
                f"Address: {co['address']}\n\n"
                f"Company {n} added. Add another?",
                reply_markup=YES_NO
            )
            return ASK_MORE
    except ValueError:
        pass

    await update.message.reply_text("Invalid choice. Pick a number from the list, or 'None of these'.")
    return ASK_PICK

async def got_usdot(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["current"]["usdot"] = update.message.text.strip()
    await update.message.reply_text("Physical address?")
    return ASK_ADDR

async def got_addr(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["current"]["address"] = update.message.text.strip()
    n = add_company(ctx, ctx.user_data.pop("current"))
    await update.message.reply_text(
        f"Company {n} added. Add another?",
        reply_markup=YES_NO
    )
    return ASK_MORE

async def got_more_yes(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.pop("manual_mode", None)
    await update.message.reply_text(
        "What's the company name?",
        reply_markup=ReplyKeyboardRemove()
    )
    return ASK_NAME

async def got_more_no(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.pop("manual_mode", None)
    companies = ctx.user_data.get("companies", [])
    if not companies:
        await update.message.reply_text("No companies to generate.", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END

    await update.message.reply_text(
        f"Generating {len(companies)} PDF(s)...",
        reply_markup=ReplyKeyboardRemove()
    )

    policy = load_policy()
    errors = []
    generated_paths = []

    for co in companies:
        try:
            logger.info(f"Bot PDF: {co['name']} | Policy: {policy} | USDOT: {co['usdot']}")
            path = make_pdf(co["name"], co["usdot"], co["address"], policy)
            generated_paths.append(path)
            with open(path, "rb") as f:
                await update.message.reply_document(
                    document=f,
                    filename=path.name,
                    caption=f"{co['name']} — {policy}",
                    read_timeout=60,
                    write_timeout=60,
                    connect_timeout=60,
                )
            policy = increment_policy(policy)
        except Exception as e:
            logger.error(f"Bot PDF failed: {co['name']} — {e}")
            errors.append(f"{co['name']}: {e}")

    save_policy(policy)

    if errors:
        await update.message.reply_text("Errors:\n" + "\n".join(errors))

    # Store generated paths for potential scan step
    ctx.user_data["generated_paths"] = generated_paths

    if generated_paths:
        await update.message.reply_text(
            f"Done! Next policy will be: {policy}\n\n"
            "Want a scanned version of the PDF(s)?",
            reply_markup=YES_NO
        )
        return ASK_SCAN
    else:
        return ConversationHandler.END


async def got_scan_yes(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    paths = ctx.user_data.get("generated_paths", [])
    if not paths:
        await update.message.reply_text("No PDFs to scan.", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END

    await update.message.reply_text(
        f"Creating scanned version(s)...",
        reply_markup=ReplyKeyboardRemove()
    )

    for path in paths:
        try:
            jpg_paths = scannify_pdf(path)
            for jpg_path in jpg_paths:
                with open(jpg_path, "rb") as f:
                    await update.message.reply_document(
                        document=f,
                        filename=jpg_path.name,
                        read_timeout=60,
                        write_timeout=60,
                        connect_timeout=60,
                    )
        except Exception as e:
            logger.error(f"Scan effect failed: {path.name} — {e}")
            await update.message.reply_text(f"Error scanning {path.name}: {e}")

    await update.message.reply_text("Done!")
    return ConversationHandler.END


async def got_scan_no(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("All done!", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# ─── UTILITY BILL HANDLERS ───────────────────────────────────────────────────

async def cmd_utility(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Utility bill generator\n\nCompany name?",
        reply_markup=ReplyKeyboardRemove()
    )
    return UT_NAME

async def _ut_generate_and_send(update, ctx, company, address):
    """Shared helper: generate utility bill, send it, ask about scan."""
    await update.message.reply_text("Generating utility bill...", reply_markup=ReplyKeyboardRemove())
    try:
        path = generate_utility(company, address)
        ctx.user_data["generated_paths"] = [path]
        with open(path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=path.name,
                caption=f"Utility bill — {company.upper()}",
                read_timeout=60, write_timeout=60, connect_timeout=60,
            )
        logger.info(f"Utility bill sent: {company} | {address}")
        await update.message.reply_text("Want a scanned version?", reply_markup=YES_NO)
        return UT_SCAN
    except Exception as e:
        logger.error(f"Utility bill failed: {company} — {e}")
        await update.message.reply_text(f"Error: {e}")
        return ConversationHandler.END

async def got_ut_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    user = update.effective_user.first_name
    logger.info(f"[{user}] Utility search: \"{query}\"")
    results = search_companies(query)

    if len(results) == 1:
        co = results[0]
        logger.info(f"[{user}] Utility match: {co['name']}")
        return await _ut_generate_and_send(update, ctx, co["name"], co["address"])

    elif len(results) > 1:
        logger.info(f"[{user}] Utility multiple matches: {len(results)}")
        ctx.user_data["ut_search_results"] = results
        lines = [f"{i+1}. {co['name']}" for i, co in enumerate(results)]
        buttons = [[str(i+1)] for i in range(len(results))]
        buttons.append(["None of these"])
        await update.message.reply_text(
            f"Found {len(results)} matches:\n\n" + "\n".join(lines) +
            "\n\nPick a number, or 'None of these' for manual entry.",
            reply_markup=ReplyKeyboardMarkup(buttons, one_time_keyboard=True, resize_keyboard=True)
        )
        return UT_PICK

    else:
        logger.info(f"[{user}] Utility no match — manual entry")
        ctx.user_data["ut_company"] = query
        await update.message.reply_text(
            f"No match found for \"{query}\".\n\nEnter the address:"
        )
        return UT_ADDR

async def got_ut_pick(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()

    if text.lower() == "none of these":
        await update.message.reply_text(
            "Enter the company name for manual entry:",
            reply_markup=ReplyKeyboardRemove()
        )
        return UT_NAME

    results = ctx.user_data.get("ut_search_results", [])
    try:
        idx = int(text) - 1
        if 0 <= idx < len(results):
            co = results[idx]
            return await _ut_generate_and_send(update, ctx, co["name"], co["address"])
    except ValueError:
        pass

    await update.message.reply_text("Invalid choice. Pick a number from the list, or 'None of these'.")
    return UT_PICK

async def got_ut_addr(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    company = ctx.user_data["ut_company"]
    address = update.message.text.strip()
    return await _ut_generate_and_send(update, ctx, company, address)

async def got_ut_scan_yes(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
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

async def got_ut_scan_no(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("All done!", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# ─── SCAN ANY PDF ─────────────────────────────────────────────────────────────

async def cmd_scan(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Send me a PDF file and I'll create a scanned version.",
        reply_markup=ReplyKeyboardRemove()
    )

async def handle_pdf_file(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("Please send a PDF file.")
        return

    await update.message.reply_text("Creating scanned version...")

    try:
        file = await doc.get_file()
        pdf_path = OUTPUT_DIR / doc.file_name
        OUTPUT_DIR.mkdir(exist_ok=True)
        await file.download_to_drive(str(pdf_path))

        jpg_paths = scannify_pdf(pdf_path)

        for jpg_path in jpg_paths:
            with open(jpg_path, "rb") as f:
                await update.message.reply_document(
                    document=f, filename=jpg_path.name,
                    read_timeout=60, write_timeout=60, connect_timeout=60,
                )

        logger.info(f"Scanned PDF: {doc.file_name} -> {len(jpg_paths)} pages")
        await update.message.reply_text("Done!")
    except Exception as e:
        logger.error(f"Scan PDF failed: {doc.file_name} — {e}")
        await update.message.reply_text(f"Error: {e}")

# ─── GENERAL ──────────────────────────────────────────────────────────────────

async def cmd_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Cancelled.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    ensure_fonts()
    load_companies_db()

    from telegram.ext import Defaults
    from telegram.request import HTTPXRequest
    request = HTTPXRequest(read_timeout=60, write_timeout=60, connect_timeout=60)
    app = Application.builder().token(TOKEN).request(request).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("new", cmd_new)],
        states={
            ASK_NAME:  [MessageHandler(filters.TEXT & ~filters.COMMAND, got_name)],
            ASK_PICK:  [MessageHandler(filters.TEXT & ~filters.COMMAND, got_pick)],
            ASK_USDOT: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_usdot)],
            ASK_ADDR:  [MessageHandler(filters.TEXT & ~filters.COMMAND, got_addr)],
            ASK_MORE:  [
                MessageHandler(filters.Regex(r"(?i)^yes$"), got_more_yes),
                MessageHandler(filters.Regex(r"(?i)^no$"),  got_more_no),
            ],
            ASK_SCAN:  [
                MessageHandler(filters.Regex(r"(?i)^yes$"), got_scan_yes),
                MessageHandler(filters.Regex(r"(?i)^no$"),  got_scan_no),
            ],
        },
        fallbacks=[CommandHandler("cancel", cmd_cancel)],
    )

    util_conv = ConversationHandler(
        entry_points=[CommandHandler("utility", cmd_utility)],
        states={
            UT_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_ut_name)],
            UT_PICK: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_ut_pick)],
            UT_ADDR: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_ut_addr)],
            UT_SCAN: [
                MessageHandler(filters.Regex(r"(?i)^yes$"), got_ut_scan_yes),
                MessageHandler(filters.Regex(r"(?i)^no$"),  got_ut_scan_no),
            ],
        },
        fallbacks=[CommandHandler("cancel", cmd_cancel)],
    )

    app.add_handler(CommandHandler("start",     cmd_start))
    app.add_handler(CommandHandler("scan",      cmd_scan))
    app.add_handler(CommandHandler("policy",    cmd_policy))
    app.add_handler(CommandHandler("setpolicy", cmd_setpolicy))
    app.add_handler(conv)
    app.add_handler(util_conv)
    app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf_file))

    async def post_init(application):
        await application.bot.set_my_commands([
            BotCommand("new",       "Generate a policy PDF"),
            BotCommand("utility",   "Generate a utility bill"),
            BotCommand("scan",      "Scan any PDF document"),
            BotCommand("policy",    "View current policy number"),
            BotCommand("setpolicy", "Override policy number"),
            BotCommand("cancel",    "Cancel current operation"),
        ])

    app.post_init = post_init
    logger.info("Bot started")
    print("Bot running...")
    app.run_polling()

if __name__ == "__main__":
    main()
