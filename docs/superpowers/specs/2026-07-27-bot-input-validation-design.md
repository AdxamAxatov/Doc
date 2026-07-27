# Telegram bot input validation — design

**Date:** 2026-07-27
**Status:** Approved, ready for implementation plan

## Problem

`/utility` produced a bill with an email address in the company-name slot and a
company name in the address slot. The PDF pipeline was not at fault. From
`log/coverwhale.log:1529`:

```
10:49:41  Utility search: "daylenis@leoempireservicesllc.com"
10:49:41  Utility no match — manual entry
10:49:48  Utility bill sent: daylenis@leoempireservicesllc.com | LEO EMPIRE SERVICES LLC
```

The email was typed at the `Company name?` prompt and the company name at the
`Enter the address:` prompt. Confirmed against
`output/Utility_DAYLENIS@LEOEMPIRESERVICESLLC.COM.pdf`:

| Slot | Rendered |
| --- | --- |
| Name, y=124, 12pt bold | `DAYLENIS@LEOEMPIRESERVICESLLC.COM` |
| Address line 1, y=167 | `LEO EMPIRE SERVICES LLC` |
| Address line 2, y=177 | *(blank — `split_address` found no comma, so line 2 was empty and the original was whited out)* |

A control run with `SAINT LOUIS FAM INC` plus a real address placed all three
values correctly in both the header block and the bottom payment slip, so
`fill_utility` and `replace_on_page` in `src/generate.py` work as intended.

The defect is that the conversation handlers in `src/bot.py` accept any string
and go straight to generation. Nothing distinguishes an email from a company
name, or a company name from an address.

## Goals

- A value that is obviously not a company name, or obviously not an address, is
  rejected with an explanation and the same question is asked again.
- A legitimate but unusual value can still be forced through.
- No silent generation from mis-ordered input.

## Non-goals

- Verifying that a company or address actually exists. This checks *shape*, not truth.
- Validating the USDOT field (`got_usdot`, `got_cw_usdot`).
- Parsing multi-line pasted lead blobs. Considered and rejected as scope creep;
  rejection alone fixes the reported failure.
- Changing anything in `src/generate.py`.

## Scope: eight call sites

Four flows take free-text name and address:

| Flow | Name handler | Address handler |
| --- | --- | --- |
| `/utility` | `got_ut_name` (src/bot.py:384) | `got_ut_addr` (src/bot.py:438) |
| `/new` | `got_name` (src/bot.py:166) | `got_addr` (src/bot.py:248) |
| `/coc` | `got_coc_name` (src/bot.py:496) | `got_coc_addr` (src/bot.py:538) |
| `/coverwhale` | `got_cw_name` (src/bot.py:601) | `got_cw_addr` (src/bot.py:648) |

Addresses arriving from `All Companies.csv` via the single-match and pick-a-number
paths are **not** validated. They are trusted data and continue straight to
generation. Only manually typed input passes through a validator.

## Component 1: pure validators

Two functions in a new module `src/validation.py`. Each returns an error message
to show the user, or `None` when the value is acceptable.

They live in their own module rather than in `src/bot.py` because `bot.py` calls
`sys.exit()` at import time when `TELEGRAM_BOT_TOKEN` is unset, and pytest is not
installed in this project. A module importing nothing outside the standard
library keeps the validator tests trivial to run.

```python
def looks_like_company(s: str) -> str | None
def looks_like_address(s: str) -> str | None
```

`looks_like_company` rejects when the value:

- is shorter than 2 characters
- contains an email address
- contains a URL (`http://`, `https://`, or `www.`)
- contains no letters at all
- looks like a mailing address — contains a 5-digit ZIP, or a comma followed by
  a US state abbreviation as a whole word (`, TX`, `, FL`)

The address-shape rule closes the other half of the reported swap. Without it,
an address typed at the company-name prompt sails through: all 16,742 real
Physical Address values in `All Companies.csv` were accepted as company names.
The rule reuses `STATE_RE`'s alternation rather than duplicating the state list,
and its word boundaries keep it off legitimate names — `SMITH, INC` does not
match via `IN`, `ACME, DELIVERY LLC` does not match via `DE`.

A "more digits than letters" rule was drafted and then dropped: it wrongly
rejected the real companies `7573 LLC` and `1524 INC`, and caught nothing that
the no-letters rule misses.

`looks_like_address` rejects when the value:

- is shorter than 6 characters
- contains an email address
- contains no digits at all
- contains none of: comma, newline, 5-digit ZIP, or a US state abbreviation

The address rule is deliberately loose — any one of comma / newline / ZIP / state
is enough. Over-accepting is the safe direction because the override exists,
whereas over-rejecting blocks real work. The state-abbreviation pattern will also
match common English words (`IN`, `OR`, `ME`); this only ever makes the check
more permissive, never wrongly strict.

### Behavior against real data

Both rules were swept over every row of `assets/All Companies.csv` (16,742
companies) before this spec was finalized:

| Check | Rows rejected |
| --- | --- |
| `looks_like_company` over every Legal Name | 0 of 16,742 |
| `looks_like_address` over every Physical Address | 0 of 16,742 |
| `looks_like_company` over every Physical Address (should reject) | 16,544 of 16,742 (98.8%) |

The first two rows are the false-positive bar: no real record is ever blocked.
The third measures the address-shape rule catching an address typed at the
company-name prompt.

Every incident and control string behaves correctly. Rejected:
`daylenis@leoempireservicesllc.com` and `www.leoempire.com` as companies;
`LEO EMPIRE SERVICES LLC` and `no digits here at all` as addresses. Accepted:
`1 GUY 1 GIRL 1 TRUCK LLC`, `LALO'S TRUCKING INC`, `7573 LLC`, `GKH` as
companies; `265 Faulkner Dr, Niantic, IL 62551`,
`8514 Fenway Dr Houston TX 77036`, and the newline-separated CSV form as
addresses.

## Component 2: shared gate helper

```python
async def check_field(update, ctx, validator, slot, prompt) -> str | None:
    """Return the value to use, or None if the caller should re-enter its own state."""
```

`check_field` owns the whole rejection interaction: the error message, the
second-strike override keyboard, and the bookkeeping. Callers stay two lines
long and know nothing about strikes or pending values.

State is kept in `ctx.user_data` under two keys derived from `slot`:

- `_pend_<slot>` — the most recently rejected value, held so it can be forced through
- `_rej_<slot>` — how many times this field has been rejected this conversation

`ConversationHandler.END` does not clear `ctx.user_data`, so "this conversation"
has to be enforced explicitly. `_clear_validation_state(ctx)` drops every
`_rej_*` / `_pend_*` key and is called from `cmd_cancel` **and** all four entry
points — re-issuing `/utility` mid-conversation restarts without passing through
`cmd_cancel`. Without this, a value rejected in a cancelled conversation could be
forced into a later, unrelated document via `Use it anyway`: the same
silent-wrong-data failure this design exists to prevent.

Logic:

1. If the message is `Use it anyway`, clear both keys and return the pending value.
   If nothing is pending, re-send `prompt` and return `None`.
2. If the message is `Try again`, clear both keys, re-send `prompt`, return `None`.
   Clearing the counter means the next bad value gets a plain error again rather
   than jumping straight back to the override keyboard.
3. Run `validator`. On `None`, clear both keys and return the text.
4. Otherwise increment `_rej_<slot>`, store the text in `_pend_<slot>`, and reply
   with the error. On the second or later rejection, attach a
   `[Use it anyway] / [Try again]` keyboard.

The override buttons arrive as ordinary text within the state the handler already
returned, so **no new conversation states are added** to any `ConversationHandler`
and `main()` is untouched.

The keyboard is a module-level constant alongside the existing `YES_NO`
(src/bot.py:95):

```python
USE_ANYWAY = ReplyKeyboardMarkup([["Use it anyway"], ["Try again"]],
                                 one_time_keyboard=True, resize_keyboard=True)
```

## Component 3: call-site edits

Name handlers validate *before* searching, so an email never reaches
`search_companies`. Address handlers validate before generating.

```python
async def got_ut_addr(update, ctx):
    address = await check_field(update, ctx, looks_like_address,
                                "ut_addr", "Enter the address:")
    if address is None:
        return UT_ADDR
    return await _ut_generate_and_send(update, ctx, ctx.user_data["ut_company"], address)
```

Re-prompt strings reuse each flow's existing wording:

| Slot | Prompt |
| --- | --- |
| `ut_name`, `coc_name`, `cw_name` | `Company name?` |
| `name` | `What's the company name?` |
| `ut_addr`, `coc_addr` | `Enter the address:` |
| `addr`, `cw_addr` | `Physical address?` |

`got_pick`'s `None of these` branch already routes back to `ASK_NAME`, which now
validates on the way through. No change needed there.

## Logging

Rejections log at INFO through the existing `coverwhale` logger, matching the
current `[{user}] ...` format, so `log/coverwhale.log` shows why a run stalled:

```
[ibo] Rejected ut_name (strike 1): "daylenis@leoempireservicesllc.com"
[ibo] Override accepted ut_addr: "PO BOX 412"
```

## Testing

pytest is **not** installed. The three existing test files are standalone
scripts run with `./venv/bin/python tests/<file>.py` that accumulate into a
`failures` list, print `FAIL:` lines and `sys.exit(1)`, or print one `PASS:`
line. New tests follow that convention.

- `tests/test_input_validation.py` — table-driven cases over both validators.
  Rejections include the exact strings from the incident; acceptances include
  real names and addresses from `All Companies.csv`, plus `7573 LLC` and
  `1524 INC` as regressions against the dropped digit rule.
- `tests/test_check_field.py` — `check_field` driven with a minimal fake
  update/context: first rejection sends a bare error, the second attaches the
  override keyboard, `Use it anyway` returns the pending value, `Try again`
  clears it and resets the strike count, and a valid value clears both keys.
- `tests/test_handler_guards.py` — all eight handlers hold their conversation
  state on bad input and never reach a generation call, while a valid address
  still flows through.

The validators are pure, so most of the suite needs no Telegram mocking. Tests
that import `bot` must set `TELEGRAM_BOT_TOKEN` before the import.

## Risks

- **Over-rejection blocks real work.** Mitigated by the loose address rule and the
  second-strike override.
- **`ctx.user_data` key collision.** The `_pend_` / `_rej_` prefixes are new; no
  existing key in `src/bot.py` uses a leading underscore.
- **Override text collides with a real value.** Someone whose company is literally
  `Try again` would be misread. Accepted as negligible.
