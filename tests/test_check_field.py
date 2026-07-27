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
