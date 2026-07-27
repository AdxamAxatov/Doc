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
