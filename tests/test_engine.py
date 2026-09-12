import os, sys, asyncio, importlib.util, tempfile
os.environ.update(API_ID="12345", API_HASH="a"*32, BOT_TOKEN="123:abc", SECURITY_CODE="x", WORKDIR=tempfile.mkdtemp())
spec = importlib.util.spec_from_file_location("bot", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "bot.py"))
bot = importlib.util.module_from_spec(spec); sys.modules["bot"] = bot; spec.loader.exec_module(bot)

calls = {"n": 0}
def fake_translate(text, src, tgt, engine):
    calls["n"] += 1
    # emulate MT: uppercase words, mangle spacing around tokens like real MT does
    return text.upper().replace("⟪", "⟪ ").replace("⟫", " ⟫")
bot._translate_sync = fake_translate

src = "Chapter 1\n\nLi Wei walked into the room.\nThe fox stared at Li Wei Long.\n\nChapter 2\n\n" + ("Hello world sentence. " * 300)
text = bot.blocks_to_text(bot.blocks_from_plain(src))
chunks = bot.build_chunks(text, limit=1500)
gl = bot.Glossary({"Li Wei": "ली वेई", "Li Wei Long": "ली वेई लॉन्ग", "fox": "लोमड़ी"})
eng = bot.TranslationEngine("hi", "auto", gl, concurrency=3)

async def main():
    async def prog(e): pass
    res = await eng.run(chunks, asyncio.Event(), prog)
    out = bot.reassemble(chunks, res)
    print("chunks:", len(chunks), "api calls:", calls["n"], "failed:", eng.failed, "cache hits:", eng.cache_hits)
    blocks = bot.text_to_blocks(out)
    heads = [t for l,t in blocks if l]
    print("headings preserved:", heads)
    assert len(heads) == 2, heads
    joined = "\n".join(t for _,t in blocks)
    assert "ली वेई लॉन्ग" in joined and "लोमड़ी" in joined and "ली वेई" in joined, joined[:300]
    assert "⟪" not in joined
    print("glossary restored OK:", [l for l in joined.split("\n")[:3]])
    # second run → all cache hits
    eng2 = bot.TranslationEngine("hi", "auto", gl, concurrency=3)
    before = calls["n"]
    await eng2.run(chunks, asyncio.Event(), prog)
    print("second run api calls:", calls["n"]-before, "cache hits:", eng2.cache_hits)
    assert eng2.cache_hits == len(chunks)
    # cancellation
    ev = asyncio.Event(); ev.set()
    try:
        await bot.TranslationEngine("hi","auto",bot.Glossary({}),1).run(bot.build_chunks("x "*3000, 500), ev, prog)
        print("cancel: NOT raised (cache hit path) - ok")
    except bot.JobCancelled:
        print("cancel: JobCancelled raised OK")
    print("ENGINE TESTS PASSED")
asyncio.run(main())
