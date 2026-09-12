import os, sys, tempfile, importlib.util
os.environ.update(API_ID="12345", API_HASH="a"*32, BOT_TOKEN="123:abc", SECURITY_CODE="x", WORKDIR=tempfile.mkdtemp())
spec = importlib.util.spec_from_file_location("bot", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "bot.py"))
bot = importlib.util.module_from_spec(spec); sys.modules["bot"] = bot; spec.loader.exec_module(bot)
print("module loaded; version", bot.VERSION, "pdf_in", bot.HAS_PDF_IN, "pdf_out", bot.HAS_PDF_OUT)

# --- text → blocks → chunks → reassemble round trip
src = "Chapter 1\n\nThe quick brown fox.\nIt jumped over the lazy dog.\n\nChapter 2\n\n" + ("Lorem ipsum dolor sit amet. " * 400) + "\n\nThe end."
blocks = bot.blocks_from_plain(src)
print("blocks:", len(blocks), "headings:", sum(1 for b in blocks if b[0]))
text = bot.blocks_to_text(blocks)
chunks = bot.build_chunks(text, limit=1000)
print("chunks:", len(chunks), "max len:", max(len(c.text) for c in chunks))
assert all(len(c.text) <= 1000 for c in chunks), "chunk too long"
out = bot.reassemble(chunks, [c.text for c in chunks])
print("roundtrip identical:", out == text)
assert out == text

# --- writers
d = tempfile.mkdtemp()
for fmt in bot.OUTPUT_FORMATS:
    p = os.path.join(d, f"out.{fmt}")
    bot.write_output(fmt, text, p, "Test Title", "hi", source_text=None)
    print(fmt, os.path.getsize(p), "bytes")
# bilingual
p = os.path.join(d, "bi.docx")
bot.write_output("docx", text, p, "Test", "ur", source_text=text)
print("bilingual docx", os.path.getsize(p))

# --- extractors (txt, md, html, docx, epub, fb2, pdf)
open(os.path.join(d,"out.md"),"w").write("# Chapter 1\n\nThe quick brown fox.\n")
for fmt in ["txt","md","html","docx","epub"]:
    p = os.path.join(d, f"out.{fmt}")
    b = bot.extract_document(p, "."+fmt)
    print("extract", fmt, "->", len(b), "chars")
    assert "quick brown fox" in b, fmt
fb2 = os.path.join(d, "x.fb2")
open(fb2,"w").write('<?xml version="1.0"?><FictionBook><body><section><title><p>Ch 1</p></title><p>The quick brown fox.</p></section></body></FictionBook>')
print("extract fb2 ->", bot.extract_document(fb2, ".fb2")[:60])
if bot.HAS_PDF_OUT and bot.HAS_PDF_IN:
    print("extract pdf ->", bot.extract_document(os.path.join(d,"out.pdf"), ".pdf")[:60])

# --- split by output size
parts = bot.split_blocks_by_output_size(text, 3)
print("split parts (3KB):", len(parts), [len(x.encode()) for x in parts][:5])

# --- glossary
g = bot.Glossary({"Fox": "लोमड़ी"})
prot = g.protect("The Fox ran.")
print("glossary protect:", prot, "-> restore:", g.restore(prot))

# --- misc helpers
print(bot.safe_filename('bad|name:*?.txt'), bot.fmt_size(1234567), bot.fmt_time(3725), bot.parse_size_kb("1.5mb"))
print("cache:", bot.TranslationCache.__name__, "store users:", len(bot.store.data["users"]))
print("ALL SMOKE TESTS PASSED")
