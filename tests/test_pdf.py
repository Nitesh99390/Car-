import os, sys, importlib.util, tempfile
os.environ.update(API_ID="12345", API_HASH="a"*32, BOT_TOKEN="123:abc", SECURITY_CODE="x", WORKDIR=tempfile.mkdtemp())
spec = importlib.util.spec_from_file_location("bot", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "bot.py"))
bot = importlib.util.module_from_spec(spec); sys.modules["bot"] = bot; spec.loader.exec_module(bot)
d=tempfile.mkdtemp()
hi = "\u2063#2 अध्याय 1\nयह एक परीक्षण है। The Fox ran fast.\nदूसरा पैराग्राफ 123."
for lang, txt in [("hi",hi),("en","Chapter 1\nPlain English text here."),("ur","\u2063#2 باب 1\nیہ ایک امتحان ہے۔"),("bn","\u2063#2 অধ্যায় ১\nএটি একটি পরীক্ষা।")]:
    p=os.path.join(d,f"{lang}.pdf")
    bot.write_output("pdf", txt, p, f"Part 1 of 3 - Novel [{lang}]", lang, source_text="Chapter 1\nEnglish source.\nSecond para." if lang=="hi" else None)
    ext = bot.extract_document(p, ".pdf")
    print(lang, "font:", bot._pdf_font(lang), os.path.getsize(p), "B  extracted:", repr(ext[:90]))
# fb2 fix check
fb2=os.path.join(d,"x.fb2"); open(fb2,"w").write('<?xml version="1.0"?><FictionBook><body><section><title><p>Ch 1</p></title><p>The quick brown fox.</p></section></body></FictionBook>')
print("fb2:", repr(bot.extract_document(fb2,".fb2")))
assert "<p>" not in bot.extract_document(fb2,".fb2")
print("PDF/FB2 OK")
