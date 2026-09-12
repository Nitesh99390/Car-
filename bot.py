#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════════════╗
║        📚 NovelTranslator PRO  —  Telegram Document Translation Bot       ║
║                    Professional single-file build  v3.0                   ║
╚══════════════════════════════════════════════════════════════════════════╝

Install
──────────────────────────────────────────────────────────────────────────
    pip install -U pyrogram tgcrypto python-docx beautifulsoup4 ebooklib deep-translator
    pip install -U pypdf reportlab matplotlib python-dotenv     # optional extras
        pypdf      → .pdf input
        reportlab  → .pdf output   (needs a Unicode TTF, see PDF_FONT_PATH)
        matplotlib → /weekly chart
        dotenv     → load .env automatically

Config  (.env  or environment variables — NEVER hardcode secrets)
──────────────────────────────────────────────────────────────────────────
    API_ID=123456
    API_HASH=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    BOT_TOKEN=123456:ABC-DEF...
    SECURITY_CODE=ChangeMe@2026
    OWNER_ID=123456789            # strongly recommended
    BACKUP_GROUP_ID=-1001234567890
    PDF_FONT_PATH=/usr/share/fonts/truetype/noto/NotoSansDevanagari-Regular.ttf

Run
──────────────────────────────────────────────────────────────────────────
    python bot.py

v3.0 — every bug of v2.0 fixed
──────────────────────────────────────────────────────────────────────────
  ✔ Py<3.12 SyntaxError (nested f-string quotes)      ✔ Janitor deleting queued sources
  ✔ Windows-invalid "|" in output names              ✔ Legacy migration hijacking OWNER
  ✔ Queue lost on restart → persisted + resumed      ✔ Paragraph structure preserved (markers)
  ✔ Cancel leaves zombie threads → gathered          ✔ 429 / ban → global cooldown gate
  ✔ Split applied to OUTPUT size, not input          ✔ Real part estimate after extraction
  ✔ Double upload → file_id reuse for backup         ✔ Malformed callback data → validated
  ✔ Security-code brute force → lockout              ✔ Own-output forwarded → explained
  ✔ Partial download left behind → cleaned           ✔ Broadcast FloodWait → retried
  ✔ EPUB dedupe by 300-char hash → full sha1         ✔ TXT file handle leak / dead branch
  ✔ LiveMessage skipped keyboard-only changes        ✔ notify_positions flood → staggered
  ✔ Sync JSON writes on loop → debounced thread      ✔ Owner can cancel anyone's job
  ✔ revoke() leaves user's jobs running → purged     ✔ EPUB "Section N" → real chapters
  ✔ RTL (ur/ar) direction in DOCX/EPUB/HTML          ✔ Graceful shutdown notifies user
  ✔ Per-user limits (pending/queued/daily quota)

v3.0 — new features
──────────────────────────────────────────────────────────────────────────
  TRANSLATION   glossary/name-lock · SQLite translation cache · source language ·
                multi-engine fallback (Google → MyMemory → Libre) · sample preview ·
                bilingual output · junk-line filter · chapter detection (EN/ZH/JA/KO/HI)
  FILES         in: epub txt docx pdf html htm fb2 md   out: txt docx epub pdf html ·
                merge to ZIP · resume checkpoints · re-translate from cache
  QUEUE         fair round-robin · owner priority · per-user caps · daily quota ·
                "your turn" ping · queue ETA · /history · persistent across restarts
  OWNER         /ban /unban /logs /restart /cancelall /priority /weekly (chart) ·
                health ping · forum topics via message_thread_id · /links
  UX            EN + HI interface · quick text translation (just send text) ·
                mobile-friendly menus · custom split
"""

from __future__ import annotations

import asyncio
import hashlib
import html
import io
import json
import logging
import os
import random
import re
import shutil
import signal
import sqlite3
import sys
import threading
import time
import uuid
import zipfile
from collections import defaultdict, deque
from dataclasses import dataclass, field
from logging.handlers import RotatingFileHandler
from typing import Any, Callable, Dict, List, Optional, Tuple

# ── optional .env ──────────────────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import docx
import ebooklib
from bs4 import BeautifulSoup
try:
    import warnings
    from bs4 import XMLParsedAsHTMLWarning
    warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)
except ImportError:                                      # very old bs4
    pass
from deep_translator import GoogleTranslator, MyMemoryTranslator
from ebooklib import epub
from pyrogram import Client, enums, filters
from pyrogram.errors import (FloodWait, MessageIdInvalid, MessageNotModified,
                             RPCError, UserIsBlocked, PeerIdInvalid)
from pyrogram.types import (CallbackQuery, ChatMemberUpdated, InlineKeyboardButton,
                            InlineKeyboardMarkup, Message)

try:
    from deep_translator import LibreTranslator
except ImportError:                                      # pragma: no cover
    LibreTranslator = None
try:
    from deep_translator.exceptions import TooManyRequests, RequestError
except ImportError:                                      # pragma: no cover
    class TooManyRequests(Exception): ...
    class RequestError(Exception): ...

try:
    from pypdf import PdfReader
    HAS_PDF_IN = True
except ImportError:
    HAS_PDF_IN = False
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
    HAS_PDF_OUT = True
except ImportError:
    HAS_PDF_OUT = False
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    HAS_MPL = True
except ImportError:
    HAS_MPL = False

# ═══════════════════════════════════════════════════════════════════════════
# 🔑 CREDENTIALS  — from environment only
# ═══════════════════════════════════════════════════════════════════════════
def _need(name: str) -> str:
    v = os.getenv(name, "").strip()
    if not v:
        sys.exit(f"[config] Missing required environment variable: {name}  (see header of this file)")
    return v


API_ID = int(_need("API_ID"))
API_HASH = _need("API_HASH")
BOT_TOKEN = _need("BOT_TOKEN")
SECURITY_CODE = os.getenv("SECURITY_CODE", "").strip() or uuid.uuid4().hex[:12]
OWNER_ID = int(os.getenv("OWNER_ID", "0") or 0)
BACKUP_GROUP_ID = int(os.getenv("BACKUP_GROUP_ID", "0") or 0)
PDF_FONT_PATH = os.getenv("PDF_FONT_PATH", "").strip()
LIBRE_URL = os.getenv("LIBRE_URL", "https://libretranslate.com").strip()

# ═══════════════════════════════════════════════════════════════════════════
# ⚙️ CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════
BOT_NAME = "NovelTranslator PRO"
VERSION = "3.0.1"
CONCURRENCY_LIMIT = 6              # parallel translation requests (free API — be gentle)
CHUNK_SIZE = 3800                  # chars per request (Google hard limit ≈ 5000)
MAX_RETRIES = 5
MAX_INPUT_MB = 100
PENDING_TTL = 30 * 60              # wizard sessions
EDIT_INTERVAL = 4.0                # dashboard edit throttle
COOLDOWN_429 = 45                  # seconds all workers pause after a 429/ban
MAX_PENDING_PER_USER = 2
MAX_QUEUED_PER_USER = 3
DAILY_QUOTA_CHARS = 3_000_000      # per user per day  (owner unlimited)
BRUTE_MAX_ATTEMPTS = 5
BRUTE_LOCK_SECONDS = 15 * 60
HEALTH_PING_HOURS = 6
QUICK_TEXT_MAX = 3000              # inline text translation limit
PREVIEW_CHARS = 1200
DEFAULT_LANG = "hi"
DEFAULT_SRC = "auto"
DEFAULT_FORMAT = "txt"
DEFAULT_SPLIT_KB = 500
DEFAULT_UI = "en"
OUTPUT_EXPANSION = {               # avg UTF-8 growth factor vs. English source
    "hi": 2.6, "bn": 2.7, "ta": 3.0, "te": 2.8, "mr": 2.6, "gu": 2.6, "kn": 2.9, "ml": 3.1,
    "pa": 2.6, "ur": 1.9, "ne": 2.6, "ar": 1.8, "en": 1.0, "es": 1.1, "fr": 1.1, "id": 1.05,
    "zh-CN": 0.9, "ja": 1.0, "ko": 1.0, "ru": 1.6, "de": 1.1, "pt": 1.1, "tr": 1.1, "vi": 1.2,
}
RTL_LANGS = {"ur", "ar", "fa", "he"}

BASE_DIR = os.path.abspath(os.getenv("WORKDIR", "workspace_textbot"))
STORAGE_DIR = os.path.join(BASE_DIR, "translated_files")
INBOX_DIR = os.path.join(BASE_DIR, "inbox")
CKPT_DIR = os.path.join(BASE_DIR, "checkpoints")
DB_FILE = os.path.join(BASE_DIR, "database.json")
CACHE_FILE = os.path.join(BASE_DIR, "cache.sqlite3")
LEGACY_AUTH_FILE = os.path.join(BASE_DIR, "authorized_users.txt")
LOG_FILE = os.path.join(BASE_DIR, "bot.log")
for _d in (BASE_DIR, STORAGE_DIR, INBOX_DIR, CKPT_DIR):
    os.makedirs(_d, exist_ok=True)

INPUT_EXTS = {".epub", ".txt", ".docx", ".html", ".htm", ".fb2", ".md"} | ({".pdf"} if HAS_PDF_IN else set())
OUTPUT_FORMATS: Dict[str, str] = {"txt": "📄 TXT", "docx": "📝 DOCX", "epub": "📚 EPUB", "html": "🌐 HTML"}
if HAS_PDF_OUT:
    OUTPUT_FORMATS["pdf"] = "🖨 PDF"
SPLIT_PRESETS = [(0, "🚫 No split"), (100, "100 KB"), (300, "300 KB"), (500, "500 KB"),
                 (1024, "1 MB"), (2048, "2 MB"), (5120, "5 MB")]

LANGUAGES: Dict[str, Tuple[str, str]] = {
    "hi": ("Hindi", "🇮🇳"), "en": ("English", "🇬🇧"), "bn": ("Bengali", "🇧🇩"),
    "ta": ("Tamil", "🇮🇳"), "te": ("Telugu", "🇮🇳"), "mr": ("Marathi", "🇮🇳"),
    "gu": ("Gujarati", "🇮🇳"), "kn": ("Kannada", "🇮🇳"), "ml": ("Malayalam", "🇮🇳"),
    "pa": ("Punjabi", "🇮🇳"), "ur": ("Urdu", "🇵🇰"), "ne": ("Nepali", "🇳🇵"),
    "es": ("Spanish", "🇪🇸"), "fr": ("French", "🇫🇷"), "ar": ("Arabic", "🇸🇦"),
    "id": ("Indonesian", "🇮🇩"), "ru": ("Russian", "🇷🇺"), "de": ("German", "🇩🇪"),
    "pt": ("Portuguese", "🇧🇷"), "tr": ("Turkish", "🇹🇷"), "vi": ("Vietnamese", "🇻🇳"),
    "zh-CN": ("Chinese", "🇨🇳"), "ja": ("Japanese", "🇯🇵"), "ko": ("Korean", "🇰🇷"),
}
SOURCE_LANGS: Dict[str, Tuple[str, str]] = {
    "auto": ("Auto detect", "🔍"), "en": ("English", "🇬🇧"), "zh-CN": ("Chinese", "🇨🇳"),
    "ja": ("Japanese", "🇯🇵"), "ko": ("Korean", "🇰🇷"), "hi": ("Hindi", "🇮🇳"),
    "ru": ("Russian", "🇷🇺"), "es": ("Spanish", "🇪🇸"), "fr": ("French", "🇫🇷"),
    "id": ("Indonesian", "🇮🇩"), "th": ("Thai", "🇹🇭"), "vi": ("Vietnamese", "🇻🇳"),
}

# ═══════════════════════════════════════════════════════════════════════════
# 🔧 LOGGING
# ═══════════════════════════════════════════════════════════════════════════
_fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(name)s | %(message)s")
_root = logging.getLogger()
_root.setLevel(logging.INFO)
_console = logging.StreamHandler(sys.stdout); _console.setFormatter(_fmt)
_file = RotatingFileHandler(LOG_FILE, maxBytes=3_000_000, backupCount=3, encoding="utf-8")
_file.setFormatter(_fmt)
_root.handlers = [_console, _file]
logging.getLogger("pyrogram").setLevel(logging.WARNING)
log = logging.getLogger("bot")

# ═══════════════════════════════════════════════════════════════════════════
# 🌍 UI STRINGS  (English / Hindi)
# ═══════════════════════════════════════════════════════════════════════════
STRINGS: Dict[str, Dict[str, str]] = {
    "en": {
        "hello": "👋 Hello, {name}!",
        "intro": "Send me a document — or just paste text —\nand I will translate it into your language.",
        "input": "Input", "output": "Output", "language": "Language", "source": "Source",
        "format": "Format", "split": "Split", "bilingual": "Bilingual", "ui_lang": "Interface",
        "tip_home": "Tip: change defaults in ⚙️ Settings\nand use ⚡ Quick Start next time.",
        "settings": "Settings", "mystats": "My Stats", "queue": "Queue", "help": "Help",
        "history": "History", "glossary": "Glossary", "owner_panel": "Owner Panel", "home": "Home",
        "back": "Back", "cancel": "Cancel", "saved": "✅ Saved", "refresh": "Refresh",
        "locked_title": "Security Locked",
        "locked": "This bot is private.\nSend the <b>security code</b> to unlock.",
        "locked_out": "🚫 Too many wrong attempts.\nTry again in {mins} min.",
        "granted": "Access Granted", "welcome": "Welcome, {name}!\nRole: {role}\n\nSend /start to begin.",
        "banned": "🚫 You are banned from this bot.",
        "unsupported": "Unsupported file", "allowed": "Allowed", "you_sent": "You sent",
        "too_large": "❌ File too large. Limit is {mb} MB.",
        "own_output": "ℹ️ That looks like a file <i>I</i> produced.\nUse ↩️ <b>Re-translate</b> below to translate it again into another language, or send the original book.",
        "too_many_pending": "⚠️ You already have {n} unfinished wizard(s). Finish or /cancel them first.",
        "too_many_queued": "⚠️ You already have {n} job(s) in the queue. Please wait for them.",
        "quota": "⚠️ Daily quota reached ({used} / {limit} chars). Resets in {hrs} h.",
        "downloading": "Downloading", "download_failed": "Download failed", "send_again": "Please send the file again.",
        "step_lang": "Step 1 · Language", "choose_lang": "Choose the target language:",
        "step_fmt": "Step 2 · Format", "choose_fmt": "Choose the output format:",
        "step_split": "Step 3 · Split size", "choose_split": "Max size per delivered part\n<i>(applies to the translated output)</i>:",
        "step_confirm": "Step 4 · Confirm", "looks_good": "Everything looks good?",
        "est_parts": "Estimated parts", "ahead": "Jobs ahead of you", "eta_queue": "Queue ETA",
        "start": "🚀 Start translation", "preview": "👀 Preview sample", "quick": "⚡ Quick Start (use my defaults)",
        "more_opts": "🔧 More options", "custom_size": "✏️ Custom size",
        "custom_prompt": "Type the size per part, e.g.\n<code>750</code> (KB) · <code>2 MB</code> · <code>900kb</code>\n<i>Allowed: 50 KB – 15 MB</i>",
        "invalid_size": "⚠️ Invalid format. Try <code>500</code> or <code>2 MB</code>.",
        "size_range": "⚠️ Size must be between 50 KB and 15 MB.",
        "expired": "Session expired", "expired_body": "Please send the file again.",
        "discarded": "📘 {name} was discarded.",
        "queued": "Queued", "position": "Position", "next_up": "▶️ Next up!", "please_wait": "🕐 Please wait…",
        "your_turn": "🔔 Your turn! Translating {name} now.",
        "processing": "Processing", "extracting": "📖 Extracting text…",
        "failed": "Failed", "no_text": "No readable text was found in this file.\n<i>Scanned PDFs / image-only books are not supported.</i>",
        "dashboard": "Live Dashboard", "part": "Part", "chunks": "Chunks", "speed": "Speed",
        "elapsed": "Elapsed", "eta": "ETA", "cached": "Cached", "retries": "Retries exhausted",
        "translating": "Translating…", "building": "Building .{fmt} file…", "cooling": "⏸ API cooldown ({s}s)…",
        "completed": "Completed", "parts": "Parts", "chars": "Characters", "time": "Time",
        "kept_original": "⚠️ {n} chunk(s) kept original text",
        "thanks": "🚀 Thank you for using {bot}!",
        "cancelled": "Cancelled", "stopped": "Job stopped.", "error": "Error",
        "went_wrong": "Something went wrong while processing.",
        "cancel_job": "❌ Cancel job", "merge_zip": "📦 All parts (ZIP)",
        "no_jobs": "ℹ️ You have no active or queued jobs.", "cancelled_n": "🛑 Cancelled {n} job(s).",
        "send_doc_hint": "📎 Send me a document to translate,\nor paste text (≤ {n} chars) for a quick translation.\n/help for instructions.",
        "quick_result": "🌐 <b>{lang}</b>\n{div}\n{text}",
        "preview_title": "Sample preview", "preview_note": "<i>First ~{n} characters translated. Continue with the full file?</i>",
        "history_empty": "No jobs yet.", "glossary_empty": "Glossary is empty.\nAdd with <code>/glossary Name=अनुवाद</code>",
        "retranslate": "↩️ Re-translate",
    },
    "hi": {
        "hello": "👋 नमस्ते, {name}!",
        "intro": "मुझे कोई डॉक्यूमेंट भेजें — या सीधे टेक्स्ट पेस्ट करें —\nमैं उसे आपकी भाषा में अनुवाद कर दूँगा।",
        "input": "इनपुट", "output": "आउटपुट", "language": "भाषा", "source": "स्रोत भाषा",
        "format": "फ़ॉर्मेट", "split": "विभाजन", "bilingual": "द्विभाषी", "ui_lang": "इंटरफ़ेस",
        "tip_home": "टिप: ⚙️ सेटिंग्स में डिफ़ॉल्ट बदलें\nऔर अगली बार ⚡ क्विक स्टार्ट दबाएँ।",
        "settings": "सेटिंग्स", "mystats": "मेरे आँकड़े", "queue": "कतार", "help": "मदद",
        "history": "इतिहास", "glossary": "शब्दकोश", "owner_panel": "ओनर पैनल", "home": "होम",
        "back": "वापस", "cancel": "रद्द", "saved": "✅ सेव हुआ", "refresh": "रिफ़्रेश",
        "locked_title": "सुरक्षा लॉक",
        "locked": "यह बॉट निजी है।\nअनलॉक करने के लिए <b>सुरक्षा कोड</b> भेजें।",
        "locked_out": "🚫 बहुत ज़्यादा गलत प्रयास।\n{mins} मिनट बाद फिर कोशिश करें।",
        "granted": "प्रवेश मिला", "welcome": "स्वागत है, {name}!\nभूमिका: {role}\n\nशुरू करने के लिए /start भेजें।",
        "banned": "🚫 आपको इस बॉट से प्रतिबंधित किया गया है।",
        "unsupported": "असमर्थित फ़ाइल", "allowed": "अनुमत", "you_sent": "आपने भेजा",
        "too_large": "❌ फ़ाइल बहुत बड़ी है। सीमा {mb} MB है।",
        "own_output": "ℹ️ यह फ़ाइल <i>मैंने</i> ही बनाई लगती है।\nदूसरी भाषा में अनुवाद के लिए नीचे ↩️ <b>पुनः अनुवाद</b> दबाएँ, या मूल किताब भेजें।",
        "too_many_pending": "⚠️ आपके {n} अधूरे विज़ार्ड हैं। पहले उन्हें पूरा या /cancel करें।",
        "too_many_queued": "⚠️ आपके {n} काम पहले से कतार में हैं। कृपया प्रतीक्षा करें।",
        "quota": "⚠️ दैनिक सीमा पूरी ({used} / {limit} अक्षर)। {hrs} घंटे में रीसेट होगी।",
        "downloading": "डाउनलोड हो रहा है", "download_failed": "डाउनलोड विफल", "send_again": "कृपया फ़ाइल फिर से भेजें।",
        "step_lang": "चरण 1 · भाषा", "choose_lang": "लक्ष्य भाषा चुनें:",
        "step_fmt": "चरण 2 · फ़ॉर्मेट", "choose_fmt": "आउटपुट फ़ॉर्मेट चुनें:",
        "step_split": "चरण 3 · भाग का आकार", "choose_split": "प्रति भाग अधिकतम आकार\n<i>(अनुवादित आउटपुट पर लागू)</i>:",
        "step_confirm": "चरण 4 · पुष्टि", "looks_good": "सब ठीक है?",
        "est_parts": "अनुमानित भाग", "ahead": "आपसे आगे काम", "eta_queue": "कतार ETA",
        "start": "🚀 अनुवाद शुरू करें", "preview": "👀 नमूना देखें", "quick": "⚡ क्विक स्टार्ट (मेरे डिफ़ॉल्ट)",
        "more_opts": "🔧 और विकल्प", "custom_size": "✏️ कस्टम आकार",
        "custom_prompt": "प्रति भाग आकार लिखें, जैसे\n<code>750</code> (KB) · <code>2 MB</code> · <code>900kb</code>\n<i>अनुमत: 50 KB – 15 MB</i>",
        "invalid_size": "⚠️ गलत फ़ॉर्मेट। <code>500</code> या <code>2 MB</code> लिखें।",
        "size_range": "⚠️ आकार 50 KB से 15 MB के बीच होना चाहिए।",
        "expired": "सत्र समाप्त", "expired_body": "कृपया फ़ाइल फिर से भेजें।",
        "discarded": "📘 {name} हटा दी गई।",
        "queued": "कतार में", "position": "स्थान", "next_up": "▶️ अगला आपका!", "please_wait": "🕐 कृपया प्रतीक्षा करें…",
        "your_turn": "🔔 आपकी बारी! {name} का अनुवाद शुरू।",
        "processing": "प्रोसेसिंग", "extracting": "📖 टेक्स्ट निकाला जा रहा है…",
        "failed": "विफल", "no_text": "इस फ़ाइल में पढ़ने योग्य टेक्स्ट नहीं मिला।\n<i>स्कैन की गई PDF / केवल-छवि किताबें समर्थित नहीं हैं।</i>",
        "dashboard": "लाइव डैशबोर्ड", "part": "भाग", "chunks": "टुकड़े", "speed": "गति",
        "elapsed": "बीता", "eta": "शेष", "cached": "कैश से", "retries": "पुनः प्रयास विफल",
        "translating": "अनुवाद हो रहा है…", "building": ".{fmt} फ़ाइल बन रही है…", "cooling": "⏸ API विराम ({s}s)…",
        "completed": "पूर्ण", "parts": "भाग", "chars": "अक्षर", "time": "समय",
        "kept_original": "⚠️ {n} टुकड़े मूल टेक्स्ट में रहे",
        "thanks": "🚀 {bot} इस्तेमाल करने के लिए धन्यवाद!",
        "cancelled": "रद्द", "stopped": "काम रोक दिया गया।", "error": "त्रुटि",
        "went_wrong": "प्रोसेसिंग में कुछ गलत हो गया।",
        "cancel_job": "❌ काम रद्द करें", "merge_zip": "📦 सभी भाग (ZIP)",
        "no_jobs": "ℹ️ आपका कोई सक्रिय या कतारबद्ध काम नहीं है।", "cancelled_n": "🛑 {n} काम रद्द किए।",
        "send_doc_hint": "📎 अनुवाद के लिए डॉक्यूमेंट भेजें,\nया त्वरित अनुवाद के लिए टेक्स्ट (≤ {n} अक्षर) पेस्ट करें।\nनिर्देश: /help",
        "quick_result": "🌐 <b>{lang}</b>\n{div}\n{text}",
        "preview_title": "नमूना पूर्वावलोकन", "preview_note": "<i>पहले ~{n} अक्षर अनुवादित। पूरी फ़ाइल के साथ जारी रखें?</i>",
        "history_empty": "अभी कोई काम नहीं।", "glossary_empty": "शब्दकोश खाली है।\nजोड़ें: <code>/glossary Name=अनुवाद</code>",
        "retranslate": "↩️ पुनः अनुवाद",
    },
}


def T(uid: int, key: str, **kw) -> str:
    """Translate a UI string for a user (falls back to English)."""
    ui = store.pref(uid, "ui", DEFAULT_UI) if uid else DEFAULT_UI
    s = STRINGS.get(ui, STRINGS["en"]).get(key) or STRINGS["en"].get(key, key)
    return s.format(**kw) if kw else s


# ═══════════════════════════════════════════════════════════════════════════
# 🗄️ PERSISTENT STORE  (JSON, debounced writes off the event loop)
# ═══════════════════════════════════════════════════════════════════════════
class Store:
    def __init__(self, path: str):
        self.path = path
        self._lock = threading.Lock()
        self._dirty = False
        self._timer: Optional[threading.Timer] = None
        self.data: Dict[str, Any] = {
            "owner_id": 0,
            "security_code": None,
            "users": {},       # uid -> {...}
            "banned": {},      # uid -> {reason, at}
            "chats": {},       # chat_id -> {title, type, added}
            "stats": {"jobs": 0, "parts": 0, "chars": 0, "failed": 0, "cache_hits": 0},
            "daily": {},       # "YYYY-MM-DD" -> {"jobs":n,"chars":n,"failed":n}
            "queue": [],       # persisted queued jobs (dicts)
        }
        self._load()
        self._migrate_legacy()

    # -- io ----------------------------------------------------------------
    def _load(self):
        if os.path.exists(self.path):
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                for k, v in saved.items():
                    self.data[k] = v
                # forward-compat defaults
                self.data["stats"].setdefault("cache_hits", 0)
                for k in ("banned", "daily", "chats", "users"):
                    self.data.setdefault(k, {})
                self.data.setdefault("queue", [])
            except Exception as e:
                log.error("DB load failed: %s — starting with backup copy", e)
                try:
                    shutil.copy(self.path, self.path + ".corrupt")
                except Exception:
                    pass

    def save(self, immediate: bool = False):
        """Debounced: coalesce writes within 1.5 s, run on a thread."""
        with self._lock:
            self._dirty = True
            if immediate:
                self._flush()
                return
            if self._timer is None:
                self._timer = threading.Timer(1.5, self._flush_from_timer)
                self._timer.daemon = True
                self._timer.start()

    def _flush_from_timer(self):
        with self._lock:
            self._flush()

    def _flush(self):
        # caller holds lock
        self._timer = None
        if not self._dirty:
            return
        try:
            payload = json.dumps(self.data, ensure_ascii=False, indent=1)
            tmp = self.path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                f.write(payload)
            os.replace(tmp, self.path)
            self._dirty = False
        except Exception as e:                      # e.g. dict mutated mid-dump → retry shortly
            log.error("DB save failed (%s) — retrying", e)
            self._timer = threading.Timer(2.0, self._flush_from_timer)
            self._timer.daemon = True
            self._timer.start()

    def _migrate_legacy(self):
        """Import old authorized_users.txt WITHOUT touching owner assignment (bug #4)."""
        if not os.path.exists(LEGACY_AUTH_FILE):
            return
        added = 0
        try:
            with open(LEGACY_AUTH_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    s = line.strip()
                    if s.isdigit() and s not in self.data["users"]:
                        self.data["users"][s] = self._default_user("Legacy User")
                        added += 1
            os.rename(LEGACY_AUTH_FILE, LEGACY_AUTH_FILE + ".migrated")
            self.save(immediate=True)
            log.info("Migrated %d users from legacy auth file", added)
        except Exception as e:
            log.error("Legacy migration failed: %s", e)

    # -- users -------------------------------------------------------------
    @staticmethod
    def _default_user(name: str) -> dict:
        return {"name": name, "joined": int(time.time()), "role": "user",
                "lang": DEFAULT_LANG, "src": DEFAULT_SRC, "fmt": DEFAULT_FORMAT,
                "split": DEFAULT_SPLIT_KB, "bilingual": False, "ui": DEFAULT_UI,
                "glossary": {}, "history": [],
                "stats": {"jobs": 0, "parts": 0, "chars": 0},
                "quota": {"day": "", "chars": 0}}

    def user(self, uid: int) -> Optional[dict]:
        return self.data["users"].get(str(uid))

    def is_authorized(self, uid: int) -> bool:
        return str(uid) in self.data["users"] and not self.is_banned(uid)

    def authorize(self, uid: int, name: str):
        u = self.data["users"].get(str(uid))
        if u is None:
            self.data["users"][str(uid)] = u = self._default_user(name)
        else:
            u["name"] = name
            for k, v in self._default_user(name).items():      # forward-compat
                u.setdefault(k, v)
        if not self.owner_id():
            self.data["owner_id"] = uid
            u["role"] = "owner"
        self.save()

    def revoke(self, uid: int) -> bool:
        if str(uid) in self.data["users"] and uid != self.owner_id():
            del self.data["users"][str(uid)]
            self.save()
            return True
        return False

    def owner_id(self) -> int:
        return OWNER_ID or int(self.data.get("owner_id") or 0)

    def is_owner(self, uid: int) -> bool:
        return uid == self.owner_id()

    def pref(self, uid: int, key: str, default=None):
        u = self.user(uid)
        return u.get(key, default) if u else default

    def set_pref(self, uid: int, key: str, value):
        u = self.user(uid)
        if u:
            u[key] = value
            self.save()

    # -- bans --------------------------------------------------------------
    def is_banned(self, uid: int) -> bool:
        return str(uid) in self.data["banned"]

    def ban(self, uid: int, reason: str = ""):
        self.data["banned"][str(uid)] = {"reason": reason, "at": int(time.time())}
        self.save()

    def unban(self, uid: int) -> bool:
        if str(uid) in self.data["banned"]:
            del self.data["banned"][str(uid)]
            self.save()
            return True
        return False

    # -- quota -------------------------------------------------------------
    @staticmethod
    def _today() -> str:
        return time.strftime("%Y-%m-%d", time.gmtime())

    def quota_used(self, uid: int) -> int:
        u = self.user(uid)
        if not u:
            return 0
        q = u.setdefault("quota", {"day": "", "chars": 0})
        if q.get("day") != self._today():
            q["day"], q["chars"] = self._today(), 0
        return int(q["chars"])

    def quota_left(self, uid: int) -> int:
        if self.is_owner(uid):
            return 10 ** 12
        return max(0, DAILY_QUOTA_CHARS - self.quota_used(uid))

    def quota_add(self, uid: int, chars: int):
        u = self.user(uid)
        if u and not self.is_owner(uid):
            self.quota_used(uid)
            u["quota"]["chars"] += chars

    # -- stats / history ---------------------------------------------------
    def bump(self, uid: int, parts: int, chars: int, failed=False, cache_hits=0):
        u = self.user(uid)
        if u:
            u["stats"]["jobs"] += 1
            u["stats"]["parts"] += parts
            u["stats"]["chars"] += chars
        s = self.data["stats"]
        s["jobs"] += 1; s["parts"] += parts; s["chars"] += chars; s["cache_hits"] += cache_hits
        d = self.data["daily"].setdefault(self._today(), {"jobs": 0, "chars": 0, "failed": 0})
        d["jobs"] += 1; d["chars"] += chars
        if failed:
            s["failed"] += 1; d["failed"] += 1
        # keep 60 days
        for k in sorted(self.data["daily"])[:-60]:
            del self.data["daily"][k]
        self.save()

    def add_history(self, uid: int, entry: dict):
        u = self.user(uid)
        if u:
            h = u.setdefault("history", [])
            h.insert(0, entry)
            del h[15:]
            self.save()

    # -- glossary ----------------------------------------------------------
    def glossary(self, uid: int) -> Dict[str, str]:
        return dict(self.pref(uid, "glossary", {}) or {})

    def glossary_set(self, uid: int, src: str, dst: str):
        u = self.user(uid)
        if u:
            u.setdefault("glossary", {})[src] = dst
            self.save()

    def glossary_del(self, uid: int, src: str) -> bool:
        u = self.user(uid)
        if u and src in u.get("glossary", {}):
            del u["glossary"][src]
            self.save()
            return True
        return False

    # -- security ----------------------------------------------------------
    def security_code(self) -> str:
        return self.data.get("security_code") or SECURITY_CODE

    def set_security_code(self, code: str):
        self.data["security_code"] = code
        self.save(immediate=True)

    # -- chats -------------------------------------------------------------
    def track_chat(self, chat_id: int, title: str, ctype: str):
        self.data["chats"][str(chat_id)] = {"title": title, "type": ctype, "added": int(time.time())}
        self.save()

    def untrack_chat(self, chat_id: int):
        if self.data["chats"].pop(str(chat_id), None) is not None:
            self.save()

    # -- queue persistence -------------------------------------------------
    def save_queue(self, jobs: List[dict]):
        self.data["queue"] = jobs
        self.save(immediate=True)


store = Store(DB_FILE)


# ═══════════════════════════════════════════════════════════════════════════
# 💾 TRANSLATION CACHE  (SQLite, thread-safe)
# ═══════════════════════════════════════════════════════════════════════════
class TranslationCache:
    def __init__(self, path: str):
        self._lock = threading.Lock()
        self.conn = sqlite3.connect(path, check_same_thread=False)
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("""CREATE TABLE IF NOT EXISTS tr (
            k TEXT PRIMARY KEY, v TEXT NOT NULL, ts INTEGER NOT NULL)""")
        self.conn.execute("""CREATE TABLE IF NOT EXISTS docs (
            k TEXT PRIMARY KEY, name TEXT, text TEXT NOT NULL, ts INTEGER NOT NULL)""")
        self.conn.commit()

    @staticmethod
    def key(text: str, src: str, tgt: str) -> str:
        return hashlib.sha1(f"{src}|{tgt}|{text}".encode("utf-8")).hexdigest()

    def get(self, text: str, src: str, tgt: str) -> Optional[str]:
        with self._lock:
            row = self.conn.execute("SELECT v FROM tr WHERE k=?", (self.key(text, src, tgt),)).fetchone()
        return row[0] if row else None

    def put(self, text: str, src: str, tgt: str, result: str):
        with self._lock:
            self.conn.execute("INSERT OR REPLACE INTO tr VALUES (?,?,?)",
                              (self.key(text, src, tgt), result, int(time.time())))
            self.conn.commit()

    # extracted-document cache (for re-translate without re-upload)
    def put_doc(self, file_uid: str, name: str, text: str):
        with self._lock:
            self.conn.execute("INSERT OR REPLACE INTO docs VALUES (?,?,?,?)",
                              (file_uid, name, text, int(time.time())))
            self.conn.commit()

    def get_doc(self, file_uid: str) -> Optional[Tuple[str, str]]:
        with self._lock:
            row = self.conn.execute("SELECT name, text FROM docs WHERE k=?", (file_uid,)).fetchone()
        return (row[0], row[1]) if row else None

    def stats(self) -> Tuple[int, int]:
        with self._lock:
            a = self.conn.execute("SELECT COUNT(*) FROM tr").fetchone()[0]
            b = self.conn.execute("SELECT COUNT(*) FROM docs").fetchone()[0]
        return a, b

    def vacuum_old(self, days: int = 90):
        cutoff = int(time.time()) - days * 86400
        with self._lock:
            self.conn.execute("DELETE FROM tr WHERE ts<?", (cutoff,))
            self.conn.execute("DELETE FROM docs WHERE ts<?", (cutoff,))
            self.conn.commit()


cache = TranslationCache(CACHE_FILE)

# ═══════════════════════════════════════════════════════════════════════════
# 🎨 UI HELPERS
# ═══════════════════════════════════════════════════════════════════════════
DIV = "━━━━━━━━━━━━━━━━━━━━"


def esc(s) -> str:
    return html.escape(str(s if s is not None else ""), quote=False)


def b(s) -> str:
    return f"<b>{esc(s)}</b>"


def code(s) -> str:
    return f"<code>{esc(s)}</code>"


def header(title: str, icon: str = "📚") -> str:
    return f"{icon} <b>{esc(title)}</b>\n{DIV}\n"


def progress_bar(ratio: float, width: int = 12) -> str:
    ratio = max(0.0, min(1.0, ratio))
    filled = int(round(ratio * width))
    return "▰" * filled + "▱" * (width - filled)


def fmt_time(seconds: float) -> str:
    seconds = int(max(0, seconds))
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:d}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


def fmt_size(num: float) -> str:
    num = float(num or 0)
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024 or unit == "GB":
            return f"{num:.0f} {unit}" if unit == "B" else f"{num:.1f} {unit}"
        num /= 1024
    return f"{num:.1f} GB"


def fmt_int(n) -> str:
    return f"{int(n or 0):,}"


def lang_label(code_: str) -> str:
    name, flag = LANGUAGES.get(code_, SOURCE_LANGS.get(code_, (code_.upper(), "🌐")))
    return f"{flag} {name}"


def lang_name(code_: str) -> str:
    return LANGUAGES.get(code_, SOURCE_LANGS.get(code_, (code_, "")))[0]


def split_label(kb: int) -> str:
    if kb == 0:
        return "No split"
    return f"{kb / 1024:g} MB" if kb >= 1024 else f"{kb} KB"


def safe_filename(s: str, limit: int = 90) -> str:
    """Strip everything invalid on Windows / Linux / macOS (bug #3)."""
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", s).strip(" .")
    s = re.sub(r"\s+", " ", s)
    return (s or "Document")[:limit]


def kb_rows(buttons: List[InlineKeyboardButton], per_row: int = 2) -> List[List[InlineKeyboardButton]]:
    return [buttons[i:i + per_row] for i in range(0, len(buttons), per_row)]


def btn(text: str, cb: str) -> InlineKeyboardButton:
    return InlineKeyboardButton(text, callback_data=cb)


def home_keyboard(uid: int) -> InlineKeyboardMarkup:
    rows = [
        [btn(f"⚙️ {T(uid, 'settings')}", "nav:settings"), btn(f"📊 {T(uid, 'mystats')}", "nav:mystats")],
        [btn(f"📋 {T(uid, 'queue')}", "nav:queue"), btn(f"🕘 {T(uid, 'history')}", "nav:history")],
        [btn(f"📖 {T(uid, 'glossary')}", "nav:glossary"), btn(f"ℹ️ {T(uid, 'help')}", "nav:help")],
    ]
    if store.is_owner(uid):
        rows.append([btn(f"👑 {T(uid, 'owner_panel')}", "nav:owner")])
    return InlineKeyboardMarkup(rows)


def back_home_kb(uid: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[btn(f"🏠 {T(uid, 'home')}", "nav:home")]])


def text_home(uid: int, name: str) -> str:
    u = store.user(uid) or {}
    return (
        header(BOT_NAME) +
        T(uid, "hello", name=b(name)) + "\n\n" + T(uid, "intro") + "\n\n"
        f"📥 {b(T(uid, 'input'))}: {', '.join(sorted(e.lstrip('.').upper() for e in INPUT_EXTS))}\n"
        f"📤 {b(T(uid, 'output'))}: {' · '.join(f.upper() for f in OUTPUT_FORMATS)}\n\n"
        f"{DIV}\n"
        f"🌐 {T(uid, 'language')}: {b(lang_label(u.get('lang', DEFAULT_LANG)))}\n"
        f"🔤 {T(uid, 'source')}: {b(lang_label(u.get('src', DEFAULT_SRC)))}\n"
        f"📄 {T(uid, 'format')}: {b(u.get('fmt', DEFAULT_FORMAT).upper())}\n"
        f"✂️ {T(uid, 'split')}: {b(split_label(u.get('split', DEFAULT_SPLIT_KB)))}\n"
        f"🔁 {T(uid, 'bilingual')}: {b('ON' if u.get('bilingual') else 'OFF')}\n"
        f"{DIV}\n<i>{T(uid, 'tip_home')}</i>"
    )


def text_help(uid: int) -> str:
    return (
        header(T(uid, "help"), "ℹ️") +
        "1️⃣ Send a book (.epub / .txt / .docx / .pdf / .html / .fb2 / .md)\n"
        "2️⃣ Choose language, format & split\n"
        "3️⃣ Watch the live dashboard\n"
        "4️⃣ Receive translated parts (+ ZIP)\n\n"
        f"{b('Commands')}\n"
        "/start · /settings · /queue · /cancel\n"
        "/mystats · /history · /help\n"
        "/glossary  – list terms\n"
        "/glossary Li Wei=ली वेई  – add term\n"
        "/glossary -Li Wei  – remove term\n"
        "/tr &lt;lang&gt; &lt;text&gt;  – quick translate\n\n"
        f"{b('Split size')}\n"
        "Applies to the <i>translated</i> output.\n"
        "Custom: 50 KB – 15 MB, e.g. <code>750</code>, <code>2 MB</code>\n\n"
        f"{b('Quick text')}\n"
        f"Just paste up to {QUICK_TEXT_MAX} characters — no file needed."
    )


def text_settings(uid: int) -> str:
    u = store.user(uid) or {}
    return (
        header(T(uid, "settings"), "⚙️") +
        f"🌐 {T(uid, 'language')}: {b(lang_label(u.get('lang', DEFAULT_LANG)))}\n"
        f"🔤 {T(uid, 'source')}: {b(lang_label(u.get('src', DEFAULT_SRC)))}\n"
        f"📄 {T(uid, 'format')}: {b(u.get('fmt', DEFAULT_FORMAT).upper())}\n"
        f"✂️ {T(uid, 'split')}: {b(split_label(u.get('split', DEFAULT_SPLIT_KB)))}\n"
        f"🔁 {T(uid, 'bilingual')}: {b('ON' if u.get('bilingual') else 'OFF')}\n"
        f"🗣 {T(uid, 'ui_lang')}: {b('हिन्दी' if u.get('ui') == 'hi' else 'English')}\n"
        f"{DIV}\n<i>Used by ⚡ Quick Start and pre-selected in the wizard.</i>"
    )


def settings_keyboard(uid: int) -> InlineKeyboardMarkup:
    u = store.user(uid) or {}
    return InlineKeyboardMarkup([
        [btn(f"🌐 {T(uid, 'language')}", "st:lang"), btn(f"🔤 {T(uid, 'source')}", "st:src")],
        [btn(f"📄 {T(uid, 'format')}", "st:fmt"), btn(f"✂️ {T(uid, 'split')}", "st:spl")],
        [btn(f"🔁 {T(uid, 'bilingual')}: {'ON' if u.get('bilingual') else 'OFF'}", "st:bil"),
         btn("🗣 English / हिन्दी", "st:ui")],
        [btn(f"🏠 {T(uid, 'home')}", "nav:home")],
    ])


def language_keyboard(prefix: str, selected: str, back_cb: str, table: Dict[str, Tuple[str, str]] = LANGUAGES,
                      back_label: str = "◀️ Back") -> InlineKeyboardMarkup:
    btns = [btn(f"{'✅ ' if c == selected else ''}{flag} {name}", f"{prefix}{c}") for c, (name, flag) in table.items()]
    rows = kb_rows(btns, 2)
    rows.append([btn(back_label, back_cb)])
    return InlineKeyboardMarkup(rows)


def format_keyboard(prefix: str, selected: str, back_cb: str, back_label: str = "◀️ Back") -> InlineKeyboardMarkup:
    btns = [btn(f"{'✅ ' if f == selected else ''}{label}", f"{prefix}{f}") for f, label in OUTPUT_FORMATS.items()]
    return InlineKeyboardMarkup(kb_rows(btns, 3) + [[btn(back_label, back_cb)]])


def split_keyboard(prefix: str, selected: int, back_cb: str, custom_cb: Optional[str],
                   custom_label: str = "✏️ Custom size", back_label: str = "◀️ Back") -> InlineKeyboardMarkup:
    btns = [btn(f"{'✅ ' if kb == selected else ''}{label}", f"{prefix}{kb}") for kb, label in SPLIT_PRESETS]
    rows = kb_rows(btns, 2)
    if custom_cb:
        rows.append([btn(custom_label, custom_cb)])
    rows.append([btn(back_label, back_cb)])
    return InlineKeyboardMarkup(rows)


# ═══════════════════════════════════════════════════════════════════════════
# 📬 FLOOD-SAFE MESSAGE UTILITIES
# ═══════════════════════════════════════════════════════════════════════════
def _kb_sig(kb: Optional[InlineKeyboardMarkup]) -> str:
    if not kb:
        return ""
    return "|".join(f"{x.text}:{x.callback_data}" for row in kb.inline_keyboard for x in row)


class LiveMessage:
    """Edits one Telegram message with throttling + full error tolerance (bug #19 fixed)."""

    def __init__(self, message: Optional[Message], min_interval: float = EDIT_INTERVAL):
        self.msg = message
        self.min_interval = min_interval
        self._last_edit = 0.0
        self._last_sig = ""
        self.alive = message is not None

    async def update(self, text: str, keyboard: Optional[InlineKeyboardMarkup] = None, force: bool = False):
        if not self.alive:
            return
        now = time.time()
        sig = text + "\x00" + _kb_sig(keyboard)
        if not force and (now - self._last_edit) < self.min_interval:
            return
        if sig == self._last_sig:
            return
        try:
            await self.msg.edit_text(text, reply_markup=keyboard, disable_web_page_preview=True)
            self._last_edit = time.time()
            self._last_sig = sig
        except MessageNotModified:
            self._last_edit = time.time(); self._last_sig = sig
        except FloodWait as e:
            self._last_edit = time.time() + float(e.value)
        except MessageIdInvalid:
            self.alive = False
        except RPCError as e:
            log.warning("edit failed: %s", e)
        except Exception as e:                                  # pragma: no cover
            log.warning("edit failed: %s", e)


async def safe_send(chat_id: int, text: str, **kwargs) -> Optional[Message]:
    if not chat_id:
        return None
    for _ in range(3):
        try:
            return await app.send_message(chat_id, text, disable_web_page_preview=True, **kwargs)
        except FloodWait as e:
            await asyncio.sleep(float(e.value) + 1)
        except (UserIsBlocked, PeerIdInvalid):
            return None
        except Exception as e:
            log.warning("send_message failed: %s", e)
            return None
    return None


async def safe_send_document(chat_id: int, document, caption: str, **kwargs) -> Optional[Message]:
    """`document` may be a path or a Telegram file_id (bug #11: reuse for backup)."""
    if not chat_id:
        return None
    for attempt in range(4):
        try:
            return await app.send_document(chat_id, document, caption=caption, **kwargs)
        except FloodWait as e:
            await asyncio.sleep(float(e.value) + 2)
        except (UserIsBlocked, PeerIdInvalid):
            return None
        except Exception as e:
            log.warning("send_document failed (%s/4): %s", attempt + 1, e)
            await asyncio.sleep(2 + attempt * 2)
    return None


async def create_backup_topic(title: str) -> Optional[int]:
    """Create a forum topic in the backup group; returns thread id or None."""
    if not BACKUP_GROUP_ID:
        return None
    try:
        topic = await app.create_forum_topic(BACKUP_GROUP_ID, title[:120])       # pyrogram ≥2.1
        return topic.id
    except AttributeError:
        pass
    except Exception as e:
        log.debug("create_forum_topic: %s", e)
    try:                                                                          # raw fallback
        from pyrogram.raw import functions
        peer = await app.resolve_peer(BACKUP_GROUP_ID)
        result = await app.invoke(functions.channels.CreateForumTopic(
            channel=peer, title=title[:120], random_id=random.randint(1, 2 ** 62)))
        for upd in result.updates:
            msg = getattr(upd, "message", None)
            if msg is not None and hasattr(msg, "id"):
                return msg.id
    except Exception as e:
        log.warning("Backup topic creation failed: %s", e)
    return None


def thread_kwargs(topic_id: Optional[int]) -> dict:
    return {"message_thread_id": topic_id} if topic_id else {}


async def backup_send(text: str, topic_id: Optional[int]):
    if not BACKUP_GROUP_ID:
        return
    try:
        await app.send_message(BACKUP_GROUP_ID, text, disable_web_page_preview=True, **thread_kwargs(topic_id))
    except TypeError:                                  # old pyrogram without message_thread_id
        kw = {"reply_to_message_id": topic_id} if topic_id else {}
        await safe_send(BACKUP_GROUP_ID, text, **kw)
    except FloodWait as e:
        await asyncio.sleep(float(e.value) + 1)
    except Exception as e:
        log.debug("backup_send: %s", e)


async def backup_doc(document, caption: str, topic_id: Optional[int]):
    if not BACKUP_GROUP_ID:
        return
    try:
        await app.send_document(BACKUP_GROUP_ID, document, caption=caption, **thread_kwargs(topic_id))
    except TypeError:
        await safe_send_document(BACKUP_GROUP_ID, document, caption,
                                 **({"reply_to_message_id": topic_id} if topic_id else {}))
    except FloodWait as e:
        await asyncio.sleep(float(e.value) + 2)
        await safe_send_document(BACKUP_GROUP_ID, document, caption)
    except Exception as e:
        log.debug("backup_doc: %s", e)


# ═══════════════════════════════════════════════════════════════════════════
# 📖 TEXT EXTRACTION  (returns list of (heading|None, paragraph) blocks)
# ═══════════════════════════════════════════════════════════════════════════
JUNK_RE = re.compile(
    r"(https?://\S+|www\.\S+\.\S+|read\s+(more|latest|free)\s+(chapters?\s+)?(at|on)\b|"
    r"visit\s+\S+\.(com|net|org|xyz|me|info)|\bnovel(full|bin|updates)\b|\bwuxiaworld\b|"
    r"translated\s+by\s+\S+\s+at\b|please\s+support\s+(us|the\s+author)\s+(at|on)\b|"
    r"如需最新章节|请访问|www\.[a-z0-9]+\.(com|net))", re.IGNORECASE)

CHAPTER_RE = re.compile(
    r"^\s*(?:"
    r"(?:chapter|ch\.?|episode|ep\.?|book|part|volume|vol\.?)\s*(?:[\d一二三四五六七八九十百千零〇]+|[IVXLCDM]+\b)[\s:.\-—–]*.{0,80}"
    r"|(?:prologue|epilogue|interlude|afterword|foreword|preface)\b.{0,60}"
    r"|第\s*[\d一二三四五六七八九十百千零〇]+\s*[章节回卷部話].{0,60}"
    r"|\d+\s*화.{0,60}"
    r"|(?:अध्याय|भाग|प्रकरण)\s*[\d०-९]+.{0,60}"
    r"|\d{1,4}\s*[.:\-—]\s+\S.{0,50}"
    r")\s*$", re.IGNORECASE)

Block = Tuple[Optional[int], str]      # (heading level or None, text)


def is_junk(line: str) -> bool:
    return bool(JUNK_RE.search(line)) and len(line) < 220


def is_heading(line: str) -> bool:
    return len(line) <= 90 and bool(CHAPTER_RE.match(line))


def blocks_from_plain(text: str) -> List[Block]:
    """Turn plain text into blocks, detecting chapter headings and filtering junk."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    blocks: List[Block] = []
    for raw in text.split("\n"):
        line = raw.strip()
        if not line or is_junk(line):
            continue
        blocks.append((2 if is_heading(line) else None, line))
    return blocks


def _html_to_blocks(markup) -> List[Block]:
    soup = BeautifulSoup(markup, "html.parser")
    for junk in soup(["script", "style", "nav", "header", "footer", "aside"]):
        junk.decompose()
    block_tags = ["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "blockquote", "div", "td", "pre"]
    blocks: List[Block] = []
    for el in soup.find_all(block_tags):
        if el.find(block_tags):
            continue
        txt = el.get_text(" ", strip=True)
        if not txt or is_junk(txt):
            continue
        if el.name.startswith("h") and len(el.name) == 2:
            blocks.append((min(int(el.name[1]), 3), txt))
        elif is_heading(txt):
            blocks.append((2, txt))
        else:
            blocks.append((None, txt))
    if not blocks:
        blocks = blocks_from_plain(soup.get_text("\n", strip=True))
    return blocks


def extract_epub(path: str) -> List[Block]:
    try:
        try:
            book = epub.read_epub(path, options={"ignore_ncx": True})
        except TypeError:
            book = epub.read_epub(path)
    except Exception as e:
        log.error("EPUB read failed: %s", e)
        return []
    items = []
    for entry in getattr(book, "spine", []):
        idref = entry[0] if isinstance(entry, (tuple, list)) else entry
        item = book.get_item_with_id(idref)
        if item is not None and item.get_type() == ebooklib.ITEM_DOCUMENT:
            items.append(item)
    if not items:
        items = [i for i in book.get_items() if i.get_type() == ebooklib.ITEM_DOCUMENT]
    out: List[Block] = []
    seen = set()
    for item in items:
        try:
            blocks = _html_to_blocks(item.get_body_content())
        except Exception:
            continue
        joined = "\n".join(t for _, t in blocks)
        if len(joined) < 30:
            continue
        key = hashlib.sha1(joined.encode("utf-8")).hexdigest()       # bug: 300-char hash dedupe
        if key in seen:
            continue
        seen.add(key)
        out.extend(blocks)
    return out


def extract_txt(path: str) -> List[Block]:
    with open(path, "rb") as f:
        raw = f.read()
    text = None
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):                 # UTF-16 BOM
        try:
            text = raw.decode("utf-16")
        except UnicodeDecodeError:
            text = None
    if text is None:
        for enc in ("utf-8-sig", "gb18030", "cp1252"):
            try:
                text = raw.decode(enc)
                break
            except (UnicodeDecodeError, LookupError):
                continue
    if text is None:
        text = raw.decode("utf-8", errors="replace")
    return blocks_from_plain(text)


def extract_md(path: str) -> List[Block]:
    blocks: List[Block] = []
    for lvl, txt in extract_txt(path):
        m = re.match(r"^(#{1,6})\s+(.*)$", txt)
        if m:
            blocks.append((min(len(m.group(1)), 3), m.group(2).strip()))
        else:
            txt = re.sub(r"[*_`>]+", "", txt)
            blocks.append((lvl, txt))
    return blocks


def extract_docx(path: str) -> List[Block]:
    try:
        d = docx.Document(path)
    except Exception as e:
        log.error("DOCX read failed: %s", e)
        return []
    blocks: List[Block] = []
    for p in d.paragraphs:
        txt = p.text.strip()
        if not txt or is_junk(txt):
            continue
        try:
            style = (p.style.name or "").lower() if p.style is not None else ""
        except Exception:
            style = ""
        m = re.match(r"heading\s*(\d)", style)
        if m or style == "title":
            blocks.append((min(int(m.group(1)) if m else 1, 3), txt))
        else:
            blocks.append((2 if is_heading(txt) else None, txt))
    return blocks


def extract_pdf(path: str) -> List[Block]:
    if not HAS_PDF_IN:
        return []
    try:
        reader = PdfReader(path)
        pages = []
        for page in reader.pages:
            t = page.extract_text() or ""
            if t.strip():
                pages.append(t)
        # join hyphenated / wrapped lines into paragraphs
        text = "\n".join(pages)
        text = re.sub(r"-\n(?=[a-z])", "", text)
        text = re.sub(r"(?<![.!?。！？:\"”])\n(?=[a-z,;])", " ", text)
        return blocks_from_plain(text)
    except Exception as e:
        log.error("PDF read failed: %s", e)
        return []


def extract_html(path: str) -> List[Block]:
    with open(path, "rb") as f:
        return _html_to_blocks(f.read())


def _xml_soup(data: bytes) -> BeautifulSoup:
    """Parse real XML (FB2) with an XML parser when lxml is present.
    html.parser treats <title> as raw text, so <title><p>Ch 1</p></title> would
    leak literal tags into the output."""
    try:
        return BeautifulSoup(data, "xml")
    except Exception:                                  # lxml missing → best effort
        return BeautifulSoup(data, "html.parser")


def extract_fb2(path: str) -> List[Block]:
    with open(path, "rb") as f:
        soup = _xml_soup(f.read())
    blocks: List[Block] = []
    body = soup.find("body") or soup
    for el in body.find_all(["title", "subtitle", "p", "v"]):
        if el.name in ("p", "v") and el.find_parent(["title", "subtitle"]):
            continue                                   # already emitted via the heading
        txt = el.get_text(" ", strip=True)
        if not txt or is_junk(txt):
            continue
        blocks.append((2 if el.name in ("title", "subtitle") else None, txt))
    return blocks


EXTRACTORS: Dict[str, Callable[[str], List[Block]]] = {
    ".epub": extract_epub, ".txt": extract_txt, ".md": extract_md, ".docx": extract_docx,
    ".pdf": extract_pdf, ".html": extract_html, ".htm": extract_html, ".fb2": extract_fb2,
}

# Serialised block format: heading lines prefixed with a private-use marker
H_MARK = "\u2063#"        # invisible separator + '#', survives translation (we protect it anyway)


def blocks_to_text(blocks: List[Block]) -> str:
    lines = []
    for lvl, txt in blocks:
        txt = re.sub(r"\s+", " ", txt).strip()
        if not txt:
            continue
        lines.append(f"{H_MARK}{lvl} {txt}" if lvl else txt)
    return "\n".join(lines)


def text_to_blocks(text: str) -> List[Block]:
    out: List[Block] = []
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        m = re.match(rf"^{re.escape(H_MARK)}(\d)\s*(.*)$", line)
        if m:
            out.append((int(m.group(1)), m.group(2).strip()))
        else:
            out.append((None, line))
    return out


def extract_document(path: str, ext: str) -> str:
    """Extract → normalised line-per-paragraph text (headings marked)."""
    blocks = EXTRACTORS.get(ext, extract_txt)(path)
    return blocks_to_text(blocks)


# ═══════════════════════════════════════════════════════════════════════════
# ✂️ CHUNKING  (paragraph-safe: paragraphs are joined with a marker that
#               Google leaves intact, then split back → structure preserved)
# ═══════════════════════════════════════════════════════════════════════════
PARA_MARK = "\n⟦¶⟧\n"                     # unusual glyphs: MT engines copy them verbatim
_PARA_SPLIT_RE = re.compile(r"\s*⟦\s*¶\s*⟧\s*")
_SENTENCE_RE = re.compile(r"(?<=[.!?。！？])\s+")


def _split_long_paragraph(par: str, limit: int) -> List[str]:
    pieces, buf = [], ""
    for sent in _SENTENCE_RE.split(par):
        if len(sent) > limit:
            if buf:
                pieces.append(buf); buf = ""
            pieces.extend(sent[i:i + limit] for i in range(0, len(sent), limit))
            continue
        if len(buf) + len(sent) + 1 > limit and buf:
            pieces.append(buf); buf = sent
        else:
            buf = f"{buf} {sent}".strip()
    if buf:
        pieces.append(buf)
    return pieces


@dataclass
class Chunk:
    text: str                 # what gets sent to the API (paragraphs joined by PARA_MARK)
    n_paras: int              # how many paragraphs are inside
    glue_after: str = "\n"    # "\n" = paragraph boundary after, " " = same paragraph continues


def build_chunks(text: str, limit: int = CHUNK_SIZE) -> List[Chunk]:
    """Group whole paragraphs into ≤limit chunks; never cuts inside a sentence."""
    chunks: List[Chunk] = []
    buf: List[str] = []
    size = 0

    def flush():
        nonlocal buf, size
        if buf:
            chunks.append(Chunk(PARA_MARK.join(buf), len(buf), "\n"))
            buf, size = [], 0

    for par in text.split("\n"):
        par = par.strip()
        if not par:
            continue
        if len(par) > limit:
            flush()
            pieces = _split_long_paragraph(par, limit)
            for p in pieces[:-1]:
                chunks.append(Chunk(p, 1, " "))
            chunks.append(Chunk(pieces[-1], 1, "\n"))
            continue
        if size + len(par) + len(PARA_MARK) > limit:
            flush()
        buf.append(par)
        size += len(par) + len(PARA_MARK)
    flush()
    return chunks


def reassemble(chunks: List[Chunk], results: List[str]) -> str:
    """Rebuild line-per-paragraph text from translated chunks.
    Paragraph count per chunk is forced back to the source count, so the output
    stays aligned 1:1 with the source (needed for bilingual mode)."""
    out: List[str] = []
    carry = ""
    for ch, res in zip(chunks, results):
        src_paras = [p.strip() for p in ch.text.split(PARA_MARK)]
        paras = [p.strip() for p in _PARA_SPLIT_RE.split(res) if p.strip()]
        if not paras:
            paras = [res.strip()]
        n = ch.n_paras
        if len(paras) > n:                       # MT split a paragraph → merge extras into the last
            paras = paras[:n - 1] + [" ".join(paras[n - 1:])]
        elif len(paras) < n:                     # MT merged paragraphs → pad with untranslated source
            paras += src_paras[len(paras):]
        if carry:
            paras[0] = f"{carry} {paras[0]}".strip()
            carry = ""
        if ch.glue_after == " ":
            carry = paras.pop() if paras else ""
        out.extend(paras)
    if carry:
        out.append(carry)
    return "\n".join(out)


# ═══════════════════════════════════════════════════════════════════════════
# 📖 GLOSSARY / NAME-LOCK
# ═══════════════════════════════════════════════════════════════════════════
class Glossary:
    """Replace protected terms with numeric tokens before MT, restore after."""

    def __init__(self, mapping: Dict[str, str]):
        # longest first so "Li Wei Long" wins over "Li Wei"
        self.items = sorted(((k, v) for k, v in mapping.items() if k.strip()), key=lambda kv: -len(kv[0]))
        self.tokens = {k: f"⟪{i}⟫" for i, (k, _) in enumerate(self.items)}
        self.pattern = re.compile("|".join(re.escape(k) for k, _ in self.items), re.IGNORECASE) if self.items else None
        self.restore_re = re.compile(r"⟪\s*(\d+)\s*⟫")
        self.lookup = {i: v for i, (_, v) in enumerate(self.items)}
        self.lower = {k.lower(): k for k, _ in self.items}

    def protect(self, text: str) -> str:
        if not self.pattern:
            return text
        return self.pattern.sub(lambda m: self.tokens[self.lower[m.group(0).lower()]], text)

    def restore(self, text: str) -> str:
        if not self.pattern:
            return text
        return self.restore_re.sub(lambda m: self.lookup.get(int(m.group(1)), m.group(0)), text)

    # heading markers must also survive: protect H_MARK as a token too
    @staticmethod
    def protect_headings(text: str) -> str:
        return re.sub(rf"{re.escape(H_MARK)}(\d)\s*", lambda m: f"⟪H{m.group(1)}⟫ ", text)

    @staticmethod
    def restore_headings(text: str) -> str:
        return re.sub(r"⟪\s*[Hh]\s*(\d)\s*⟫\s*", lambda m: f"{H_MARK}{m.group(1)} ", text)


# ═══════════════════════════════════════════════════════════════════════════
# 📝 OUTPUT WRITERS  (blocks → file; RTL aware; real chapter TOC; bilingual)
# ═══════════════════════════════════════════════════════════════════════════
def _pairs(blocks: List[Block], src_blocks: Optional[List[Block]]) -> List[Tuple[Block, Optional[str]]]:
    """Zip translated blocks with source paragraphs for bilingual output."""
    if not src_blocks or len(src_blocks) != len(blocks):
        return [(bl, None) for bl in blocks]
    return [(bl, src[1]) for bl, src in zip(blocks, src_blocks)]


def write_txt(blocks, path, title, lang, src_blocks=None):
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"{title}\n{'=' * min(len(title), 60)}\n\n")
        for (lvl, txt), src in _pairs(blocks, src_blocks):
            if lvl:
                f.write(f"\n\n{'#' * lvl} {txt}\n\n")
            else:
                f.write(txt + "\n")
            if src and not lvl:
                f.write(f"    ⟶ {src}\n")
            f.write("\n")


def write_docx(blocks, path, title, lang, src_blocks=None):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
    d = docx.Document()
    rtl = lang in RTL_LANGS

    def _rtl(p):
        if not rtl:
            return
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pPr = p._p.get_or_add_pPr()
        bidi = OxmlElement("w:bidi"); bidi.set(qn("w:val"), "1"); pPr.append(bidi)

    _rtl(d.add_heading(title, level=0))
    for (lvl, txt), src in _pairs(blocks, src_blocks):
        if lvl:
            _rtl(d.add_heading(txt, level=min(lvl, 3)))
        else:
            _rtl(d.add_paragraph(txt))
            if src:
                p = d.add_paragraph()
                r = p.add_run(src); r.italic = True; r.font.size = Pt(9)
    d.save(path)


def _html_body(blocks, src_blocks, rtl: bool) -> str:
    out = []
    for (lvl, txt), src in _pairs(blocks, src_blocks):
        if lvl:
            out.append(f"<h{min(lvl + 1, 4)}>{html.escape(txt)}</h{min(lvl + 1, 4)}>")
        else:
            out.append(f"<p>{html.escape(txt)}</p>")
            if src:
                out.append(f'<p class="src" dir="auto"><i>{html.escape(src)}</i></p>')
    return "\n".join(out)


_HTML_CSS = ("body{max-width:46em;margin:2em auto;padding:0 1em;font-family:serif;line-height:1.7}"
             "h1,h2,h3,h4{margin-top:1.6em}p{margin:.6em 0;text-align:justify}"
             ".src{color:#666;font-size:.9em;margin:-.3em 0 1em 1.5em}")


def write_html(blocks, path, title, lang, src_blocks=None):
    rtl = lang in RTL_LANGS
    direction = "rtl" if rtl else "ltr"
    doc = (f'<!DOCTYPE html><html lang="{html.escape(lang)}" dir="{direction}"><head>'
           f'<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
           f"<title>{html.escape(title)}</title><style>{_HTML_CSS}</style></head><body>"
           f"<h1>{html.escape(title)}</h1>{_html_body(blocks, src_blocks, rtl)}"
           f'<p style="color:#999;font-size:.8em;margin-top:3em">{BOT_NAME} v{VERSION}</p></body></html>')
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)


def write_epub(blocks, path, title, lang, src_blocks=None):
    """Real chapters from detected headings (falls back to 350-paragraph sections)."""
    book = epub.EpubBook()
    book.set_identifier(uuid.uuid4().hex)
    book.set_title(title)
    book.set_language(lang.split("-")[0])
    book.add_author(BOT_NAME)
    rtl = lang in RTL_LANGS
    if rtl:
        try:
            book.set_direction("rtl")
        except Exception:
            pass
    css = epub.EpubItem(uid="style", file_name="style/main.css", media_type="text/css",
                        content=(_HTML_CSS + ("body{direction:rtl;text-align:right}" if rtl else "")).encode())
    book.add_item(css)

    pairs = _pairs(blocks, src_blocks)
    # group into chapters
    chapters: List[Tuple[str, List[Tuple[Block, Optional[str]]]]] = []
    cur_title: str = title
    cur: List[Tuple[Block, Optional[str]]] = []
    heading_count = sum(1 for (lvl, _), _ in pairs if lvl)
    use_headings = 1 <= heading_count <= 2000
    for (lvl, txt), src in pairs:
        if use_headings and lvl and lvl <= 2:
            if cur:
                chapters.append((cur_title, cur))
            cur_title, cur = txt, [((lvl, txt), None)]
        else:
            cur.append(((lvl, txt), src))
            if not use_headings and len(cur) >= 350:
                chapters.append((cur_title if not chapters else f"Section {len(chapters) + 1}", cur)); cur = []
    if cur:
        chapters.append((cur_title if not chapters or use_headings else f"Section {len(chapters) + 1}", cur))
    if not chapters:
        chapters = [(title, [((None, ""), None)])]

    items = []
    for n, (ctitle, cblocks) in enumerate(chapters, 1):
        ch = epub.EpubHtml(title=ctitle[:120], file_name=f"ch_{n:04d}.xhtml", lang=lang.split("-")[0])
        body_html = _html_body([bl for bl, _ in cblocks], [(None, s or "") for _, s in cblocks] if src_blocks else None, rtl)
        if n == 1 and (not use_headings or cblocks[0][0][0] is None):
            body_html = f"<h1>{html.escape(title)}</h1>" + body_html
        body_attr = ' dir="rtl"' if rtl else ""
        ch.content = (f'<html xmlns="http://www.w3.org/1999/xhtml"><head><title>{html.escape(ctitle)}</title>'
                      f'<link rel="stylesheet" href="style/main.css" type="text/css"/></head>'
                      f'<body{body_attr}>{body_html}</body></html>')
        ch.add_item(css)
        book.add_item(ch)
        items.append(ch)
    book.toc = tuple(items)
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    book.spine = ["nav"] + items
    epub.write_epub(path, book)


# Noto script fonts contain ONLY their own script — NotoSansDevanagari has no
# Latin glyphs and NotoSans has no Devanagari.  So we register one font per
# target language plus a Latin fallback, and switch fonts inline per run.
_NOTO_DIR = "/usr/share/fonts/truetype/noto"
_SCRIPT_FONTS: Dict[str, List[str]] = {
    "hi": ["NotoSansDevanagari-Regular.ttf"], "mr": ["NotoSansDevanagari-Regular.ttf"],
    "ne": ["NotoSansDevanagari-Regular.ttf"], "bn": ["NotoSansBengali-Regular.ttf"],
    "ta": ["NotoSansTamil-Regular.ttf"], "te": ["NotoSansTelugu-Regular.ttf"],
    "gu": ["NotoSansGujarati-Regular.ttf"], "kn": ["NotoSansKannada-Regular.ttf"],
    "ml": ["NotoSansMalayalam-Regular.ttf"], "pa": ["NotoSansGurmukhi-Regular.ttf"],
    "ur": ["NotoNastaliqUrdu-Regular.ttf", "NotoSansArabic-Regular.ttf"],
    "ar": ["NotoSansArabic-Regular.ttf"], "fa": ["NotoSansArabic-Regular.ttf"],
    "he": ["NotoSansHebrew-Regular.ttf"], "th": ["NotoSansThai-Regular.ttf"],
    "ja": ["NotoSansCJK-Regular.ttc", "NotoSansJP-Regular.ttf"],
    "ko": ["NotoSansCJK-Regular.ttc", "NotoSansKR-Regular.ttf"],
    "zh-CN": ["NotoSansCJK-Regular.ttc", "NotoSansSC-Regular.ttf"],
    "ru": ["NotoSans-Regular.ttf", "DejaVuSans.ttf"],
}
_LATIN_CANDIDATES = [
    os.path.join(_NOTO_DIR, "NotoSans-Regular.ttf"),
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "C:/Windows/Fonts/arial.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf",
]
_PDF_FONTS: Dict[str, str] = {}          # language → registered reportlab font name
_PDF_LATIN: Optional[str] = None
_LATIN_RUN = re.compile(r"[A-Za-z0-9\u00C0-\u024F][A-Za-z0-9\u00C0-\u024F\s.,;:!?'\"()\-–—/&%#@*+=\[\]]*")


def _register_font(name: str, path: str) -> bool:
    if not path or not os.path.exists(path) or not path.lower().endswith(".ttf"):
        return False
    try:
        pdfmetrics.registerFont(TTFont(name, path))
        return True
    except Exception as e:
        log.debug("font %s: %s", path, e)
        return False


def _pdf_latin_font() -> str:
    global _PDF_LATIN
    if _PDF_LATIN:
        return _PDF_LATIN
    for p in _LATIN_CANDIDATES:
        if _register_font("LatinFont", p):
            _PDF_LATIN = "LatinFont"
            return _PDF_LATIN
    _PDF_LATIN = "Helvetica"
    return _PDF_LATIN


def _pdf_font(lang: str = "") -> str:
    """Return a registered font that covers *lang*'s script (PDF_FONT_PATH wins)."""
    if lang in _PDF_FONTS:
        return _PDF_FONTS[lang]
    candidates = [PDF_FONT_PATH] if PDF_FONT_PATH else []
    candidates += [os.path.join(_NOTO_DIR, f) for f in _SCRIPT_FONTS.get(lang, [])]
    candidates += [os.path.join(_NOTO_DIR, "NotoSansDevanagari-Regular.ttf") if lang in ("", "hi") else "",
                   "C:/Windows/Fonts/Nirmala.ttf"] + _LATIN_CANDIDATES
    name = f"BookFont_{re.sub(r'[^A-Za-z0-9]', '', lang) or 'default'}"
    for p in candidates:
        if _register_font(name, p):
            _PDF_FONTS[lang] = name
            return name
    _PDF_FONTS[lang] = "Helvetica"
    log.warning("No Unicode TTF found for %s — PDF may show boxes. Set PDF_FONT_PATH.", lang or "default")
    return "Helvetica"


def _pdf_markup(text: str, script_font: str, latin_font: str) -> str:
    """Escape for reportlab Paragraph and wrap Latin runs in the Latin font so
    'Part 1 of 3 — नॉवेल' renders both scripts (script fonts lack Latin glyphs)."""
    if script_font == latin_font or script_font == "Helvetica":
        return html.escape(text)
    return _LATIN_RUN.sub(lambda m: f'<font name="{latin_font}">{html.escape(m.group(0))}</font>',
                          html.escape(text))


def write_pdf(blocks, path, title, lang, src_blocks=None):
    if not HAS_PDF_OUT:
        raise RuntimeError("reportlab not installed")
    font = _pdf_font(lang)
    latin = _pdf_latin_font()
    rtl = lang in RTL_LANGS
    base = ParagraphStyle("b", fontName=font, fontSize=11, leading=17, spaceAfter=6,
                          alignment=2 if rtl else 4)
    h = {1: ParagraphStyle("h1", parent=base, fontSize=20, leading=26, spaceBefore=18, spaceAfter=12),
         2: ParagraphStyle("h2", parent=base, fontSize=16, leading=22, spaceBefore=16, spaceAfter=10),
         3: ParagraphStyle("h3", parent=base, fontSize=13, leading=19, spaceBefore=12, spaceAfter=8)}
    small = ParagraphStyle("s", parent=base, fontSize=9, leading=13, textColor=colors.HexColor("#666666"))
    doc = SimpleDocTemplate(path, pagesize=A4, title=title, author=BOT_NAME,
                            leftMargin=18 * mm, rightMargin=18 * mm, topMargin=18 * mm, bottomMargin=18 * mm)
    story = [Paragraph(_pdf_markup(title, font, latin), h[1]), Spacer(1, 8)]
    for (lvl, txt), src in _pairs(blocks, src_blocks):
        story.append(Paragraph(_pdf_markup(txt, font, latin), h.get(lvl, base) if lvl else base))
        if src and not lvl:
            story.append(Paragraph(f"<i>{_pdf_markup(src, font, latin)}</i>", small))
    doc.build(story)


WRITERS: Dict[str, Callable] = {"txt": write_txt, "docx": write_docx, "epub": write_epub, "html": write_html}
if HAS_PDF_OUT:
    WRITERS["pdf"] = write_pdf


def write_output(fmt: str, translated_text: str, path: str, title: str, lang: str,
                 source_text: Optional[str] = None):
    blocks = text_to_blocks(translated_text)
    src_blocks = text_to_blocks(source_text) if source_text else None
    WRITERS[fmt](blocks, path, title, lang, src_blocks)


def make_zip(paths: List[str], zip_path: str):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            if os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))


def split_blocks_by_output_size(text: str, max_kb: int) -> List[str]:
    """Split *translated* line-per-paragraph text into ≤max_kb UTF-8 parts on
    paragraph boundaries, preferring chapter headings as cut points (bug #9)."""
    if max_kb <= 0:
        return [text]
    max_bytes = max_kb * 1024
    lines = text.split("\n")
    parts, cur, size = [], [], 0
    for i, line in enumerate(lines):
        lb = len(line.encode("utf-8")) + 1
        is_head = line.startswith(H_MARK)
        # cut at a heading if we're already past 70 % of the budget
        if cur and ((size + lb > max_bytes) or (is_head and size > max_bytes * 0.7)):
            parts.append("\n".join(cur)); cur, size = [], 0
        cur.append(line); size += lb
    if cur:
        parts.append("\n".join(cur))
    return parts


# ═══════════════════════════════════════════════════════════════════════════
# 🧩 JOB MODEL & FAIR QUEUE  (persistent, round-robin, owner priority)
# ═══════════════════════════════════════════════════════════════════════════
class JobCancelled(Exception):
    pass


@dataclass(eq=False)                    # identity semantics → QUEUE.index / `in` are safe & fast
class Job:
    job_id: str
    chat_id: int
    user_id: int
    user_name: str
    file_path: str                      # source file ("" when text comes from doc cache)
    ext: str
    novel_name: str
    file_size: int
    lang: str = DEFAULT_LANG
    src: str = DEFAULT_SRC
    out_format: str = DEFAULT_FORMAT
    split_kb: int = DEFAULT_SPLIT_KB
    bilingual: bool = False
    doc_key: str = ""                   # extracted-text cache key (for re-translate)
    priority: int = 0                   # higher = sooner (owner = 10)
    status: str = "pending"             # pending|queued|running|done|cancelled|failed
    created: float = field(default_factory=time.time)
    est_chars: int = 0
    msg: Optional[Message] = None
    msg_id: int = 0
    cancel: asyncio.Event = field(default_factory=asyncio.Event)
    cancelled_by_owner: bool = False

    @property
    def lang_name(self) -> str:
        return lang_name(self.lang)

    @property
    def out_dir(self) -> str:
        return os.path.join(STORAGE_DIR, f"{self.chat_id}_{self.job_id}")

    @property
    def ckpt_path(self) -> str:
        return os.path.join(CKPT_DIR, f"{self.job_id}.json")

    def cleanup(self, keep_source: bool = False):
        try:
            if not keep_source and self.file_path and os.path.exists(self.file_path):
                os.remove(self.file_path)
        except Exception:
            pass
        shutil.rmtree(self.out_dir, ignore_errors=True)
        try:
            os.remove(self.ckpt_path)
        except Exception:
            pass

    # -- persistence ---------------------------------------------------------
    def to_dict(self) -> dict:
        return {k: getattr(self, k) for k in (
            "job_id", "chat_id", "user_id", "user_name", "file_path", "ext", "novel_name", "file_size",
            "lang", "src", "out_format", "split_kb", "bilingual", "doc_key", "priority", "created",
            "est_chars", "msg_id")}

    @classmethod
    def from_dict(cls, d: dict) -> "Job":
        j = cls(**{k: d[k] for k in ("job_id", "chat_id", "user_id", "user_name", "file_path", "ext",
                                     "novel_name", "file_size")})
        for k in ("lang", "src", "out_format", "split_kb", "bilingual", "doc_key", "priority", "created",
                  "est_chars", "msg_id"):
            if k in d:
                setattr(j, k, d[k])
        j.status = "queued"
        return j


PENDING: Dict[str, Job] = {}
QUEUE: List[Job] = []
ACTIVE: Optional[Job] = None
QUEUE_WAKE = asyncio.Event()
USER_STATE: Dict[int, dict] = {}                 # chat_id -> {"state":..., ...}
RECENT_DURATIONS: deque = deque(maxlen=20)       # seconds per 100k chars, for ETA
SHUTTING_DOWN = False


def persist_queue():
    store.save_queue([j.to_dict() for j in QUEUE])


def queue_position(job: Job) -> int:
    try:
        return QUEUE.index(job) + 1
    except ValueError:
        return 0


def user_queued_count(uid: int) -> int:
    return sum(1 for j in QUEUE if j.user_id == uid) + (1 if ACTIVE and ACTIVE.user_id == uid else 0)


def user_pending_count(uid: int) -> int:
    return sum(1 for j in PENDING.values() if j.user_id == uid)


def _fair_insert(job: Job):
    """Round-robin fairness.  A user's k-th queued job belongs to 'round k'; it is
    inserted right before the first job that belongs to a later round.  So one user
    dumping 10 books can never block another user's first book.
    Priority jobs (owner / /priority) stay at the front, ordered among themselves."""
    if job.priority > 0:
        idx = 0
        while idx < len(QUEUE) and QUEUE[idx].priority >= job.priority:
            idx += 1
        QUEUE.insert(idx, job)
        return
    my_round = sum(1 for j in QUEUE if j.user_id == job.user_id) + 1
    counts: Dict[int, int] = defaultdict(int)
    pos = len(QUEUE)
    for i, j in enumerate(QUEUE):
        if j.priority > 0:
            continue
        counts[j.user_id] += 1
        if counts[j.user_id] > my_round:
            pos = i
            break
    QUEUE.insert(pos, job)


def chars_per_second() -> float:
    if not RECENT_DURATIONS:
        return 900.0                               # conservative default
    return sum(RECENT_DURATIONS) / len(RECENT_DURATIONS)


def queue_eta_seconds(before: List[Job]) -> float:
    cps = chars_per_second()
    total = sum(max(j.est_chars, 20_000) for j in before)
    if ACTIVE:
        total += max(ACTIVE.est_chars, 20_000) * 0.5
    return total / cps


def summary_block(job: Job, uid: Optional[int] = None) -> str:
    uid = uid or job.user_id
    return (
        f"📘 File: {b(job.novel_name)}\n"
        f"💾 Size: {b(fmt_size(job.file_size))}"
        + (f"  ·  🔤 {b(fmt_int(job.est_chars))} chars" if job.est_chars else "") + "\n"
        f"🌐 {T(uid, 'language')}: {b(lang_label(job.lang))}"
        + (f"  ·  {T(uid, 'source')}: {b(lang_name(job.src))}" if job.src != 'auto' else "") + "\n"
        f"📄 {T(uid, 'format')}: {b(job.out_format.upper())}"
        + ("  ·  🔁 " + T(uid, 'bilingual') if job.bilingual else "") + "\n"
        f"✂️ {T(uid, 'split')}: {b(split_label(job.split_kb))}\n"
    )


def queued_text(job: Job, pos: int) -> str:
    uid = job.user_id
    eta = queue_eta_seconds(QUEUE[:pos - 1])
    return (
        header(T(uid, "queued"), "⏳") + summary_block(job) + f"{DIV}\n"
        f"📍 {T(uid, 'position')}: {b('#' + str(pos))}  ·  ⏱ {T(uid, 'eta_queue')}: {b('~' + fmt_time(eta))}\n"
        f"{T(uid, 'next_up') if pos == 1 else T(uid, 'please_wait')}"
    )


def cancel_kb(job: Job) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[btn(T(job.user_id, "cancel_job"), f"qx:{job.job_id}")]])


async def enqueue(job: Job, message: Optional[Message]):
    job.msg = message
    job.msg_id = message.id if message else 0
    job.status = "queued"
    if store.is_owner(job.user_id):
        job.priority = max(job.priority, 5)
    _fair_insert(job)
    persist_queue()
    await LiveMessage(message).update(queued_text(job, queue_position(job)), cancel_kb(job), force=True)
    QUEUE_WAKE.set()


async def notify_positions():
    """Staggered so 20 queued users don't trigger a FloodWait (bug #20)."""
    for i, j in enumerate(list(QUEUE), 1):
        if j not in QUEUE:
            continue
        await LiveMessage(j.msg).update(queued_text(j, i), cancel_kb(j), force=True)
        await asyncio.sleep(0.35)


def remove_from_queue(job: Job):
    if job in QUEUE:
        QUEUE.remove(job)
        persist_queue()


async def restore_queue():
    """Re-hydrate persisted jobs after a restart (bug #5)."""
    saved = store.data.get("queue") or []
    restored = 0
    for d in saved:
        try:
            j = Job.from_dict(d)
            if j.file_path and not os.path.exists(j.file_path) and not j.doc_key:
                continue
            if j.msg_id:
                try:
                    j.msg = await app.get_messages(j.chat_id, j.msg_id)
                except Exception:
                    j.msg = None
            if j.msg is None:
                j.msg = await safe_send(j.chat_id, header("Restored", "♻️") + summary_block(j) +
                                        f"{DIV}\nBot restarted — your job is back in the queue.")
                j.msg_id = j.msg.id if j.msg else 0
            QUEUE.append(j)
            restored += 1
        except Exception as e:
            log.warning("restore job failed: %s", e)
    QUEUE.sort(key=lambda x: (-x.priority, x.created))
    persist_queue()
    if restored:
        log.info("Restored %d queued job(s)", restored)
        QUEUE_WAKE.set()
        asyncio.create_task(notify_positions())


async def queue_worker():
    global ACTIVE
    log.info("Queue worker started")
    while not SHUTTING_DOWN:
        if not QUEUE:
            QUEUE_WAKE.clear()
            await QUEUE_WAKE.wait()
            continue
        job = QUEUE.pop(0)
        persist_queue()
        ACTIVE = job
        job.status = "running"
        asyncio.create_task(notify_positions())
        if time.time() - job.created > 90:                 # only ping if the user actually waited
            await safe_send(job.chat_id, T(job.user_id, "your_turn", name=b(job.novel_name)))
        try:
            await process_job(job)
        except Exception as e:
            log.exception("Worker error: %s", e)
        finally:
            ACTIVE = None
            if job.status in ("done", "cancelled", "failed") and not SHUTTING_DOWN:
                job.cleanup()


async def janitor():
    """Drop stale wizards & temp files — never touching files owned by live jobs (bug #2)."""
    while not SHUTTING_DOWN:
        await asyncio.sleep(600)
        now = time.time()
        for jid, job in list(PENDING.items()):
            if now - job.created > PENDING_TTL:
                job.cleanup()
                PENDING.pop(jid, None)
                if USER_STATE.get(job.chat_id, {}).get("job_id") == jid:
                    USER_STATE.pop(job.chat_id, None)
                log.info("Janitor removed stale wizard %s", jid)
        protected = {j.file_path for j in list(PENDING.values()) + QUEUE + ([ACTIVE] if ACTIVE else [])}
        for name in os.listdir(INBOX_DIR):
            p = os.path.join(INBOX_DIR, name)
            if p in protected:
                continue
            try:
                if now - os.path.getmtime(p) > 4 * 3600:
                    os.remove(p)
            except Exception:
                pass
        live_dirs = {os.path.basename(j.out_dir) for j in QUEUE + ([ACTIVE] if ACTIVE else [])}
        for name in os.listdir(STORAGE_DIR):
            if name not in live_dirs:
                shutil.rmtree(os.path.join(STORAGE_DIR, name), ignore_errors=True)
        live_ck = {j.job_id + ".json" for j in QUEUE + ([ACTIVE] if ACTIVE else [])}
        for name in os.listdir(CKPT_DIR):
            p = os.path.join(CKPT_DIR, name)
            if name not in live_ck and now - os.path.getmtime(p) > 48 * 3600:
                try:
                    os.remove(p)
                except Exception:
                    pass
        if random.random() < 0.05:
            await asyncio.to_thread(cache.vacuum_old)


async def health_ping():
    while not SHUTTING_DOWN:
        await asyncio.sleep(HEALTH_PING_HOURS * 3600)
        owner = store.owner_id()
        if owner:
            s = store.data["stats"]
            await safe_send(owner, header("Health", "💚") +
                            f"Alive · uptime {fmt_time(time.time() - STARTED_AT)}\n"
                            f"📚 Jobs: {b(s['jobs'])}  ·  ❌ Failed: {b(s['failed'])}\n"
                            f"▶️ Active: {b(ACTIVE.novel_name if ACTIVE else '—')}  ·  ⏳ Queue: {b(len(QUEUE))}")


# ═══════════════════════════════════════════════════════════════════════════
# 🌐 TRANSLATION ENGINE  (cache → Google → MyMemory → Libre, global 429 gate)
# ═══════════════════════════════════════════════════════════════════════════
class RateGate:
    """When any worker hits a 429 / ban, every worker pauses together (bug #8)."""

    def __init__(self):
        self.resume_at = 0.0

    def trip(self, seconds: float = COOLDOWN_429):
        self.resume_at = max(self.resume_at, time.time() + seconds)

    @property
    def remaining(self) -> float:
        return max(0.0, self.resume_at - time.time())

    async def wait(self, cancel: asyncio.Event):
        while self.remaining > 0:
            if cancel.is_set():
                raise JobCancelled()
            await asyncio.sleep(min(1.0, self.remaining))


GATE = RateGate()
_BAD_RESPONSE = re.compile(r"(Error 5\d\d|Server Error|Our systems have detected unusual traffic|captcha)", re.I)


_MM_CODES: Optional[Dict[str, str]] = None


def _mymemory_lang(c: str) -> str:
    """MyMemory wants region codes (hi-IN, en-GB…) — map lazily from the library's own table."""
    global _MM_CODES
    if _MM_CODES is None:
        try:
            _MM_CODES = {v.lower(): v for v in MyMemoryTranslator.get_supported_languages(as_dict=True).values()}
        except Exception:
            _MM_CODES = {}
    c = "en" if c == "auto" else c
    low = c.lower()
    if low in _MM_CODES:
        return _MM_CODES[low]
    base = low.split("-")[0]
    for k, v in _MM_CODES.items():
        if k == base or k.startswith(base + "-"):
            return v
    return c


def _translate_sync(text: str, src: str, tgt: str, engine: str) -> str:
    if engine == "google":
        return GoogleTranslator(source=src, target=tgt).translate(text)
    if engine == "mymemory":
        if len(text) > 500:                              # MyMemory hard limit
            raise RequestError("too long for mymemory")
        return MyMemoryTranslator(source=_mymemory_lang(src), target=_mymemory_lang(tgt)).translate(text)
    if engine == "libre" and LibreTranslator is not None:
        return LibreTranslator(source=src, target=tgt.split("-")[0], base_url=LIBRE_URL).translate(text)
    raise RequestError(f"engine {engine} unavailable")


class TranslationEngine:
    ENGINES = ("google", "mymemory", "libre")

    def __init__(self, target: str, source: str = "auto", glossary: Optional[Glossary] = None,
                 concurrency: int = CONCURRENCY_LIMIT):
        self.target, self.source = target, source
        self.glossary = glossary or Glossary({})
        self.sem = asyncio.Semaphore(concurrency)
        self.total = self.done = self.failed = self.cache_hits = 0
        self.started = time.time()
        self.phase = ""

    async def _translate_one(self, idx: int, chunk: Chunk, cancel: asyncio.Event) -> Tuple[int, str]:
        raw = chunk.text
        if not raw.strip():
            self.done += 1
            return idx, raw
        prepared = self.glossary.protect(Glossary.protect_headings(raw))
        cached = await asyncio.to_thread(cache.get, prepared, self.source, self.target)
        if cached is not None:
            self.done += 1; self.cache_hits += 1
            return idx, Glossary.restore_headings(self.glossary.restore(cached))

        # attempt plan: Google ×3 → Libre (if available) → MyMemory (short chunks only)
        plan = ["google"] * max(1, MAX_RETRIES - 2)
        if LibreTranslator is not None:
            plan.append("libre")
        if len(prepared) <= 500:
            plan.append("mymemory")

        async with self.sem:
            delay = 2.0
            for attempt, engine in enumerate(plan):
                if cancel.is_set():
                    raise JobCancelled()
                await GATE.wait(cancel)
                try:
                    res = await asyncio.to_thread(_translate_sync, prepared, self.source, self.target, engine)
                    if not res or not isinstance(res, str) or _BAD_RESPONSE.search(res):
                        raise RuntimeError("Bad API response")
                    # sanity: MT must not silently swallow 80 % of the text
                    if len(res) < len(prepared) * 0.2 and len(prepared) > 200:
                        raise RuntimeError("Suspiciously short response")
                    await asyncio.to_thread(cache.put, prepared, self.source, self.target, res)
                    self.done += 1
                    return idx, Glossary.restore_headings(self.glossary.restore(res))
                except JobCancelled:
                    raise
                except TooManyRequests:
                    GATE.trip(COOLDOWN_429)
                    log.warning("429 from %s → global cooldown %ss", engine, COOLDOWN_429)
                except Exception as e:
                    msg = str(e).lower()
                    if "429" in msg or "too many" in msg or "unusual traffic" in msg or "captcha" in msg:
                        GATE.trip(COOLDOWN_429)
                    log.debug("chunk %d attempt %d (%s) failed: %s", idx, attempt + 1, engine, e)
                    await asyncio.sleep(delay + random.uniform(0, 1.5))
                    delay = min(delay * 2, 30)
            self.done += 1; self.failed += 1
            return idx, raw

    async def run(self, chunks: List[Chunk], cancel: asyncio.Event, on_progress) -> List[str]:
        self.total = len(chunks)
        self.started = time.time()
        results: List[Optional[str]] = [None] * self.total
        tasks = [asyncio.create_task(self._translate_one(i, c, cancel)) for i, c in enumerate(chunks)]
        try:
            for fut in asyncio.as_completed(tasks):
                idx, text = await fut
                results[idx] = text
                await on_progress(self)
        except BaseException:
            for t in tasks:
                t.cancel()
            await asyncio.gather(*tasks, return_exceptions=True)          # bug #7: no zombie tasks
            raise
        return [results[i] if results[i] is not None else chunks[i].text for i in range(self.total)]

    @property
    def speed(self) -> float:
        return self.done / max(time.time() - self.started, 0.001)

    @property
    def eta(self) -> float:
        return (self.total - self.done) / self.speed if self.speed > 0 else 0.0


def dashboard_text(job: Job, part: int, total_parts: int, eng: TranslationEngine, phase: str) -> str:
    uid = job.user_id
    ratio = eng.done / eng.total if eng.total else 0
    cool = GATE.remaining
    return (
        header(T(uid, "dashboard"), "🖥") +
        f"📘 {b(job.novel_name)}\n"
        f"🌐 {lang_label(job.lang)}  ·  📄 {job.out_format.upper()}"
        + ("  ·  🔁" if job.bilingual else "") + "\n"
        f"📦 {T(uid, 'part')} {b(f'{part} / {total_parts}')}\n{DIV}\n"
        f"{progress_bar(ratio)} {b(f'{ratio * 100:.1f}%')}\n"
        f"🧩 {T(uid, 'chunks')}: {eng.done} / {eng.total}"
        + (f"  ·  💾 {T(uid, 'cached')}: {eng.cache_hits}" if eng.cache_hits else "") + "\n"
        f"⚡ {T(uid, 'speed')}: {eng.speed:.2f}/s  ·  ⏱ {fmt_time(time.time() - eng.started)}  ·  {T(uid, 'eta')} {fmt_time(eng.eta)}\n"
        + (f"⚠️ {T(uid, 'retries')}: {eng.failed}\n" if eng.failed else "")
        + (T(uid, "cooling", s=int(cool)) + "\n" if cool > 0 else "")
        + f"{DIV}\n<i>{esc(phase)}</i>"
    )


# ═══════════════════════════════════════════════════════════════════════════
# 🚀 MAIN PIPELINE  (extract → split by output → translate → write → deliver)
# ═══════════════════════════════════════════════════════════════════════════
def _load_ckpt(job: Job) -> dict:
    try:
        with open(job.ckpt_path, "r", encoding="utf-8") as f:
            d = json.load(f)
        if d.get("lang") == job.lang and d.get("fmt") == job.out_format and d.get("split") == job.split_kb:
            return d
    except Exception:
        pass
    return {"lang": job.lang, "fmt": job.out_format, "split": job.split_kb, "topic": None, "done_parts": {}}


def _save_ckpt(job: Job, d: dict):
    try:
        with open(job.ckpt_path, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
    except Exception as e:
        log.debug("ckpt save: %s", e)


async def get_source_text(job: Job) -> str:
    """Extract (or fetch from doc cache) → normalised text; caches result."""
    if job.doc_key:
        hit = await asyncio.to_thread(cache.get_doc, job.doc_key)
        if hit:
            return hit[1]
    if not job.file_path or not os.path.exists(job.file_path):
        return ""
    text = await asyncio.to_thread(extract_document, job.file_path, job.ext)
    if text and job.doc_key:
        await asyncio.to_thread(cache.put_doc, job.doc_key, job.novel_name, text)
    return text


async def translate_text_block(text: str, job: Job, live: Optional[LiveMessage], part: int, total_parts: int,
                               eng: TranslationEngine) -> str:
    chunks = build_chunks(text)

    async def on_progress(e: TranslationEngine):
        if job.cancel.is_set():
            raise JobCancelled()
        if live:
            await live.update(dashboard_text(job, part, total_parts, e, T(job.user_id, "translating")), cancel_kb(job))

    results = await eng.run(chunks, job.cancel, on_progress)
    return reassemble(chunks, results)


async def process_job(job: Job):
    uid = job.user_id
    live = LiveMessage(job.msg)
    os.makedirs(job.out_dir, exist_ok=True)
    started = time.time()
    total_chars = failed_chunks = cache_hits = 0
    delivered: List[str] = []
    ck = _load_ckpt(job)

    try:
        # 1 ─ extract ───────────────────────────────────────────────────────
        await live.update(header(T(uid, "processing"), "🔄") + summary_block(job) + f"{DIV}\n{T(uid, 'extracting')}",
                          cancel_kb(job), force=True)
        text = await get_source_text(job)
        if len(text) < 20:
            job.status = "failed"
            await live.update(header(T(uid, "failed"), "❌") + summary_block(job) + f"{DIV}\n{T(uid, 'no_text')}",
                              back_home_kb(uid), force=True)
            return
        if job.cancel.is_set():
            raise JobCancelled()
        total_chars = len(text)
        job.est_chars = total_chars

        # quota check on real size
        if store.quota_left(uid) < total_chars:
            job.status = "failed"
            hrs = 24 - time.gmtime().tm_hour
            await live.update(header(T(uid, "failed"), "❌") + summary_block(job) + f"{DIV}\n" +
                              T(uid, "quota", used=fmt_int(store.quota_used(uid)), limit=fmt_int(DAILY_QUOTA_CHARS), hrs=hrs),
                              back_home_kb(uid), force=True)
            return

        # 2 ─ pre-split the SOURCE using the expansion factor so OUTPUT parts hit the target (bug #9)
        factor = OUTPUT_EXPANSION.get(job.lang, 1.5) * (1.9 if job.bilingual else 1.0)
        src_kb = max(1, int(job.split_kb / factor)) if job.split_kb else 0
        parts = split_blocks_by_output_size(text, src_kb) if src_kb else [text]
        total_parts = len(parts)

        # 3 ─ backup topic ──────────────────────────────────────────────────
        topic_id = ck.get("topic") or await create_backup_topic(job.novel_name)
        ck["topic"] = topic_id
        _save_ckpt(job, ck)
        await backup_send(header("New Job", "📥") + summary_block(job) +
                          f"👤 {b(job.user_name)} ({code(job.user_id)})\n"
                          f"🧩 Parts: {b(total_parts)}  ·  🔤 Chars: {b(fmt_int(total_chars))}", topic_id)

        glossary = Glossary(store.glossary(uid))

        # 4 ─ translate part by part ────────────────────────────────────────
        for i, part_text in enumerate(parts, 1):
            if job.cancel.is_set():
                raise JobCancelled()
            part_title = f"{job.novel_name} — Part {i} of {total_parts}" if total_parts > 1 else job.novel_name
            out_name = safe_filename(f"Part {i} of {total_parts} - {job.novel_name} [{job.lang_name}]", 120) + f".{job.out_format}"
            out_path = os.path.join(job.out_dir, out_name)

            done_info = ck["done_parts"].get(str(i))
            if done_info and done_info.get("file_id"):
                # resume: already delivered before a restart
                delivered.append(out_path if os.path.exists(out_path) else "")
                continue

            eng = TranslationEngine(job.lang, job.src, glossary)
            translated = await translate_text_block(part_text, job, live, i, total_parts, eng)
            failed_chunks += eng.failed
            cache_hits += eng.cache_hits

            await live.update(dashboard_text(job, i, total_parts, eng, T(uid, "building", fmt=job.out_format)),
                              cancel_kb(job), force=True)
            await asyncio.to_thread(write_output, job.out_format, translated, out_path, part_title, job.lang,
                                    part_text if job.bilingual else None)
            delivered.append(out_path)

            caption = (f"📘 {b(job.novel_name)}\n"
                       f"📦 {T(uid, 'part')} {i} / {total_parts}  ·  🌐 {lang_label(job.lang)}\n"
                       f"💾 {fmt_size(os.path.getsize(out_path))}")
            sent = await safe_send_document(job.chat_id, out_path, caption)
            file_id = sent.document.file_id if sent and sent.document else None
            # backup reuses the uploaded file_id → no second upload (bug #11)
            await backup_doc(file_id or out_path, caption + f"\n👤 {b(job.user_name)} ({code(job.user_id)})", topic_id)
            ck["done_parts"][str(i)] = {"file_id": file_id or "", "name": out_name}
            _save_ckpt(job, ck)

        # 5 ─ ZIP when multiple parts ───────────────────────────────────────
        existing = [p for p in delivered if p and os.path.exists(p)]
        if total_parts > 1 and len(existing) == total_parts:
            zip_path = os.path.join(job.out_dir, safe_filename(f"{job.novel_name} [{job.lang_name}]") + ".zip")
            await asyncio.to_thread(make_zip, existing, zip_path)
            if os.path.getsize(zip_path) < 1990 * 1024 * 1024:
                zsent = await safe_send_document(job.chat_id, zip_path, f"📦 {b(job.novel_name)} — {T(uid, 'merge_zip')}")
                if zsent and zsent.document:
                    await backup_doc(zsent.document.file_id, f"📦 {b(job.novel_name)} ZIP", topic_id)

        # 6 ─ done ──────────────────────────────────────────────────────────
        job.status = "done"
        elapsed = time.time() - started
        if total_chars > 5000 and elapsed > 5:
            RECENT_DURATIONS.append(total_chars / elapsed)
        store.quota_add(uid, total_chars)
        store.bump(uid, total_parts, total_chars, cache_hits=cache_hits)
        store.add_history(uid, {"name": job.novel_name, "lang": job.lang, "fmt": job.out_format,
                                "parts": total_parts, "chars": total_chars, "at": int(time.time()),
                                "doc_key": job.doc_key, "topic": topic_id})
        await live.update(
            header(T(uid, "completed"), "✅") + summary_block(job) + f"{DIV}\n"
            f"🧩 {T(uid, 'parts')}: {b(total_parts)}  ·  🔤 {T(uid, 'chars')}: {b(fmt_int(total_chars))}\n"
            f"⏱ {T(uid, 'time')}: {b(fmt_time(elapsed))}"
            + (f"  ·  💾 {T(uid, 'cached')}: {cache_hits}" if cache_hits else "") + "\n"
            + (T(uid, "kept_original", n=failed_chunks) + "\n" if failed_chunks else "")
            + f"{DIV}\n{T(uid, 'thanks', bot=BOT_NAME)}",
            InlineKeyboardMarkup(([[btn(T(uid, "retranslate"), f"rt:{job.doc_key}")]] if job.doc_key else [])
                                 + [[btn(f"🏠 {T(uid, 'home')}", "nav:home")]]), force=True)

    except JobCancelled:
        job.status = "cancelled"
        if SHUTTING_DOWN:                                   # paused for restart — message handled by shutdown
            return
        who = " (by owner)" if job.cancelled_by_owner else ""
        await live.update(header(T(uid, "cancelled"), "🛑") + summary_block(job) + f"{DIV}\n{T(uid, 'stopped')}{who}",
                          back_home_kb(uid), force=True)
        await backup_send(f"🛑 Job cancelled{who}: {b(job.novel_name)} by {code(job.user_id)}", ck.get("topic"))
    except Exception as e:
        job.status = "failed"
        store.bump(uid, 0, 0, failed=True)
        log.exception("Processing error for %s: %s", job.novel_name, e)
        await live.update(header(T(uid, "error"), "⚠️") + summary_block(job) + f"{DIV}\n"
                          f"{T(uid, 'went_wrong')}\n{code(str(e)[:200])}", back_home_kb(uid), force=True)
        await backup_send(f"⚠️ Job failed: {b(job.novel_name)} — {code(str(e)[:150])}", ck.get("topic"))


# ═══════════════════════════════════════════════════════════════════════════
# 🤖 CLIENT & FILTERS
# ═══════════════════════════════════════════════════════════════════════════
app = Client("novel_translator_pro", api_id=API_ID, api_hash=API_HASH, bot_token=BOT_TOKEN,
             workdir=BASE_DIR, parse_mode=enums.ParseMode.HTML)
STARTED_AT = time.time()


async def _auth_msg(_, __, m: Message):
    return bool(m.from_user and store.is_authorized(m.from_user.id))


async def _auth_cb(_, __, q: CallbackQuery):
    return bool(q.from_user and store.is_authorized(q.from_user.id))


async def _owner_msg(_, __, m: Message):
    return bool(m.from_user and store.is_owner(m.from_user.id))


authorized = filters.create(_auth_msg)
authorized_cb = filters.create(_auth_cb)
owner_only = filters.create(_owner_msg)
PRIVATE = filters.private & filters.incoming

# matches ONLY our own output pattern  "Part 3 of 7 - Name [Hindi].txt"  (not "Part 1 - The Beginning.epub")
OUTPUT_NAME_RE = re.compile(r"^Part\s*\d+\s*of\s*\d+\s*[|\-–].*\[[^\]]+\]\.\w+$", re.IGNORECASE)

# ── brute-force protection (bug #13) ───────────────────────────────────────
_BRUTE: Dict[int, dict] = {}          # uid -> {"n": attempts, "until": lock_ts, "last_reply": ts}


def brute_locked(uid: int) -> float:
    st = _BRUTE.get(uid)
    return max(0.0, st["until"] - time.time()) if st else 0.0


def brute_fail(uid: int) -> float:
    if len(_BRUTE) > 5000:                                   # memory cap
        for k in [k for k, v in _BRUTE.items() if v["until"] < time.time()][:2500]:
            _BRUTE.pop(k, None)
    st = _BRUTE.setdefault(uid, {"n": 0, "until": 0.0, "last_reply": 0.0})
    st["n"] += 1
    if st["n"] >= BRUTE_MAX_ATTEMPTS:
        st["n"] = 0
        st["until"] = time.time() + BRUTE_LOCK_SECONDS
        return BRUTE_LOCK_SECONDS
    return 0.0


# ═══════════════════════════════════════════════════════════════════════════
# 🔒 UNAUTHORIZED
# ═══════════════════════════════════════════════════════════════════════════
@app.on_message(PRIVATE & ~authorized)
async def unauthorized_message(_, m: Message):
    if not m.from_user:
        return
    uid = m.from_user.id
    if store.is_banned(uid):
        return                                                 # silent
    if brute_locked(uid) > 0:
        st = _BRUTE[uid]
        if time.time() - st["last_reply"] > 60:                # at most one reply / min while locked
            st["last_reply"] = time.time()
            await m.reply(T(0, "locked_out", mins=int(brute_locked(uid) // 60) + 1))
        return
    if m.text and m.text.strip() == store.security_code():
        _BRUTE.pop(uid, None)
        store.authorize(uid, m.from_user.first_name or "User")
        role = "👑 Owner" if store.is_owner(uid) else "👤 User"
        await m.reply(header(T(uid, "granted"), "✅") + T(uid, "welcome", name=b(m.from_user.first_name), role=b(role)))
        log.info("Authorized %s (%s)", uid, role)
        owner = store.owner_id()
        if owner and owner != uid:
            await safe_send(owner, f"🆕 New user unlocked: {b(m.from_user.first_name)} ({code(uid)})")
        return
    if m.text:
        lock = brute_fail(uid)
        if lock:
            log.warning("Brute-force lockout for %s", uid)
            return await m.reply(T(0, "locked_out", mins=int(lock // 60)))
    await m.reply(header(T(0, "locked_title"), "🔒") + T(0, "locked"))


@app.on_callback_query(~authorized_cb)
async def unauthorized_callback(_, q: CallbackQuery):
    await q.answer("🔒 Not authorized.", show_alert=True)


# ═══════════════════════════════════════════════════════════════════════════
# 💬 USER COMMANDS
# ═══════════════════════════════════════════════════════════════════════════
@app.on_message(filters.command("start") & PRIVATE & authorized)
async def cmd_start(_, m: Message):
    USER_STATE.pop(m.chat.id, None)
    await m.reply(text_home(m.from_user.id, m.from_user.first_name), reply_markup=home_keyboard(m.from_user.id))


@app.on_message(filters.command("help") & PRIVATE & authorized)
async def cmd_help(_, m: Message):
    await m.reply(text_help(m.from_user.id), reply_markup=back_home_kb(m.from_user.id))


@app.on_message(filters.command("settings") & PRIVATE & authorized)
async def cmd_settings(_, m: Message):
    await m.reply(text_settings(m.from_user.id), reply_markup=settings_keyboard(m.from_user.id))


def text_queue(uid: int) -> str:
    lines = [header(T(uid, "queue"), "📋")]
    if ACTIVE:
        who = "you" if ACTIVE.user_id == uid else esc(ACTIVE.user_name)
        lines.append(f"▶️ Running: {b(ACTIVE.novel_name)}\n   <i>by {who}</i>\n")
    else:
        lines.append("▶️ Running: <i>nothing</i>\n")
    if QUEUE:
        lines.append(f"⏳ Waiting: {b(len(QUEUE))}")
        for i, j in enumerate(QUEUE[:10], 1):
            mark = "🟢" if j.user_id == uid else ("👑" if j.priority > 0 else "⚪️")
            lines.append(f"{mark} #{i} {esc(j.novel_name[:32])}  <i>~{fmt_time(queue_eta_seconds(QUEUE[:i - 1]))}</i>")
        if len(QUEUE) > 10:
            lines.append(f"… and {len(QUEUE) - 10} more")
    else:
        lines.append("⏳ Waiting: <i>empty</i>")
    if GATE.remaining > 0:
        lines.append(f"\n⏸ API cooldown: {int(GATE.remaining)}s")
    return "\n".join(lines)


def queue_kb(uid: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[btn(f"🔄 {T(uid, 'refresh')}", "nav:queue"), btn(f"🏠 {T(uid, 'home')}", "nav:home")]])


@app.on_message(filters.command("queue") & PRIVATE & authorized)
async def cmd_queue(_, m: Message):
    await m.reply(text_queue(m.from_user.id), reply_markup=queue_kb(m.from_user.id))


async def cancel_user_jobs(uid: int, by_owner: bool = False) -> int:
    """Cancel everything belonging to a user (used by /cancel, /deluser, /ban)."""
    n = 0
    if ACTIVE and ACTIVE.user_id == uid:
        ACTIVE.cancelled_by_owner = by_owner
        ACTIVE.cancel.set(); n += 1
    for j in [j for j in QUEUE if j.user_id == uid]:
        remove_from_queue(j); j.status = "cancelled"; j.cleanup(); n += 1
        await LiveMessage(j.msg).update(header(T(uid, "cancelled"), "🛑") + summary_block(j), back_home_kb(uid), force=True)
    for jid, j in [(k, v) for k, v in PENDING.items() if v.user_id == uid]:
        PENDING.pop(jid, None); j.cleanup(); n += 1
    if n:
        asyncio.create_task(notify_positions())
    return n


@app.on_message(filters.command("cancel") & PRIVATE & authorized)
async def cmd_cancel(_, m: Message):
    uid = m.from_user.id
    USER_STATE.pop(m.chat.id, None)
    if len(m.command) > 1 and store.is_owner(uid):                 # /cancel <job_id>  (owner)
        jid = m.command[1]
        if ACTIVE and ACTIVE.job_id == jid:
            ACTIVE.cancelled_by_owner = True; ACTIVE.cancel.set()
            return await m.reply(f"🛑 Stopping {b(ACTIVE.novel_name)}…")
        for j in QUEUE:
            if j.job_id == jid:
                remove_from_queue(j); j.status = "cancelled"; j.cleanup()
                await LiveMessage(j.msg).update(header(T(j.user_id, "cancelled"), "🛑") + summary_block(j) + "\n(by owner)",
                                                back_home_kb(j.user_id), force=True)
                asyncio.create_task(notify_positions())
                return await m.reply(f"🛑 Removed {b(j.novel_name)}.")
        return await m.reply("⚠️ Job id not found.")
    n = await cancel_user_jobs(uid)
    await m.reply(T(uid, "cancelled_n", n=b(n)) if n else T(uid, "no_jobs"))


def text_mystats(uid: int) -> str:
    u = store.user(uid) or {}
    s = u.get("stats", {})
    since = time.strftime("%d %b %Y", time.localtime(u.get("joined", time.time())))
    left = store.quota_left(uid)
    return (
        header(T(uid, "mystats"), "📊") +
        f"👤 {b(u.get('name', 'User'))}  ·  {code(uid)}\n"
        f"🎖 Role: {b(u.get('role', 'user').title())}  ·  📅 {b(since)}\n{DIV}\n"
        f"📚 Files: {b(s.get('jobs', 0))}  ·  🧩 Parts: {b(s.get('parts', 0))}\n"
        f"🔤 Characters: {b(fmt_int(s.get('chars', 0)))}\n"
        f"📈 Today's quota left: {b('∞' if left > 10 ** 11 else fmt_int(left))}\n"
        f"📖 Glossary terms: {b(len(u.get('glossary', {})))}"
    )


@app.on_message(filters.command("mystats") & PRIVATE & authorized)
async def cmd_mystats(_, m: Message):
    await m.reply(text_mystats(m.from_user.id), reply_markup=back_home_kb(m.from_user.id))


def text_history(uid: int) -> Tuple[str, InlineKeyboardMarkup]:
    h = store.pref(uid, "history", []) or []
    lines = [header(T(uid, "history"), "🕘")]
    rows = []
    if not h:
        lines.append(T(uid, "history_empty"))
    for i, e in enumerate(h[:10], 1):
        when = time.strftime("%d %b %H:%M", time.localtime(e.get("at", 0)))
        lines.append(f"{i}. {b(e.get('name', '?')[:40])}\n    {lang_label(e.get('lang', ''))} · {e.get('fmt', '').upper()} · "
                     f"{e.get('parts', 0)} part(s) · {fmt_int(e.get('chars', 0))} chars · <i>{when}</i>")
        if e.get("doc_key"):
            rows.append(btn(f"↩️ {i}. {e.get('name', '?')[:22]}", f"rt:{e['doc_key']}"))
    kb = kb_rows(rows, 2) + [[btn(f"🏠 {T(uid, 'home')}", "nav:home")]]
    return "\n".join(lines), InlineKeyboardMarkup(kb)


@app.on_message(filters.command("history") & PRIVATE & authorized)
async def cmd_history(_, m: Message):
    t, kb = text_history(m.from_user.id)
    await m.reply(t, reply_markup=kb)


def text_glossary(uid: int) -> str:
    g = store.glossary(uid)
    lines = [header(T(uid, "glossary"), "📖")]
    if not g:
        lines.append(T(uid, "glossary_empty"))
    else:
        for k, v in sorted(g.items())[:60]:
            lines.append(f"• {esc(k)} → {b(v)}")
        if len(g) > 60:
            lines.append(f"… {len(g) - 60} more")
        lines.append(f"\n<i>Add: <code>/glossary Term=Translation</code>\nRemove: <code>/glossary -Term</code>\nClear: <code>/glossary --clear</code></i>")
    return "\n".join(lines)


@app.on_message(filters.command("glossary") & PRIVATE & authorized)
async def cmd_glossary(_, m: Message):
    uid = m.from_user.id
    arg = m.text.split(None, 1)[1].strip() if len(m.command) > 1 else ""
    if not arg:
        return await m.reply(text_glossary(uid), reply_markup=back_home_kb(uid))
    if arg == "--clear":
        store.set_pref(uid, "glossary", {})
        return await m.reply("🗑 Glossary cleared.")
    if arg.startswith("-"):
        ok = store.glossary_del(uid, arg[1:].strip())
        return await m.reply("🗑 Removed." if ok else "⚠️ Term not found.")
    added = 0
    for pair in re.split(r"[;\n]+", arg):                          # allow several: A=B; C=D
        if "=" in pair:
            k, v = pair.split("=", 1)
            k, v = k.strip(), v.strip()
            if 1 <= len(k) <= 60 and 1 <= len(v) <= 80 and len(store.glossary(uid)) < 300:
                store.glossary_set(uid, k, v); added += 1
    if not added:
        return await m.reply("Usage: <code>/glossary Li Wei=ली वेई</code>")
    await m.reply(f"✅ {added} term(s) saved.\n\n" + text_glossary(uid))


async def quick_translate(m: Message, text: str, tgt: str, src: str = "auto"):
    uid = m.from_user.id
    if len(text) > QUICK_TEXT_MAX:
        return await m.reply(f"⚠️ Max {QUICK_TEXT_MAX} characters for quick translation. Send a file instead.")
    if store.quota_left(uid) < len(text):
        return await m.reply(T(uid, "quota", used=fmt_int(store.quota_used(uid)), limit=fmt_int(DAILY_QUOTA_CHARS), hrs=24 - time.gmtime().tm_hour))
    status = await m.reply("⏳ …")
    eng = TranslationEngine(tgt, src, Glossary(store.glossary(uid)), concurrency=2)
    job = Job("quick", m.chat.id, uid, m.from_user.first_name or "User", "", ".txt", "quick", len(text), lang=tgt, src=src)
    try:
        out = await translate_text_block(blocks_to_text(blocks_from_plain(text)), job, None, 1, 1, eng)
        out = "\n\n".join(t for _, t in text_to_blocks(out))
        store.quota_add(uid, len(text))
        body = T(uid, "quick_result", lang=lang_label(tgt), div=DIV, text=esc(out))
        for i in range(0, len(body), 4000):
            if i == 0:
                await status.edit_text(body[:4000])
            else:
                await m.reply(body[i:i + 4000])
    except Exception as e:
        await status.edit_text(f"⚠️ {esc(str(e)[:200])}")


@app.on_message(filters.command("tr") & PRIVATE & authorized)
async def cmd_tr(_, m: Message):
    """/tr <lang> <text>   or reply to a message with /tr <lang>"""
    uid = m.from_user.id
    parts = m.text.split(None, 2)
    tgt = parts[1] if len(parts) > 1 and parts[1] in LANGUAGES else None
    text = parts[2] if len(parts) > 2 else None
    if not text and m.reply_to_message:
        text = m.reply_to_message.text or m.reply_to_message.caption
    if not tgt or not text:
        return await m.reply("Usage: <code>/tr hi Hello world</code>\nor reply to a message with <code>/tr hi</code>\n"
                             f"Languages: {', '.join(LANGUAGES)}")
    await quick_translate(m, text, tgt, store.pref(uid, "src", DEFAULT_SRC))


# ═══════════════════════════════════════════════════════════════════════════
# 👑 OWNER COMMANDS
# ═══════════════════════════════════════════════════════════════════════════
def text_owner() -> str:
    s = store.data["stats"]
    d = store.data["daily"].get(Store._today(), {"jobs": 0, "chars": 0, "failed": 0})
    tr_n, doc_n = cache.stats()
    return (
        header("Owner Panel", "👑") +
        f"👥 Users: {b(len(store.data['users']))}  ·  🚫 Banned: {b(len(store.data['banned']))}\n"
        f"💬 Tracked chats: {b(len(store.data['chats']))}\n"
        f"📚 Jobs: {b(s['jobs'])}  ·  ❌ Failed: {b(s['failed'])}\n"
        f"🧩 Parts: {b(s['parts'])}  ·  🔤 Chars: {b(fmt_int(s['chars']))}\n"
        f"💾 Cache: {b(fmt_int(tr_n))} chunks · {b(doc_n)} docs · {b(fmt_int(s.get('cache_hits', 0)))} hits\n"
        f"📅 Today: {b(d['jobs'])} jobs · {b(fmt_int(d['chars']))} chars\n"
        f"▶️ Active: {b(ACTIVE.novel_name if ACTIVE else '—')}  ·  ⏳ Queue: {b(len(QUEUE))}\n"
        f"⏱ Uptime: {b(fmt_time(time.time() - STARTED_AT))}  ·  v{VERSION}\n"
        f"{DIV}\n{b('Commands')}\n"
        "/users · /adduser &lt;id&gt; · /deluser &lt;id&gt;\n"
        "/ban &lt;id&gt; [reason] · /unban &lt;id&gt;\n"
        "/broadcast &lt;text&gt; (or reply)\n"
        "/setcode &lt;new code&gt; · /links\n"
        "/cancelall · /cancel &lt;job_id&gt; · /priority &lt;job_id&gt;\n"
        "/logs · /weekly · /restart · /stats"
    )


def owner_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [btn("👥 Users", "ow:users"), btn("📋 Queue", "nav:queue")],
        [btn("📈 Weekly", "ow:weekly"), btn("🧾 Logs", "ow:logs")],
        [btn("🔄 Refresh", "nav:owner"), btn("🏠 Home", "nav:home")],
    ])


@app.on_message(filters.command("stats") & PRIVATE & owner_only)
async def cmd_stats(_, m: Message):
    await m.reply(text_owner(), reply_markup=owner_kb())


def text_users() -> str:
    users = store.data["users"]
    lines = [header(f"Users ({len(users)})", "👥")]
    for uid, u in sorted(users.items(), key=lambda kv: -kv[1]["stats"].get("jobs", 0))[:60]:
        crown = "👑 " if u.get("role") == "owner" else ""
        lines.append(f"{crown}{esc(u.get('name', 'User'))} — {code(uid)} · {u['stats'].get('jobs', 0)} jobs · {fmt_int(u['stats'].get('chars', 0))}")
    if len(users) > 60:
        lines.append(f"… and {len(users) - 60} more")
    if store.data["banned"]:
        lines.append(f"\n🚫 Banned: {', '.join(code(x) for x in list(store.data['banned'])[:30])}")
    return "\n".join(lines)


@app.on_message(filters.command("users") & PRIVATE & owner_only)
async def cmd_users(_, m: Message):
    await m.reply(text_users())


def _uid_arg(m: Message) -> Optional[int]:
    if len(m.command) < 2:
        return None
    a = m.command[1]
    return int(a) if a.lstrip("-").isdigit() else None


@app.on_message(filters.command("adduser") & PRIVATE & owner_only)
async def cmd_adduser(_, m: Message):
    uid = _uid_arg(m)
    if uid is None:
        return await m.reply("Usage: <code>/adduser 123456789</code>")
    store.unban(uid)
    store.authorize(uid, "Added by owner")
    await m.reply(f"✅ User {code(uid)} authorized.")
    await safe_send(uid, header("Access Granted", "✅") + "You have been authorized.\nSend /start to begin.")


@app.on_message(filters.command("deluser") & PRIVATE & owner_only)
async def cmd_deluser(_, m: Message):
    uid = _uid_arg(m)
    if uid is None:
        return await m.reply("Usage: <code>/deluser 123456789</code>")
    n = await cancel_user_jobs(uid, by_owner=True)                   # bug #23
    ok = store.revoke(uid)
    await m.reply(f"🗑 User {code(uid)} removed. Cancelled {n} job(s)." if ok else "⚠️ Not found (or is the owner).")


@app.on_message(filters.command("ban") & PRIVATE & owner_only)
async def cmd_ban(_, m: Message):
    uid = _uid_arg(m)
    if uid is None:
        return await m.reply("Usage: <code>/ban 123456789 [reason]</code>")
    if store.is_owner(uid):
        return await m.reply("⚠️ Cannot ban the owner.")
    reason = m.text.split(None, 2)[2] if len(m.command) > 2 else ""
    n = await cancel_user_jobs(uid, by_owner=True)
    store.revoke(uid)
    store.ban(uid, reason)
    await m.reply(f"🚫 Banned {code(uid)}. Cancelled {n} job(s).")
    await safe_send(uid, T(uid, "banned"))


@app.on_message(filters.command("unban") & PRIVATE & owner_only)
async def cmd_unban(_, m: Message):
    uid = _uid_arg(m)
    if uid is None:
        return await m.reply("Usage: <code>/unban 123456789</code>")
    await m.reply(f"✅ Unbanned {code(uid)}." if store.unban(uid) else "⚠️ Not banned.")


@app.on_message(filters.command("setcode") & PRIVATE & owner_only)
async def cmd_setcode(_, m: Message):
    if len(m.command) < 2:
        return await m.reply("Usage: <code>/setcode NewSecret123</code>")
    new_code = m.text.split(None, 1)[1].strip()
    if len(new_code) < 8:
        return await m.reply("⚠️ Code must be at least 8 characters.")
    store.set_security_code(new_code)
    await m.reply(f"🔐 Security code updated to {code(new_code)}\n<i>Delete this message for safety.</i>")


@app.on_message(filters.command("broadcast") & PRIVATE & owner_only)
async def cmd_broadcast(_, m: Message):
    src = m.reply_to_message
    text = m.text.split(None, 1)[1] if len(m.command) > 1 else None
    if not src and not text:
        return await m.reply("Usage: <code>/broadcast Hello everyone</code>\nor reply to a message with /broadcast")
    sent = failed = 0
    status = await m.reply("📣 Broadcasting…")
    for uid in list(store.data["users"].keys()):
        for attempt in range(2):                                   # bug #16: retry after FloodWait
            try:
                if src:
                    await src.copy(int(uid))
                else:
                    await app.send_message(int(uid), header("Announcement", "📣") + esc(text))
                sent += 1
                break
            except FloodWait as e:
                await asyncio.sleep(float(e.value) + 1)
            except Exception:
                failed += 1
                break
        await asyncio.sleep(0.1)
    await status.edit_text(f"📣 Broadcast done.\n✅ Sent: {b(sent)}  ·  ❌ Failed: {b(failed)}")


@app.on_message(filters.command("cancelall") & PRIVATE & owner_only)
async def cmd_cancelall(_, m: Message):
    n = 0
    for j in list(QUEUE):
        remove_from_queue(j); j.status = "cancelled"; j.cleanup(); n += 1
        await LiveMessage(j.msg).update(header(T(j.user_id, "cancelled"), "🛑") + summary_block(j) + "\n(by owner)",
                                        back_home_kb(j.user_id), force=True)
    if ACTIVE:
        ACTIVE.cancelled_by_owner = True; ACTIVE.cancel.set(); n += 1
    await m.reply(f"🛑 Cancelled {b(n)} job(s).")


@app.on_message(filters.command("priority") & PRIVATE & owner_only)
async def cmd_priority(_, m: Message):
    if len(m.command) < 2:
        listing = "\n".join(f"{code(j.job_id)} · {esc(j.novel_name[:30])} · {esc(j.user_name)}" for j in QUEUE[:20]) or "<i>queue empty</i>"
        return await m.reply(f"Usage: <code>/priority &lt;job_id&gt;</code>\n\n{listing}")
    jid = m.command[1]
    for j in QUEUE:
        if j.job_id == jid:
            remove_from_queue(j); j.priority = 10; _fair_insert(j); persist_queue()
            asyncio.create_task(notify_positions())
            return await m.reply(f"⏫ {b(j.novel_name)} moved to the front.")
    await m.reply("⚠️ Job not found in queue.")


@app.on_message(filters.command("logs") & PRIVATE & owner_only)
async def cmd_logs(_, m: Message):
    try:
        await m.reply_document(log_tail_file(), caption=f"🧾 Last 300 log lines · {fmt_size(os.path.getsize(LOG_FILE))} total")
    except Exception as e:
        await m.reply(f"⚠️ {esc(e)}")


def log_tail_file(lines: int = 300) -> io.BytesIO:
    with open(LOG_FILE, "rb") as f:
        f.seek(max(0, os.path.getsize(LOG_FILE) - 400_000))
        data = f.read()
    tail = "\n".join(data.decode("utf-8", errors="replace").splitlines()[-lines:])
    buf = io.BytesIO(tail.encode("utf-8"))
    buf.name = "bot_tail.log"
    return buf


async def send_weekly_chart(chat_id: int):
    days = [time.strftime("%Y-%m-%d", time.gmtime(time.time() - 86400 * i)) for i in range(6, -1, -1)]
    daily = store.data["daily"]
    jobs = [daily.get(d, {}).get("jobs", 0) for d in days]
    chars = [daily.get(d, {}).get("chars", 0) / 1e6 for d in days]
    text = header("Weekly", "📈") + "\n".join(f"{d[5:]}  ·  {j} jobs  ·  {c:.2f}M chars" for d, j, c in zip(days, jobs, chars))
    if not HAS_MPL:
        return await safe_send(chat_id, text + "\n\n<i>pip install matplotlib for a chart</i>")
    buf = io.BytesIO()

    def _draw():
        fig, ax1 = plt.subplots(figsize=(8, 4), dpi=130)
        ax1.bar([d[5:] for d in days], jobs, color="#4F8EF7", label="Jobs")
        ax1.set_ylabel("Jobs")
        ax2 = ax1.twinx()
        ax2.plot([d[5:] for d in days], chars, color="#F76F4F", marker="o", label="M chars")
        ax2.set_ylabel("Million chars")
        ax1.set_title(f"{BOT_NAME} — last 7 days")
        fig.tight_layout(); fig.savefig(buf, format="png"); plt.close(fig)

    await asyncio.to_thread(_draw)
    buf.seek(0)
    buf.name = "weekly.png"
    try:
        await app.send_photo(chat_id, buf, caption=text)
    except Exception as e:
        await safe_send(chat_id, text + f"\n⚠️ chart: {esc(e)}")


@app.on_message(filters.command("weekly") & PRIVATE & owner_only)
async def cmd_weekly(_, m: Message):
    await send_weekly_chart(m.chat.id)


@app.on_message(filters.command("restart") & PRIVATE & owner_only)
async def cmd_restart(_, m: Message):
    global _RESTART
    await m.reply("♻️ Restarting… queued jobs are persisted and will resume.")
    _RESTART = True
    if STOP_EVENT:
        STOP_EVENT.set()


@app.on_message(filters.command("links") & PRIVATE & owner_only)
async def cmd_links(client: Client, m: Message):
    status = await m.reply("🔄 Collecting invite links…")
    chats = dict(store.data["chats"])
    if BACKUP_GROUP_ID and str(BACKUP_GROUP_ID) not in chats:
        chats[str(BACKUP_GROUP_ID)] = {"title": "Backup Group", "type": "supergroup"}
    lines = []
    for cid in chats:
        try:
            chat = await client.get_chat(int(cid))
            me = await client.get_chat_member(int(cid), "me")
            if me.status not in (enums.ChatMemberStatus.ADMINISTRATOR, enums.ChatMemberStatus.OWNER):
                continue
            link = chat.invite_link or await client.export_chat_invite_link(int(cid))
            lines.append(f"📌 {b(chat.title)}\n🔗 {esc(link)}")
        except Exception as e:
            log.debug("links: %s -> %s", cid, e)
    if lines:
        await status.edit_text(header("Admin Invite Links", "🔗") + "\n\n".join(lines), disable_web_page_preview=True)
    else:
        await status.edit_text("❌ No chats found where the bot is admin.\n<i>Add the bot to a group/channel as admin first.</i>")


@app.on_chat_member_updated()
async def track_membership(_, upd: ChatMemberUpdated):
    try:
        new = upd.new_chat_member
        if not new or not new.user or not new.user.is_self:
            return
        if new.status in (enums.ChatMemberStatus.LEFT, enums.ChatMemberStatus.BANNED):
            store.untrack_chat(upd.chat.id)
        else:
            store.track_chat(upd.chat.id, upd.chat.title or "Chat", str(upd.chat.type).split(".")[-1].lower())
    except Exception as e:
        log.debug("membership tracking: %s", e)


# ═══════════════════════════════════════════════════════════════════════════
# 📥 DOCUMENT INTAKE → WIZARD
# ═══════════════════════════════════════════════════════════════════════════
def wizard_lang_text(job: Job) -> str:
    return header(T(job.user_id, "step_lang"), "🌐") + summary_block(job) + f"{DIV}\n{T(job.user_id, 'choose_lang')}"


def wizard_lang_kb(job: Job) -> InlineKeyboardMarkup:
    uid = job.user_id
    kb = language_keyboard(f"jl:{job.job_id}:", job.lang, f"jx:{job.job_id}")
    rows = [[btn(T(uid, "quick"), f"jq:{job.job_id}")]]
    rows += kb.inline_keyboard[:-1]
    rows.append([btn(T(uid, "more_opts"), f"jo:{job.job_id}"), btn(f"❌ {T(uid, 'cancel')}", f"jx:{job.job_id}")])
    return InlineKeyboardMarkup(rows)


def wizard_opts_text(job: Job) -> str:
    return header(T(job.user_id, "more_opts"), "🔧") + summary_block(job)


def wizard_opts_kb(job: Job) -> InlineKeyboardMarkup:
    uid = job.user_id
    return InlineKeyboardMarkup([
        [btn(f"🔤 {T(uid, 'source')}: {lang_name(job.src)}", f"jsrc:{job.job_id}:menu")],
        [btn(f"🔁 {T(uid, 'bilingual')}: {'ON' if job.bilingual else 'OFF'}", f"jbil:{job.job_id}")],
        [btn(f"◀️ {T(uid, 'back')}", f"jb:{job.job_id}:lang")],
    ])


def wizard_fmt_text(job: Job) -> str:
    return header(T(job.user_id, "step_fmt"), "📄") + summary_block(job) + f"{DIV}\n{T(job.user_id, 'choose_fmt')}"


def wizard_fmt_kb(job: Job) -> InlineKeyboardMarkup:
    return format_keyboard(f"jf:{job.job_id}:", job.out_format, f"jb:{job.job_id}:lang", f"◀️ {T(job.user_id, 'back')}")


def wizard_split_text(job: Job) -> str:
    return header(T(job.user_id, "step_split"), "✂️") + summary_block(job) + f"{DIV}\n{T(job.user_id, 'choose_split')}"


def wizard_split_kb(job: Job) -> InlineKeyboardMarkup:
    return split_keyboard(f"js:{job.job_id}:", job.split_kb, f"jb:{job.job_id}:fmt", f"jc:{job.job_id}",
                          T(job.user_id, "custom_size"), f"◀️ {T(job.user_id, 'back')}")


def wizard_confirm_text(job: Job) -> str:
    uid = job.user_id
    factor = OUTPUT_EXPANSION.get(job.lang, 1.5) * (1.9 if job.bilingual else 1.0)
    out_bytes = (job.est_chars or job.file_size) * (factor if job.est_chars else 0.9)
    est_parts = max(1, -(-int(out_bytes) // (job.split_kb * 1024))) if job.split_kb else 1
    ahead = len(QUEUE) + (1 if ACTIVE else 0)
    return (header(T(uid, "step_confirm"), "✅") + summary_block(job) + f"{DIV}\n"
            f"🧩 {T(uid, 'est_parts')}: {b('~' + str(est_parts))}\n"
            f"👥 {T(uid, 'ahead')}: {b(ahead)}  ·  ⏱ {T(uid, 'eta_queue')}: {b('~' + fmt_time(queue_eta_seconds(QUEUE)))}\n\n"
            f"{T(uid, 'looks_good')}")


def wizard_confirm_kb(job: Job) -> InlineKeyboardMarkup:
    uid = job.user_id
    return InlineKeyboardMarkup([
        [btn(T(uid, "start"), f"jgo:{job.job_id}")],
        [btn(T(uid, "preview"), f"jp:{job.job_id}")],
        [btn(f"◀️ {T(uid, 'back')}", f"jb:{job.job_id}:spl"), btn(f"❌ {T(uid, 'cancel')}", f"jx:{job.job_id}")],
    ])


def new_job_for(uid: int, chat_id: int, name: str, novel_name: str, ext: str, size: int, doc_key: str) -> Job:
    return Job(job_id=uuid.uuid4().hex[:8], chat_id=chat_id, user_id=uid, user_name=name or "User",
               file_path="", ext=ext, novel_name=novel_name, file_size=size,
               lang=store.pref(uid, "lang", DEFAULT_LANG), src=store.pref(uid, "src", DEFAULT_SRC),
               out_format=store.pref(uid, "fmt", DEFAULT_FORMAT) if store.pref(uid, "fmt", DEFAULT_FORMAT) in OUTPUT_FORMATS else DEFAULT_FORMAT,
               split_kb=int(store.pref(uid, "split", DEFAULT_SPLIT_KB)),
               bilingual=bool(store.pref(uid, "bilingual", False)), doc_key=doc_key)


def _limits_ok(uid: int, check_pending: bool = True) -> Optional[str]:
    if check_pending and user_pending_count(uid) >= MAX_PENDING_PER_USER and not store.is_owner(uid):
        return T(uid, "too_many_pending", n=user_pending_count(uid))
    if user_queued_count(uid) >= MAX_QUEUED_PER_USER and not store.is_owner(uid):
        return T(uid, "too_many_queued", n=user_queued_count(uid))
    if store.quota_left(uid) <= 0:
        return T(uid, "quota", used=fmt_int(store.quota_used(uid)), limit=fmt_int(DAILY_QUOTA_CHARS), hrs=24 - time.gmtime().tm_hour)
    return None


@app.on_message(filters.document & PRIVATE & authorized)
async def handle_document(_, m: Message):
    uid = m.from_user.id
    doc = m.document
    file_name = doc.file_name or "Document"
    ext = os.path.splitext(file_name)[1].lower()
    if OUTPUT_NAME_RE.search(file_name):                                        # bug #14: explain, offer re-translate
        kb = InlineKeyboardMarkup([[btn(T(uid, "retranslate"), "nav:history")]])
        return await m.reply(T(uid, "own_output"), reply_markup=kb)
    if ext not in INPUT_EXTS:
        return await m.reply(header(T(uid, "unsupported"), "❌") +
                             f"{T(uid, 'allowed')}: {b(', '.join(sorted(INPUT_EXTS)))}\n"
                             f"{T(uid, 'you_sent')}: {code(ext or 'no extension')}")
    if doc.file_size and doc.file_size > MAX_INPUT_MB * 1024 * 1024:
        return await m.reply(T(uid, "too_large", mb=MAX_INPUT_MB))
    err = _limits_ok(uid)
    if err:
        return await m.reply(err)

    USER_STATE.pop(m.chat.id, None)
    novel_name = safe_filename(os.path.splitext(file_name)[0])
    doc_key = hashlib.sha1(f"{doc.file_unique_id}".encode()).hexdigest()[:16]
    job = new_job_for(uid, m.chat.id, m.from_user.first_name, novel_name, ext, doc.file_size or 0, doc_key)
    job.file_path = os.path.join(INBOX_DIR, f"src_{job.job_id}{ext}")

    status = await m.reply(header(T(uid, "downloading"), "📥") + f"📘 {b(novel_name)}\n{progress_bar(0)} 0%")
    live = LiveMessage(status, min_interval=3.0)

    async def dl_progress(current: int, total: int):
        r = current / total if total else 0
        await live.update(header(T(uid, "downloading"), "📥") + f"📘 {b(novel_name)}\n"
                          f"{progress_bar(r)} {r * 100:.0f}%\n{fmt_size(current)} / {fmt_size(total)}")

    cached_doc = await asyncio.to_thread(cache.get_doc, doc_key)
    if cached_doc:
        job.est_chars = len(cached_doc[1])
    else:
        try:
            await m.download(file_name=job.file_path, progress=dl_progress)
        except Exception as e:
            log.error("Download failed: %s", e)
            job.cleanup()                                                        # bug #15
            return await live.update(header(T(uid, "download_failed"), "❌") + T(uid, "send_again"), force=True)
        # extract right away → real estimate in the wizard (bug #10) + doc cache for re-translate
        await live.update(header(T(uid, "processing"), "🔄") + f"📘 {b(novel_name)}\n{T(uid, 'extracting')}", force=True)
        try:
            text = await get_source_text(job)
            job.est_chars = len(text)
        except Exception as e:
            log.warning("pre-extract failed: %s", e)
        if job.est_chars < 20:
            job.cleanup()
            return await live.update(header(T(uid, "failed"), "❌") + f"📘 {b(novel_name)}\n{T(uid, 'no_text')}",
                                     back_home_kb(uid), force=True)
    PENDING[job.job_id] = job
    await live.update(wizard_lang_text(job), wizard_lang_kb(job), force=True)


async def start_retranslate(q: CallbackQuery, doc_key: str):
    """Re-translate a previously extracted document straight from the doc cache."""
    uid = q.from_user.id
    hit = await asyncio.to_thread(cache.get_doc, doc_key)
    if not hit:
        return await q.answer("⌛ Source no longer cached — please send the file again.", show_alert=True)
    err = _limits_ok(uid)
    if err:
        return await q.answer(re.sub(r"<[^>]+>", "", err)[:190], show_alert=True)
    name, text = hit
    job = new_job_for(uid, q.message.chat.id, q.from_user.first_name, name, ".txt", len(text.encode("utf-8")), doc_key)
    job.est_chars = len(text)
    PENDING[job.job_id] = job
    await safe_send(q.message.chat.id, wizard_lang_text(job), reply_markup=wizard_lang_kb(job))
    await q.answer()


# ═══════════════════════════════════════════════════════════════════════════
# 🔘 CALLBACKS
# ═══════════════════════════════════════════════════════════════════════════
async def _edit(q: CallbackQuery, text: str, kb: Optional[InlineKeyboardMarkup] = None):
    try:
        await q.message.edit_text(text, reply_markup=kb, disable_web_page_preview=True)
    except MessageNotModified:
        pass
    except FloodWait as e:
        await asyncio.sleep(float(e.value) + 1)
        try:
            await q.message.edit_text(text, reply_markup=kb, disable_web_page_preview=True)
        except Exception:
            pass
    except Exception as e:
        log.warning("callback edit failed: %s", e)


def _pending(q: CallbackQuery, job_id: str) -> Optional[Job]:
    job = PENDING.get(job_id)
    if not job or job.chat_id != q.message.chat.id or job.user_id != q.from_user.id:
        return None
    return job


async def run_preview(q: CallbackQuery, job: Job):
    uid = job.user_id
    await _edit(q, header(T(uid, "preview_title"), "👀") + summary_block(job) + f"{DIV}\n⏳ …")
    text = await get_source_text(job)
    sample_lines, n = [], 0
    for line in text.split("\n"):
        if n >= PREVIEW_CHARS:
            break
        sample_lines.append(line); n += len(line)
    sample = "\n".join(sample_lines)
    eng = TranslationEngine(job.lang, job.src, Glossary(store.glossary(uid)), concurrency=3)
    try:
        out = await translate_text_block(sample, job, None, 1, 1, eng)
        out = "\n\n".join(t for _, t in text_to_blocks(out))
    except Exception as e:
        out = f"⚠️ {e}"
    body = header(T(uid, "preview_title"), "👀") + esc(out[:2600]) + f"\n\n{T(uid, 'preview_note', n=PREVIEW_CHARS)}"
    await _edit(q, body, wizard_confirm_kb(job))


@app.on_callback_query(authorized_cb)
async def callbacks(_, q: CallbackQuery):
    if not q.message:
        return await q.answer()
    data = q.data or ""
    uid = q.from_user.id
    chat_id = q.message.chat.id
    parts = data.split(":")
    kind = parts[0]
    arg1 = parts[1] if len(parts) > 1 else ""
    arg2 = parts[2] if len(parts) > 2 else ""

    try:
        # ── navigation ──────────────────────────────────────────────────
        if kind == "nav":
            USER_STATE.pop(chat_id, None)
            if arg1 == "home":
                await _edit(q, text_home(uid, q.from_user.first_name), home_keyboard(uid))
            elif arg1 == "help":
                await _edit(q, text_help(uid), back_home_kb(uid))
            elif arg1 == "settings":
                await _edit(q, text_settings(uid), settings_keyboard(uid))
            elif arg1 == "queue":
                await _edit(q, text_queue(uid), queue_kb(uid))
            elif arg1 == "mystats":
                await _edit(q, text_mystats(uid), back_home_kb(uid))
            elif arg1 == "history":
                t, kb = text_history(uid); await _edit(q, t, kb)
            elif arg1 == "glossary":
                await _edit(q, text_glossary(uid), back_home_kb(uid))
            elif arg1 == "owner" and store.is_owner(uid):
                await _edit(q, text_owner(), owner_kb())
            return await q.answer()

        # ── owner panel buttons ─────────────────────────────────────────
        if kind == "ow" and store.is_owner(uid):
            if arg1 == "users":
                await _edit(q, text_users(), owner_kb())
            elif arg1 == "weekly":
                await send_weekly_chart(chat_id)
            elif arg1 == "logs":
                await app.send_document(chat_id, log_tail_file())
            return await q.answer()

        # ── settings ────────────────────────────────────────────────────
        if kind == "st":
            bk = f"◀️ {T(uid, 'back')}"
            if arg1 == "lang":
                await _edit(q, header(T(uid, "language"), "🌐") + T(uid, "choose_lang"),
                            language_keyboard("sl:", store.pref(uid, "lang", DEFAULT_LANG), "nav:settings", back_label=bk))
            elif arg1 == "src":
                await _edit(q, header(T(uid, "source"), "🔤") + "Source language of your books:",
                            language_keyboard("ssrc:", store.pref(uid, "src", DEFAULT_SRC), "nav:settings", SOURCE_LANGS, bk))
            elif arg1 == "fmt":
                await _edit(q, header(T(uid, "format"), "📄") + T(uid, "choose_fmt"),
                            format_keyboard("sf:", store.pref(uid, "fmt", DEFAULT_FORMAT), "nav:settings", bk))
            elif arg1 == "spl":
                await _edit(q, header(T(uid, "split"), "✂️") + T(uid, "choose_split"),
                            split_keyboard("ss:", int(store.pref(uid, "split", DEFAULT_SPLIT_KB)), "nav:settings", None, back_label=bk))
            elif arg1 == "bil":
                store.set_pref(uid, "bilingual", not store.pref(uid, "bilingual", False))
                await _edit(q, text_settings(uid), settings_keyboard(uid))
            elif arg1 == "ui":
                store.set_pref(uid, "ui", "hi" if store.pref(uid, "ui", "en") == "en" else "en")
                await _edit(q, text_settings(uid), settings_keyboard(uid))
            return await q.answer()

        if kind in ("sl", "ssrc", "sf", "ss"):
            if kind == "sl" and arg1 in LANGUAGES:
                store.set_pref(uid, "lang", arg1)
            elif kind == "ssrc" and arg1 in SOURCE_LANGS:
                store.set_pref(uid, "src", arg1)
            elif kind == "sf" and arg1 in OUTPUT_FORMATS:
                store.set_pref(uid, "fmt", arg1)
            elif kind == "ss" and arg1.isdigit() and int(arg1) in dict(SPLIT_PRESETS):
                store.set_pref(uid, "split", int(arg1))
            await _edit(q, text_settings(uid), settings_keyboard(uid))
            return await q.answer(T(uid, "saved"))

        # ── re-translate from doc cache ─────────────────────────────────
        if kind == "rt" and arg1:
            return await start_retranslate(q, arg1)

        # ── cancel queued / running job (owner may cancel anyone's) ─────
        if kind == "qx":
            is_owner = store.is_owner(uid)
            if ACTIVE and ACTIVE.job_id == arg1 and (ACTIVE.user_id == uid or is_owner):
                ACTIVE.cancelled_by_owner = is_owner and ACTIVE.user_id != uid
                ACTIVE.cancel.set()
                return await q.answer("🛑 Stopping…")
            for j in QUEUE:
                if j.job_id == arg1 and (j.user_id == uid or is_owner):
                    remove_from_queue(j); j.status = "cancelled"; j.cleanup()
                    await _edit(q, header(T(uid, "cancelled"), "🛑") + summary_block(j), back_home_kb(uid))
                    asyncio.create_task(notify_positions())
                    return await q.answer("Removed from queue")
            return await q.answer("Job already finished.")

        # ── wizard ──────────────────────────────────────────────────────
        if kind in ("jl", "jf", "js", "jc", "jx", "jq", "jgo", "jb", "jo", "jsrc", "jbil", "jp"):
            job = _pending(q, arg1)
            if not job:
                await _edit(q, header(T(uid, "expired"), "⌛") + T(uid, "expired_body"), back_home_kb(uid))
                return await q.answer(T(uid, "expired"), show_alert=True)

            if kind == "jx":
                PENDING.pop(arg1, None); job.cleanup(); USER_STATE.pop(chat_id, None)
                await _edit(q, header(T(uid, "cancelled"), "❌") + T(uid, "discarded", name=b(job.novel_name)), back_home_kb(uid))
            elif kind == "jl":
                if arg2 in LANGUAGES:
                    job.lang = arg2
                await _edit(q, wizard_fmt_text(job), wizard_fmt_kb(job))
            elif kind == "jo":
                await _edit(q, wizard_opts_text(job), wizard_opts_kb(job))
            elif kind == "jsrc":
                if arg2 == "menu":
                    await _edit(q, header(T(uid, "source"), "🔤") + summary_block(job),
                                language_keyboard(f"jsrc:{job.job_id}:", job.src, f"jo:{job.job_id}", SOURCE_LANGS, f"◀️ {T(uid, 'back')}"))
                else:
                    if arg2 in SOURCE_LANGS:
                        job.src = arg2
                    await _edit(q, wizard_opts_text(job), wizard_opts_kb(job))
            elif kind == "jbil":
                job.bilingual = not job.bilingual
                await _edit(q, wizard_opts_text(job), wizard_opts_kb(job))
            elif kind == "jf":
                if arg2 in OUTPUT_FORMATS:
                    job.out_format = arg2
                await _edit(q, wizard_split_text(job), wizard_split_kb(job))
            elif kind == "js":
                if arg2.isdigit() and int(arg2) in dict(SPLIT_PRESETS):            # bug #12
                    job.split_kb = int(arg2)
                await _edit(q, wizard_confirm_text(job), wizard_confirm_kb(job))
            elif kind == "jc":
                USER_STATE[chat_id] = {"state": "await_size", "job_id": arg1}
                await _edit(q, header(T(uid, "custom_size"), "✏️") + summary_block(job) + f"{DIV}\n{T(uid, 'custom_prompt')}",
                            InlineKeyboardMarkup([[btn(f"◀️ {T(uid, 'back')}", f"jb:{arg1}:spl")]]))
            elif kind == "jb":
                USER_STATE.pop(chat_id, None)
                if arg2 == "lang":
                    await _edit(q, wizard_lang_text(job), wizard_lang_kb(job))
                elif arg2 == "fmt":
                    await _edit(q, wizard_fmt_text(job), wizard_fmt_kb(job))
                else:
                    await _edit(q, wizard_split_text(job), wizard_split_kb(job))
            elif kind == "jp":
                await q.answer("👀 …")
                return await run_preview(q, job)
            elif kind in ("jq", "jgo"):
                err = _limits_ok(uid, check_pending=False)
                if err:
                    return await q.answer(re.sub(r"<[^>]+>", "", err)[:190], show_alert=True)
                PENDING.pop(arg1, None)
                USER_STATE.pop(chat_id, None)
                await enqueue(job, q.message)
            return await q.answer()

        await q.answer()
    except Exception as e:                                                   # never leave the spinner hanging
        log.exception("callback error: %s", e)
        try:
            await q.answer("⚠️ Error, try again.", show_alert=False)
        except Exception:
            pass


# ═══════════════════════════════════════════════════════════════════════════
# ⌨️ FREE TEXT  (custom size input · quick text translation)
# ═══════════════════════════════════════════════════════════════════════════
def parse_size_kb(text: str) -> Optional[int]:
    t = text.strip().lower().replace(" ", "")
    mt = re.fullmatch(r"(\d+(?:\.\d+)?)(mb|m|kb|k)?", t)
    if not mt:
        return None
    val = float(mt.group(1))
    unit = mt.group(2) or "kb"
    return int(val * 1024) if unit.startswith("m") else int(val)


@app.on_message(filters.text & PRIVATE & authorized & ~filters.regex(r"^/"))
async def handle_text(_, m: Message):
    uid = m.from_user.id
    chat_id = m.chat.id
    state = USER_STATE.get(chat_id)
    if state and state.get("state") == "await_size":
        size_kb = parse_size_kb(m.text)
        if size_kb is None:
            return await m.reply(T(uid, "invalid_size"))
        if not 50 <= size_kb <= 15360:
            return await m.reply(T(uid, "size_range"))
        job = PENDING.get(state["job_id"])
        USER_STATE.pop(chat_id, None)
        if not job or job.chat_id != chat_id:
            return await m.reply(f"⌛ {T(uid, 'expired_body')}")
        job.split_kb = size_kb
        return await m.reply(wizard_confirm_text(job), reply_markup=wizard_confirm_kb(job))

    text = m.text.strip()
    if len(text) < 2:
        return await m.reply(T(uid, "send_doc_hint", n=QUICK_TEXT_MAX))
    if len(text) > QUICK_TEXT_MAX:
        return await m.reply(f"⚠️ {len(text)} chars — too long for quick mode (max {QUICK_TEXT_MAX}).\n"
                             f"Save it as a .txt file and send it as a document.")
    await quick_translate(m, text, store.pref(uid, "lang", DEFAULT_LANG), store.pref(uid, "src", DEFAULT_SRC))


# ═══════════════════════════════════════════════════════════════════════════
# ▶️ ENTRYPOINT / GRACEFUL SHUTDOWN
# ═══════════════════════════════════════════════════════════════════════════
_RESTART = False
STOP_EVENT: Optional[asyncio.Event] = None


async def graceful_shutdown():
    """Persist queue, tell the active user, stop cleanly (bug #26)."""
    global SHUTTING_DOWN
    if SHUTTING_DOWN:
        return
    SHUTTING_DOWN = True
    if ACTIVE:
        # put the running job back at the head of the queue so it resumes (checkpoint keeps delivered parts)
        ACTIVE.cancel.set()
        QUEUE.insert(0, ACTIVE)
        await LiveMessage(ACTIVE.msg).update(header("Paused", "⏸") + summary_block(ACTIVE) + f"{DIV}\n"
                                              "Bot is restarting — your job will resume automatically.", force=True)
        await asyncio.sleep(1.0)                      # let the worker unwind
    persist_queue()
    store.save(immediate=True)
    try:
        await app.stop()
    except Exception:
        pass


async def main():
    global STARTED_AT, STOP_EVENT
    STOP_EVENT = asyncio.Event()
    loop = asyncio.get_running_loop()
    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            loop.add_signal_handler(sig, STOP_EVENT.set)
        except (NotImplementedError, RuntimeError):     # Windows
            pass
    await app.start()
    STARTED_AT = time.time()
    me = await app.get_me()
    asyncio.create_task(queue_worker())
    asyncio.create_task(janitor())
    asyncio.create_task(health_ping())
    await restore_queue()
    log.info("%s v%s online as @%s | users=%d | pdf_in=%s pdf_out=%s", BOT_NAME, VERSION, me.username,
             len(store.data["users"]), HAS_PDF_IN, HAS_PDF_OUT)
    owner = store.owner_id()
    if owner:
        await safe_send(owner, header("Bot Online", "🟢") +
                        f"@{me.username} v{VERSION} is running.\n"
                        f"👥 Users: {b(len(store.data['users']))}  ·  ⏳ Restored queue: {b(len(QUEUE))}\n"
                        f"📄 PDF in: {b('yes' if HAS_PDF_IN else 'no')}  ·  PDF out: {b('yes' if HAS_PDF_OUT else 'no')}\n"
                        f"🗄 Backup group: {b('set' if BACKUP_GROUP_ID else 'not set')}")
    else:
        log.warning("No owner yet — first user to send the security code becomes owner. "
                    "Security code: %s", store.security_code())
    if not os.getenv("SECURITY_CODE"):
        log.warning("SECURITY_CODE not set — generated one-time code: %s", SECURITY_CODE)
    try:
        await STOP_EVENT.wait()
    except (asyncio.CancelledError, KeyboardInterrupt):
        pass
    await graceful_shutdown()
    log.info("Bot stopped cleanly.")


if __name__ == "__main__":
    try:
        app.run(main())
    except KeyboardInterrupt:
        pass
    finally:
        if _RESTART:
            log.info("Re-executing process…")
            os.execv(sys.executable, [sys.executable] + sys.argv)
