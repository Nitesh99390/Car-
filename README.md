# 📚 NovelTranslator PRO — Telegram Document Translation Bot (v3.0)

Single-file Telegram bot (`bot.py`) that translates whole novels / documents
(EPUB · TXT · DOCX · PDF · HTML · FB2 · MD) into 24 languages and returns them
as TXT · DOCX · EPUB · HTML · PDF — with chapter detection, glossary
name-locking, SQLite translation cache, fair queue, resume-after-restart and an
owner admin panel.

## Features

| Area | Highlights |
|------|------------|
| **Translation** | Google → LibreTranslate → MyMemory fallback · glossary / name-lock · SQLite cache · source-language override · sample preview · bilingual output · junk-line filter |
| **Files** | In: `epub txt docx pdf html htm fb2 md` · Out: `txt docx epub html pdf` · size-based split (on *output* size) · ZIP merge · resume checkpoints · re-translate from cache |
| **Queue** | fair round-robin · owner priority · per-user caps · daily quota · ETA · `/history` · persisted across restarts |
| **Owner** | `/ban /unban /adduser /deluser /setcode /logs /restart /cancelall /priority /weekly /links /broadcast` · health ping · forum-topic backups |
| **UX** | EN + HI interface · quick inline text translation · mobile-friendly menus · custom split size |

## Quick start

```bash
git clone <this repo> && cd <repo>
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

cp .env.example .env        # fill in API_ID / API_HASH / BOT_TOKEN / SECURITY_CODE / OWNER_ID
python bot.py
```

Users unlock the bot by sending the `SECURITY_CODE` once. The first user to
do so becomes owner unless `OWNER_ID` is set (recommended).

### Docker

```bash
cp .env.example .env && $EDITOR .env
docker compose up -d --build
docker compose logs -f
```

Runtime data (inbox, outputs, DB, cache, logs) lives in `./data`.

## Configuration (`.env`)

| Variable | Required | Description |
|----------|----------|-------------|
| `API_ID`, `API_HASH` | ✅ | From <https://my.telegram.org> |
| `BOT_TOKEN` | ✅ | From @BotFather |
| `SECURITY_CODE` | recommended | Access code users must send once (random one is generated & logged if unset) |
| `OWNER_ID` | recommended | Telegram user id of the owner |
| `BACKUP_GROUP_ID` | optional | Forum-enabled supergroup where all inputs/outputs are archived (one topic per job) |
| `PDF_FONT_PATH` | optional | Override TTF used for PDF output |
| `LIBRE_URL` | optional | Self-hosted LibreTranslate endpoint |
| `WORKDIR` | optional | Data directory (default `workspace_textbot`) |

### PDF output fonts

Noto script fonts only contain their own script (e.g. `NotoSansDevanagari` has
no Latin glyphs). The bot therefore auto-picks a per-language Noto font and
falls back to `NotoSans` / `DejaVuSans` for Latin runs inside the same
paragraph. Install `fonts-noto-core` (Debian/Ubuntu) or set `PDF_FONT_PATH`.

## Commands

**Everyone:** `/start /help /settings /queue /cancel /mystats /history /glossary /tr`  
**Owner:** `/stats /users /adduser /deluser /ban /unban /setcode /broadcast /cancelall /priority /logs /weekly /restart /links`

Just send a document to start the wizard, or send plain text (≤ 3000 chars)
for an instant translation.

## Testing

Offline smoke tests (no Telegram / network needed):

```bash
python tests/test_core.py
python tests/test_engine.py
python tests/test_pdf.py
```

## Changelog

**v3.0** — see header of `bot.py` for the full list of ~30 bug fixes and new
features over v2.0.

**v3.0.1 (this repo)**
- PDF: per-language font selection + Latin fallback → mixed-script text
  (e.g. `Part 1 of 3 — नॉवेल`) no longer renders as blank boxes.
- FB2: parse with a real XML parser; `<title><p>…</p></title>` no longer leaks
  literal `<p>` tags into the output, headings no longer duplicated.
- Silence `XMLParsedAsHTMLWarning` noise in logs.
- Added `requirements.txt`, `.env.example`, `.gitignore`, `Dockerfile`,
  `docker-compose.yml`, smoke tests.
