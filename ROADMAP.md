# HRIS Dashboard — Roadmap

## Shipped

### GitHub Actions Migration — COMPLETE (4 June 2026)

Migrated from Windows Task Scheduler to GitHub Actions with a self-hosted runner on DESKTOP-MJDJM64.

- Self-hosted runner installed as Windows service at C:\\actions-runner\\
- Workflow updated: runs-on self-hosted, Python setup removed, commit step in PowerShell
- Dashboard confirmed working from home without VPN (Oxford MDM-enrolled)
- Runner auto-starts on boot

### Linda AI Integration — COMPLETE (03 July 2026)

Added Linda AI panel (right column) powered by `hr-kb-ai.kevinlelitte.workers.dev` (shared Cloudflare Worker with HR FA Knowledge Base).

- Auto-analyses full queue on page load: overall summary + per-analyst breakdown with dividers
- Coloured insight cards: ✦ Oxford blue (summary), ⚠ red (warnings), ✓ green (ok), ℹ grey (info)
- Input unlocks after auto-analysis — Kevin can paste ticket threads for deep-dive / KB lookups
- CORS configured for both `kb.lelitte.co.uk` and `begb0037admin.github.io`
- Simon Burford excluded from display and analysis

### UI Polish — COMPLETE (03 July 2026)

- Inter font throughout
- Analyst role titles in card headers and Linda dividers
- Linda panel widened to 450px
- Header subtitle Oxford blue
- Ticket numbers from `task_num` field
- Unified colour palette across Linda cards and status pills
- Improved table typography

## Ongoing

- Dashboard auto-updates hourly 8am–5pm via GitHub Actions (self-hosted runner DESKTOP-MJDJM64)
- Session refresh via `Refresh Session.bat` when Oxford SSO expires (~every few weeks)
- Linda AI live on every page load

## Backlog

### NEXT: Linda Voice — TTS + STT

Give Linda a voice, matching the implementation already working in the HR FA Knowledge Base.

- **STT (Speech-to-Text):** Mic button → Scribe v2 via `/stt` route on `hr-kb-ai` Worker → transcribed text inserted into Linda input
- **TTS (Text-to-Speech):** Linda responses read aloud via ElevenLabs `/tts` route on same Worker
- **Worker:** No new Worker needed — routes already exist on `hr-kb-ai.kevinlelitte.workers.dev`; just wire into Linda's UI
- **Secrets:** `ELEVENLABS_API_KEY` already stored in the Worker — nothing new to configure in Cloudflare
- **Reference:** `begb0037admin/hr-fa-knowledge-base/index.html` — copy mic + listen button pattern from there

### Other

- [ ] Move to `oms.lelitte.co.uk` — Kevin to configure domain + Cloudflare Pages
- [ ] Update Cloudflare `ALLOWED_ORIGIN` when domain is live (remove GitHub Pages entry)
- [ ] Ticket click-through / expand to open in OSM directly (post-domain move)
- [ ] Consider scheduled Linda briefing export (daily summary email or Command Centre task)

---
Last updated: 03 July 2026 — Lelitte Co.
