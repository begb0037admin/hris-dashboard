# HRIS Team — Open Tickets Dashboard
## Design Reference — Mockup v2 (LOCKED)

Locked: 2026-07-02  
Locked by: Kevin Lelitte  
Status: **Approved baseline — ready for production build**

---

## Artifact URL

```
https://claude.ai/code/artifact/e18d06d6-8c3e-4a87-924b-bf1c0a94cbf4
```

This is the canonical reference for the approved v2 design. Open this URL to view the locked mockup.

---

## Locked Design Specification

### Layout
- **Left sidebar:** 310px, flat Oxford navy `#002147` (no gradient)
- **Main column:** Analyst card list (scrollable)
- **Right column:** Linda AI panel
- **Font:** Inter throughout

### Sidebar contents (top to bottom)
- University of Oxford crest + "HRIS DASHBOARD" wordmark
- Live clock (HH:MM:SS) + date
- **Queue Summary tile:**
  - Three stats only: Total / Assigned / Unassigned
  - Secondary rows: Stale (>30 days) / Oldest ticket / Unassigned (with coloured values)
- Print Report button
- **Links section:** OSM — Oxford Service Manager, Command Centre, HRIS Launcher, FA Knowledge Base, Access Group Support

### Analyst cards (main column)
One card per analyst. Order: FA Team Lead group first, then others.
Cards confirmed in mockup:
- Kevin (FA Team Lead — shown with 3 open tickets)
- James
- Asta
- Michael
- Simon Burford (Senior Manager — shown with 1 open ticket)

Each card contains:
- Analyst name + role label
- Ticket table: TICKET / SUMMARY / STATUS / PRI / DAYS columns
- Status pills: Waiting (amber), Assigned (green), Accepted (blue), etc.
- DAYS column in red when aged

### Linda AI panel (right column)
- Header: "Analysis" with timestamp
- Auto-generated summary per analyst section
- Warning triangles for aged/at-risk items
- Info bullets for notable tickets
- "ASK LINDA" divider
- Chat input: "Ask Linda about your tickets..." + Send button
- Kevin's message bubble (dark) / Linda's reply (light)

---

## What Is NOT Yet Built

The mockup is static. The production build requires:

1. **Live data source** — OSM API or exported ticket feed (TBC with Kevin)
2. **Auth** — access control for the dashboard (TBC)
3. **Linda AI integration** — Cloudflare Worker + Claude API (same pattern as hr-kb-ai worker)
4. **Auto-refresh** — ticket data refresh interval (TBC)
5. **Print/export** — Print Report button implementation

---

## Production Build — Outstanding Decisions (Kevin to confirm)

| # | Decision | Status |
|---|---|---|
| 1 | Data source for live tickets (OSM API / CSV export / manual JSON?) | TBC |
| 2 | Refresh interval | TBC |
| 3 | Authentication / access restriction | TBC |
| 4 | Hosting — hris-dashboard GitHub Pages or new repo? | TBC |
| 5 | Linda AI — use existing hr-kb-ai Cloudflare Worker or new worker? | TBC |

---

## Session History

| Date | Event |
|---|---|
| 2026-07-02 | Mockup v1 produced (HTML — process violation; corrected by Kevin) |
| 2026-07-02 | Mockup v2 produced as Claude Artifact (correct process) |
| 2026-07-02 | v2 approved and locked by Kevin |
| 2026-07-02 | This design reference document committed to hris-dashboard repo |

---

## Governance

- Mockup produced as Claude Artifact per CONSTITUTION.md Section 11
- Production build requires Kevin approval before any commit to repo
- Follow AGENT_MODEL.md v2.5 and CONSTITUTION.md v2.1 for all build work
