# AGENT_MODEL.md
# Runtime Operating Model

Version : 1.1
Status  : Ratified
Updated : 2026-06-18
Author  : Kevin Lelitte, HR Systems, University of Oxford

Governed by: CONSTITUTION.md
Scope      : All work repositories (begb0037admin)

---

## Preamble

This document defines the current implementation of the four-role
model established in CONSTITUTION.md Section 1. It assigns tools
and software to roles, defines dispatch mechanics, and records
platform context.

This document changes more frequently than the constitution. When
the tooling changes, this document is updated. The constitutional
principles do not change with it.

If this document conflicts with CONSTITUTION.md, the constitution
wins. See CONSTITUTION.md Section 6.

---

## Section 1 — Platform Context

Two machines are in scope for work operations.

**Work machine (Kevin)**
Operator : Kevin Lelitte
OS       : Windows
Username : begb0037 | Domain: AD-OAK
Path root: C:\Users\begb0037.AD-OAK\

**Personal machine (Hope)**
Operator : Hope (personal domain)
OS       : macOS
Scope    : AIMM and personal projects only — out of scope for
           all work repositories

GitHub is the authoritative source of truth for all governed
repositories and acts as the shared storage layer between machines.
Repository content is authoritative; local copies are working
copies only.

Never hardcode machine-specific paths in repository files. Use
GitHub URLs as the stable reference wherever possible.

---

## Section 2 — Role Assignments

**Seat A — Reasoning Seat → Claude Chat**
Thinks, plans, architects, and routes. All sessions begin here.

**Seat B — Human Seat → Kevin (work) / Hope (personal)**
Executes read-only terminal commands on instruction from Seat A.
Human authority supersedes all in-flight decisions when invoked.

**Seat C — Execution Seat → Claude Code**
The sole seat authorised to implement approved changes. Acts
only on complete, explicit briefs from Seat A.

**Seat D — Verification Seat → Chrome (browser)**
Confirms live behaviour. Read-only. Reports what it observes.

---

## Section 3 — Dispatch Protocol

Dispatches are issued by Seat A only. Strictly sequential.

🔵 RUN SCRIPT   → Seat B (read-only terminal)
🟡 COWORK BRIEF → Seat C (implementation)
🔴 CHROME BRIEF → Seat D (verification only)

---

## Section 4 — Session Discipline

1. Large output → write to file, never paste to chat.
2. No session closes without documentation updated.
3. After a sustained session, proactively offer a handover brief.

---

## Section 5 — Cross-Domain Model

**Kevin** — work domain. All Oxford HR Systems repositories.
**Hope** — personal domain only.
**Failover chain (work):** Kevin → Hope

---

## Section 6 — GitHub Access

URL pattern : https://api.github.com/repos/begb0037admin/
              {repo}/contents/{path}?ref=main
Auth header : Authorization: token {PAT}

Secrets never committed to any file.

---

## Section 7 — Repository Scope

| Repository           | Status         | Notes                    |
|----------------------|----------------|-------------------------|
| clockify             | Active         | Gold standard / template |
| hris-dashboard       | Active         |                          |
| hr-fa-knowledge-base | Active         |                          |
| work-inbox           | Active         |                          |
| meeting-records      | Active         |                          |
| hr-projects          | Active         |                          |
| command-centre       | Active         |                          |
| ag-flexpoints        | Active         |                          |
| hris-launcher        | Active         |                          |
| hris-change-requests | Active         |                          |
| aimm                 | Out of scope   | Personal domain — Hope  |
| personal-finance     | Out of scope   | Personal domain — Hope  |

---

## Version History

| Version | Date       | Change                              |
|---------|------------|-------------------------------------|
| 1.0     | 2026-06-06 | Initial ratification.               |
| 1.1     | 2026-06-18 | Repo scope updated. Seat C updated  |
|         |            | to Claude Code.                     |
