# CONSTITUTION.md
# The Operating Constitution

Version : 1.0
Status  : Ratified
Ratified: 2026-06-06
Author  : Kevin Lelitte, HR Systems, University of Oxford

---

## Preamble

This document defines the enduring principles under which all work
is conducted across all repositories and projects. It is not a
description of current tools, workflows, or conventions. Those live
in AGENT_MODEL.md and CLAUDE.md respectively.

When this document conflicts with any other document, Section 6
governs.

---

## Section 1 — The Separation of Concerns

The system operates across four distinct roles. Each role has a
defined authority boundary.

1. **The Reasoning Seat** — thinks, plans, designs, and routes.
2. **The Human Seat** — executes read-only operations on instruction. Human authority supersedes all in-flight decisions when invoked.
3. **The Execution Seat** — the sole authority for implementing approved changes.
4. **The Verification Seat** — confirms live behaviour. Read-only.

---

## Section 2 — Dispatch Quality

A dispatch must be complete enough that the receiving role requires
no architectural decisions to carry it out. An incomplete dispatch
is a reasoning seat failure.

One dispatch at a time.

---

## Section 3 — Implementation Authority

One role, and only one role, implements approved changes.
If any other role believes a change is needed, the output is a
handoff — not an implementation attempt.

---

## Section 4 — Rollback Before Change

Before any change, a restore point must be recorded. If the change
fails, restore immediately — do not attempt a fix.

---

## Section 5 — Documentation Permanence

Conversation is temporary. Documentation is permanent.

No session closes without documentation updated to reflect current
state.

---

## Section 6 — The Source of Truth Hierarchy

1. The operator's current AI preferences
2. This constitution
3. AGENT_MODEL.md
4. CLAUDE.md
5. STATUS.md / HANDOVER.md

---

## Section 7 — Amendment Process

Amendments require: documented reason, session note, version bump,
propagation to all repositories carrying this file.

---

## Section 8 — Universal Applicability

This constitution applies to every repository, every project, and
every session regardless of technology stack, workflow, or tooling.

---

## Version History

| Version | Date       | Change               |
|---------|------------|----------------------|
| 1.0     | 2026-06-06 | Initial ratification |
