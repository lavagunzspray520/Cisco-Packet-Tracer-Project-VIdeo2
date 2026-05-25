# AGENT.md

> Persistent context for AI coding agents (Kiro, Claude Code, Cursor, Codex, Aider, etc.)
> entering this repository. Read this first before making changes.

---

## 1. Repository Purpose

This repository hosts **visual learning artifacts** about **Cisco Packet Tracer**.
It is *not* a software project — there is no build system, package manifest, or
runtime. Deliverables are static, hand-drawn-style infographics shipped as SVG +
HTML viewers.

The primary audience is learners (CCNA / NetAcad students), educators, and
visual thinkers. Output should always feel like a **premium graphic-recording
session** — dense with insight, never cluttered.

---

## 2. Repository Map

```
Cisco-Packet-Tracer-Project-VIdeo2/
├── README.md                                    quick links to artifacts
├── AGENT.md                                     this file
├── SKILLS.md                                    reusable visual techniques
└── sketchnote/
    ├── cisco-packet-tracer-sketchnote.svg       A4 portrait infographic
    └── index.html                               viewer + Print-to-PDF action
```

When new infographics are added, follow the same pattern:
`<topic-slug>/<topic-slug>-sketchnote.svg` plus an `index.html` viewer.

---

## 3. House Style — Sketchnote Visual Identity

These rules are **non-negotiable** unless the user explicitly overrides them.

### Format
- A4 portrait, viewBox `0 0 800 1131` (≈ 96 DPI).
- Single self-contained SVG — no external image references.
- Wrap SVG in an `index.html` viewer that loads Google Fonts and exposes a
  Print / Save-as-PDF action.

### Color Palette (strict)
| Role          | Hex      |
|---------------|----------|
| Ink / outline | `#1A1A1A` |
| Paper         | `#FEFEFB` |
| Teal accent   | `#2A9D8F` |
| Orange accent | `#E76F51` |
| Muted red     | `#C75D5D` |
| Yellow wash   | `#F4C95D` (highlighter only, ≤ 30% opacity) |

Do **not** introduce other hues. Use accent colors as low-opacity washes
(`opacity: 0.16–0.30`) layered behind black ink — never as solid fills behind
text.

### Typography
- Titles: `Permanent Marker` (fallback: `Marker Felt`, `Impact`, `Comic Sans MS`).
- Section headings: `Caveat` 700 (fallback: `Bradley Hand`, `Comic Sans MS`).
- Body: `Architects Daughter` (fallback: `Caveat`, `Comic Sans MS`).
- Monospace (CLI snippets only): `Courier New`.
- All headings and titles in **ALL CAPS**.
- Type scale: 46 / 22 / 18 / 13 / 11 / 9 px for title / section / sub / body /
  small / caption.

### Hand-Drawn Feel
- Apply an `feTurbulence` + `feDisplacementMap` filter (`#rough`) to all
  structural shapes, doodles, and arrows.
- Do **not** apply the displacement filter to body text — it hurts legibility.
- Vary `seed` between filter instances so adjacent elements look unique.
- Stroke widths 1.5–3 px, `stroke-linecap="round"`, `stroke-linejoin="round"`.

### Layout
- Bold central 3D-style title box (front face + offset shadow projection).
- Content arranged radially / in a 2-column grid below the title.
- Hand-drawn dashed connectors from title to top-row sections.
- One full-width example block (diagram + CLI + result speech bubble).
- Final bottom banner with stats and an anchor mantra.

### Section Template (default)
1. Title 3D box + decorative starbursts / speech bubbles.
2. Definition strip / what is it?
3. Key features grid (6 dashed-tile icon cells).
4. Workflow (numbered colored circles + curving arrows).
5. Why it matters (4 benefit pairs with icons).
6. Concrete example (network diagram + CLI snippet + ping success bubble).
7. Pro tips + by-the-numbers stats card.
8. Anchor line: `~ VERB · VERB · VERB · VERB ~`.

---

## 4. Reusable Icon Library

All doodles are defined once in `<defs>` as `<g id="icon-…">` and re-used via
`<use href="#icon-…"/>`. When extending, add new icons to `<defs>` rather than
inlining.

Existing icon set:
- Network gear: `icon-router`, `icon-switch`, `icon-pc`, `icon-server`
- Humans: `icon-stick`, `icon-people`
- Connectivity: `icon-cloud`, `icon-globe`, `icon-wifi`
- Concepts: `icon-bulb`, `icon-shield`, `icon-dollar`, `icon-eye`, `icon-cert`
- Packet-Tracer-specific: `icon-cli`, `icon-iot`, `icon-wand`, `icon-topology`
- Decoration: `icon-star`, `icon-check`, `icon-speech`, `icon-box`

See `SKILLS.md` for the implementation patterns.

---

## 5. Workflow Conventions

### Branching
- Never commit directly to `main`.
- Visual artifacts: `sketchnote/<topic-slug>`.
- Meta / docs / chores: `chore/<short-description>`.

### Commits
- Conventional style preferred: `feat: …`, `chore: …`, `docs: …`.
- Body of commit message lists what changed and why.

### Pull Requests
- Always link to where the artifact can be viewed (GitHub renders SVGs
  directly — link the SVG file itself).
- Include a short description of the section structure.

### Local Preview
GitHub strips external fonts from raw SVG (CSP). For the full hand-drawn font
appearance:
1. Open `<topic>/index.html` in a desktop browser (Google Fonts will load).
2. Use **Print → Save as PDF**, paper size **A4 portrait**, margins **None**,
   for a print-ready export.

---

## 6. What This Repo Is *Not*
- Not a Cisco Packet Tracer plugin or .pkt file collection.
- Not a tutorial in markdown form — the entire pedagogy is the visual.
- Not a generic infographic factory — every artifact must teach something
  about networking, Cisco, or Packet Tracer specifically.

---

## 7. Chat-History Highlights (Memory)

Recorded so future sessions don't re-litigate established decisions.

- **2026-05-25 — PR #1: Cisco Packet Tracer visual field guide.**
  Established the entire house style: A4 portrait SVG, three-color palette
  (teal / orange / muted-red), turbulence-displacement hand-drawn filter,
  3D shadow-projection title box, 7-section template (title → definition →
  features → workflow → why → example lab → pro tips), reusable icon
  library in `<defs>`, HTML viewer with Google Fonts + Print-to-PDF action.
  Anchor mantra: *Practice · Break · Learn · Repeat.*

- **2026-05-25 — Agent memory files added.**
  This file (`AGENT.md`) and `SKILLS.md` were added to persist context
  across future agent sessions.

---

## 8. How to Extend

When the user asks for a new infographic on a related topic
(e.g. "subnetting", "OSI model", "VLAN troubleshooting"):

1. Read `SKILLS.md` first for the technique recipes.
2. Re-use the existing `<defs>` block (icons, filters, patterns, markers).
3. Follow the section template in §3 above.
4. Drop new artifact under `<topic-slug>/<topic-slug>-sketchnote.svg`.
5. Add an `index.html` viewer in the same folder (copy from `sketchnote/`).
6. Update `README.md` with a link.
7. Open a PR on a branch named `sketchnote/<topic-slug>`.
