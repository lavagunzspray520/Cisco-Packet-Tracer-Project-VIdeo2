---
inclusion: always
---

# Project Steering — Cisco Packet Tracer Sketchnotes

This repository produces hand-drawn-style infographics about Cisco Packet Tracer.

**Before doing anything, read these two files:**

#[[file:../../AGENT.md]]

#[[file:../../SKILLS.md]]

`AGENT.md` defines the rules (color palette, typography, layout, branching).
`SKILLS.md` provides copy-paste-ready SVG technique recipes.

**Quick reminders:**

- Format: A4 portrait SVG, viewBox `0 0 800 1131`.
- Palette: ink `#1A1A1A`, paper `#FEFEFB`, teal `#2A9D8F`, orange `#E76F51`,
  muted red `#C75D5D`, yellow wash `#F4C95D`. **Do not introduce other hues.**
- Hand-drawn feel: SVG `feTurbulence` + `feDisplacementMap` filter on
  borders / doodles / arrows (never on body text).
- Typography: Permanent Marker / Caveat / Architects Daughter, all caps for
  titles and headings.
- Always wrap each new SVG in an `index.html` viewer that loads Google Fonts
  and exposes a Print-to-PDF action.
- Branch naming: `sketchnote/<topic-slug>` for new artifacts;
  `chore/<short>` for meta / docs.
- Never push to `main` &mdash; always open a PR.
