# SKILLS.md

> Reusable, copy-paste-ready techniques for producing premium hand-drawn
> sketchnote infographics from code.
>
> Pair this file with `AGENT.md` (which sets the rules) — `SKILLS.md` shows
> the moves.

---

## Skill 1 — Hand-Drawn Wobble via SVG Turbulence

The single most important trick. `feTurbulence` generates noise; `feDisplacementMap`
uses that noise to physically distort path geometry. The result: clean code,
ink-on-paper output.

```xml
<filter id="rough" x="-2%" y="-2%" width="104%" height="104%">
  <feTurbulence type="fractalNoise" baseFrequency="0.025" numOctaves="2" seed="5"/>
  <feDisplacementMap in="SourceGraphic" scale="1.4"/>
</filter>

<filter id="rough-soft" x="-1%" y="-1%" width="102%" height="102%">
  <feTurbulence type="fractalNoise" baseFrequency="0.018" numOctaves="2" seed="9"/>
  <feDisplacementMap in="SourceGraphic" scale="0.9"/>
</filter>
```

**Tuning guide**

| `scale` | Effect            | Use for                           |
|---------|-------------------|-----------------------------------|
| 0.5–0.9 | Subtle wobble     | Connector lines, dashed arrows    |
| 1.0–1.5 | Medium hand-drawn | Card borders, doodles, headers    |
| 1.5–2.5 | Heavy ink         | Title-block frames, bold stickers |
| > 2.5   | Cartoonish        | Use sparingly                     |

**Rules**
- **Never** apply to body text — destroys legibility. Headings 18 px+ are okay.
- Vary `seed` between filter instances so neighbouring elements look unique.
- Apply via `filter="url(#rough)"` on a `<g>` wrapper, not per-element.

---

## Skill 2 — 3D Title Box (Shadow-Projection Trick)

Stack three filled paths in this draw order: bottom face → right face → front
face. The result reads as 3D without any actual perspective math.

```xml
<!-- Shadow projection (drawn first) -->
<g filter="url(#rough)">
  <path d="M152 100 L716 100 L730 114 L166 114 Z" fill="#2A9D8F"/>   <!-- top  bevel -->
  <path d="M716 100 L730 114 L730 218 L716 204 Z" fill="#2A9D8F"/>   <!-- right side  -->
  <path d="M152 204 L716 204 L730 218 L166 218 Z" fill="#1F7A6F"/>   <!-- bottom edge -->
</g>

<!-- Front face (drawn last, on top) -->
<g filter="url(#rough)">
  <rect x="138" y="100" width="578" height="104" rx="4"
        fill="#FEFEFB" stroke="#1A1A1A" stroke-width="3"/>
</g>
```

The offset `(14, 14)` between front face and shadow controls the apparent depth.

---

## Skill 3 — Reusable Icon Library

Define every doodle **once** in `<defs>` as a `<g id="icon-…">`, then place it
anywhere with `<use href="#icon-…" x=… y=…/>` or `transform="translate(...)"`.

```xml
<defs>
  <g id="icon-router">
    <line x1="10" y1="14" x2="6"  y2="2" class="ink" stroke-width="1.8"/>
    <line x1="40" y1="14" x2="44" y2="2" class="ink" stroke-width="1.8"/>
    <circle cx="6"  cy="2" r="2" fill="#E76F51"/>
    <circle cx="44" cy="2" r="2" fill="#E76F51"/>
    <rect x="2" y="14" width="46" height="22" rx="4"
          fill="#FEFEFB" stroke="#1A1A1A" stroke-width="2"/>
    <!-- ↔ glyph + LEDs -->
  </g>
</defs>

<use href="#icon-router" transform="translate(86,16)"/>
```

Benefits: smaller file, consistent visual language, single place to edit, easy
to scale via `transform="scale(0.7)"`.

**The current set:** see `AGENT.md` §4 for the full inventory.

---

## Skill 4 — Highlighter-Wash Accents

Don't fill shapes with solid accent colors — that buries the ink. Instead, lay
a **low-opacity rectangle** of the accent color *over* the area you want to
"highlight":

```xml
<rect x="46" y="420" width="340" height="22" fill="#E76F51" opacity="0.18"/>
<text x="50" y="437" class="heading-font" font-size="16">"NO CABLES. NO BUDGET."</text>
```

This emulates dragging a marker across paper. Keep opacity in the **0.16–0.30**
range. Anything stronger fights the ink.

---

## Skill 5 — Hatch-Pattern Shading

For the shadow side of 3D objects or shaded callouts, use a `<pattern>` of
diagonal lines instead of a flat fill:

```xml
<pattern id="hatch-teal" width="6" height="6"
         patternUnits="userSpaceOnUse" patternTransform="rotate(45)">
  <line x1="0" y1="0" x2="0" y2="6"
        stroke="#2A9D8F" stroke-width="1.4" opacity="0.55"/>
</pattern>

<rect x="0" y="0" width="100" height="40" fill="url(#hatch-teal)"/>
```

Vary the `patternTransform="rotate(…)"` per color so hatching feels intentional.

---

## Skill 6 — Hand-Drawn Arrows

The default SVG arrowhead looks too clean. Custom marker shaped like a hollow
chevron reads as ink:

```xml
<marker id="arrow" viewBox="0 0 12 12" refX="10" refY="6"
        markerWidth="7" markerHeight="7" orient="auto-start-reverse">
  <path d="M0 0 L12 6 L0 12 L3 6 Z" fill="#1A1A1A"/>
</marker>

<path d="M104 658 Q130 650 156 658"
      stroke="#1A1A1A" stroke-width="1.8" fill="none"
      marker-end="url(#arrow)"/>
```

Use **quadratic Bézier (`Q`)** curves for the arrow body — they look
hand-flicked. Pair with `filter="url(#rough-soft)"` for the wobble.

---

## Skill 7 — Numbered Step Circles

Workflow steps look great as colored-wash circles connected by curving arrows:

```xml
<g filter="url(#rough-soft)">
  <circle cx="80" cy="660" r="22" fill="#FEFEFB" stroke="#1A1A1A" stroke-width="2.2"/>
  <circle cx="80" cy="660" r="22" fill="#2A9D8F" opacity="0.18"/>
  <text x="80" y="667" text-anchor="middle"
        font-family="'Permanent Marker', cursive" font-size="22">1</text>
</g>
```

Alternate accent colors per step (teal → orange → red → orange → teal) to
create visual rhythm.

---

## Skill 8 — Embedded CLI / Code Block

For example labs, embed a faux terminal directly in the SVG:

```xml
<rect x="0" y="0" width="220" height="84" rx="6" fill="#1A1A1A"/>
<rect x="0" y="0" width="220" height="14" fill="#2A2A2A"/>
<circle cx="6"  cy="7" r="2" fill="#C75D5D"/>
<circle cx="14" cy="7" r="2" fill="#E76F51"/>
<circle cx="22" cy="7" r="2" fill="#2A9D8F"/>
<text font-family="'Courier New', monospace" font-size="8.5" fill="#FEFEFB">
  <tspan x="6" y="28">R1&gt; enable</tspan>
  <tspan x="6" y="41">R1# config t</tspan>
</text>
```

The 3-dot "macOS window controls" + dark fill instantly read as "terminal".

---

## Skill 9 — Speech / Result Bubbles

Tail-equipped speech bubbles celebrate successes (e.g. ping replies):

```xml
<g id="icon-speech">
  <path d="M4 4 Q4 0 8 0 L72 0 Q76 0 76 4 L76 22 Q76 26 72 26
           L24 26 L16 34 L18 26 L8 26 Q4 26 4 22 Z"
        fill="#FEFEFB" stroke="#1A1A1A" stroke-width="2"/>
</g>
```

Combine with green text (`fill="#2A9D8F"`) for success states.

---

## Skill 10 — GitHub-Compatible Self-Contained SVGs

GitHub's CSP **strips `@import` and external `<link>`** inside raw SVGs. Always
specify cursive **system fallbacks** so the file still reads as hand-drawn when
viewed directly on github.com:

```xml
<text font-family="'Permanent Marker', 'Marker Felt', 'Impact',
                   'Comic Sans MS', cursive">…</text>
```

For the *real* hand-drawn fonts (Permanent Marker, Caveat, Architects
Daughter), provide an `index.html` viewer that loads Google Fonts:

```html
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin/>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=
  Architects+Daughter&family=Caveat:wght@400;700&family=Permanent+Marker
  &display=swap"/>

<object data="my-sketchnote.svg" type="image/svg+xml"></object>
```

Loaded fonts cascade into the embedded `<object>`'s SVG.

---

## Skill 11 — Print-to-PDF Action

A single line in the viewer saves users a trip to an export tool:

```html
<a href="javascript:window.print()">Print / Save PDF</a>
```

Pair with print-only CSS that hides chrome and removes shadows:

```css
@media print {
  body { background: white; }
  header, footer { display: none; }
  main { margin: 0; padding: 0; max-width: none; }
  .paper { box-shadow: none; }
}
```

Set the user's print dialog to **A4 portrait, margins: None** — and the SVG's
`viewBox="0 0 800 1131"` fills the page exactly.

---

## Skill 12 — Visual Hierarchy Cheat Sheet

A working scale that keeps a dense A4 sketchnote readable:

| Role            | Size | Font               | Weight |
|-----------------|------|--------------------|--------|
| Page title      | 46 px | Permanent Marker   | —      |
| Section header  | 22 px | Permanent Marker   | —      |
| Sub-heading     | 18 px | Caveat             | 700    |
| Body            | 13 px | Architects Daughter| —      |
| Small / labels  | 11 px | Architects Daughter| —      |
| Captions / tags |  9 px | Architects Daughter| —      |

Anchor on these sizes — when adding new sections, reach for an existing tier
before inventing a new one.

---

## Skill 13 — Negative-Space Discipline

A premium sketchnote feels dense **without** crowding. Heuristics:

- 30–40 px of breathing room between top-level sections.
- 14–18 px padding inside cards.
- Never extend strokes within 6 px of a card edge.
- If a section feels packed, drop a tile rather than shrink fonts.

---

## Skill 14 — Decorative "Stickers"

Tiny rotated colored chips break up grid rigidity:

```xml
<g transform="translate(745,558) rotate(12)" filter="url(#rough)">
  <rect x="0" y="0" width="48" height="22" fill="#2A9D8F"/>
  <text x="24" y="16" text-anchor="middle"
        font-family="'Permanent Marker', cursive" font-size="12"
        fill="#FEFEFB">$0</text>
</g>
```

Place 1–3 per page — any more and they fight for attention.

---

## Skill 15 — Connector Lines from Title (Radial Layout)

Hand-drawn dashed connectors from the title to top-row sections sell the
"radial" composition:

```xml
<g filter="url(#rough-soft)">
  <path d="M260 235 Q230 250 215 285"
        stroke="#1A1A1A" stroke-width="1.4" fill="none"
        stroke-dasharray="3 4"/>
</g>
```

Use `stroke-dasharray="3 4"` for short connectors, `2 5` (with `opacity="0.5"`)
for longer "trail-off" lines to far sections.
