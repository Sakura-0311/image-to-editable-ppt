# Acceptance Records

Public verification log for this skill's bundled tools and demo assets. Each entry states what was actually run and observed — structural checks verify structure only; they are not claimed as generation-quality scores.

## 2026-11-26 — Bundled tools (hygiene.py / audit.py)

Method: a deliberately dirty 6-slide sample was built (empty author, zeroed statistics, single-quoted XML declaration, missing font slots, script-style shape names `card_01` / `!!hero_jar`), then audited, repaired, and re-audited.

- `audit.py` on the dirty sample: **4 BLOCK + 6 WARN**, exit code 1 (delivery correctly refused).
- `hygiene.py` (seed 42, delivery time pinned to 2026-11-26T15:30): created `2026-11-24T10:47:34`, modified `2026-11-24T11:38:37` — a 51-minute span inside the 09:00–23:00 window, non-round minutes/seconds; `TotalTime` 24 minutes (47% of span); revision 3. All within the Human-Plausible Timestamps rules.
- Re-audit after hygiene: **0 BLOCK**. The deck opens and converts to PDF in LibreOffice without error.
- Boundary cases: delivery at 08:15 and at 23:50 both keep timestamps inside the allowed window; a 40-slide deck yields a 366-minute span and revision 18 (scales with deck size, stays under the ~8-hour cap).
- Known behaviour, by design: `hygiene.py` never renames shapes (renaming could break Morph pairing), so decks with `!!`-style names keep an audit WARN until renamed manually.

## 2026-11-26 — Motion monotony check (audit.py MOTION_MONOTONY)

Method: transitions were injected into copies of a real 12-slide deck via slide-XML edits.

- 5 slides all carrying Morph transitions → audit reports the motion distribution `{morph: 5}` and raises **WARN MOTION_MONOTONY**, as designed.
- 5 slides with mixed types (morph ×2, fade, push, wipe) → distribution reported, **no warning**.
- Both files open and convert in LibreOffice without error.

## 2026-11-29 — Font portability check (audit.py FONT_PORTABILITY) + real-machine render test

Method: the showcase deck was opened on a real Windows machine via a COM bridge (harness-anything `cli-anything-wps`) and exported to PDF by the real application; a LibreOffice render of the same file served as the control.

- Real-application render vs LibreOffice: layout skeleton identical (positions, rules, stat column), **but every glyph was substituted** — the deck's fonts (Playfair Display, Source Sans 3, IBM Plex Mono) are not installed on that machine, and the app fell back to a handwriting-style font for the whole page. The LibreOffice render had silently approximated the intended serif/sans/mono look, so LibreOffice-only QA could not see this failure.
- New check: `audit.py` raises **WARN FONT_PORTABILITY** listing every used font that is not preinstalled on typical Windows/macOS machines, with delivery guidance (install the fonts on the target machine, or switch to Tier A preinstalled-safe fonts).
- Verification: the showcase deck → WARN lists exactly `IBM Plex Mono, Playfair Display, Source Sans 3`. A copy of the same deck with fonts swapped to Georgia / Calibri / Consolas → no FONT_PORTABILITY finding.
- Note: the two PDFs returned from the real-machine test (bridge export vs manual export) rendered pixel-identically, so they count as one application's evidence, not two independent ones.

## 2026-11-26 — Demo video

- File: `image-to-editable-ppt-demo.mp4`, 77 seconds, 1920×1080, H.264/AAC, English narration and subtitles.
- Verified by frame inspection at 3s / 15s / 30s / 40s / 48s / 60s / 74s: hook giveaways, object-selection handles, the five human touches, the file-history comparison (3-second AI span vs 5 h 33 min human span, TotalTime 118 min, revision 9 — internally consistent), and the end card. No garbled text or misspellings found.
- The English-trigger variant (`-en`) replaces only the end-card segment; the splice was frame-checked at 74s.
