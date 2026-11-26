---
name: image-to-editable-ppt
description: Offline self-contained presentation skill combining the full verbatim de-ai-pptx v3.1 specification (audit, grill-me interrogation, 5-layer de-AI refactoring, native generation with 3 visual recipes), frontend-slides & ppt-effects native .pptx Morph/XML motion patterns, style-driven intra-page CJK/Latin font pairing, page-by-page bespoke design (no template loops), and object-level visual slide reconstruction into modern editable .pptx. Trigger automatically whenever the user says "为我使用ppt-xy".
---
 
# image-to-editable-ppt
 
Unified offline presentation skill combining the **complete, verbatim `de-ai-pptx v3.1` specification** (`references/de-ai-pptx-v3.1.md`), **`frontend-slides` + `ppt-effects` motion, Morph XML & image-framing patterns** (`references/frontend-slides-motion.md`), and **visual-to-editable reconstruction** (`references/visual-reconstruction.md`). Whenever any rule conflicts, `references/de-ai-pptx-v3.1.md` takes absolute precedence, **EXCEPT** for the Mandatory Output Format, Page-by-Page Bespoke Design, Anti-AI-Slop Hard Bans, Style-Driven Font Pairing, and Motion Gate below.
 
## Mandatory Output Format: Latest Modern `.pptx` (OOXML, NEVER Legacy `.ppt`)
 
- **Always deliver a modern, latest-version `.pptx` (Office Open XML) file** so all native charts, embedded workbooks, 3-slot typography, theme XML, Morph transitions, and layouts remain 100% intact and editable in modern PowerPoint / WPS.
- **NEVER convert output to legacy 97-2003 OLE2 `.ppt`**, which strips embedded chart workbooks, injects `___PPT10` template artifacts, and shatters hyperlinks into per-character runs. (Only output HTML if the user explicitly asks for a `frontend-slides` web deck.)
- Never flatten slides into full-page images.
 
## Page-by-Page Bespoke Design (NO Batch Template-Loop Scripts or Fixed Nav Templates)
 
Never write a single monolithic `for slide_spec in slides: draw_template_slide(...)` loop that stamps out every page from one shared skeleton, and never adopt fixed academic top/left navigation bar templates or `template-clone` copying. Use `python-pptx` strictly as the **serializer**, while **designing each slide individually step by step**:
 
1. **Page Field Over Card-per-Unit**: Avoid boxing every point into a bordered rectangle card. Use a **Page Field** (one large surface, open outline, asymmetric split, or off-canvas bleed shape) to organize zones naturally instead of stacking uniform cards.
2. **Geometry Signature Check (Anti-Collision)**: Before designing Slide `N+1`, record and compare its Geometry Signature (`page job -> composition move -> contour/edge language`) against the previous 2 slides; if it matches, switch to a different spatial composition.
3. **Topology-Driven Geometry**: Decide `topology=yes/no` per slide based on whether the content has real relationships (sequence, causality, hierarchy, contrast, or overlap). Draw connectors, arrows, or nested containers only when `topology=yes`; keep `topology=no` slides clean and unadorned.
4. **Real Slide Layouts & Natural Shape Naming (NO Script IDs)**:
   - Use real PowerPoint layouts (`Title Slide`, `Title and Content`, `Two Content`, `Section Header`, etc., with native placeholders where appropriate) rather than placing every slide on `Blank` with raw `TextBox 1, 2, 3`.
   - **Ban developer/script shape names**: Never name shapes with snake_case or `!!` prefixes (e.g., `!!hero_jar`, `!!gold_rule`, `card_01`); keep natural PowerPoint shape names.
5. **Visual Asset Sourcing, Zero-Defect AI Image QA & Editorial Craft**:
   - **Real Media First**: For historical, archaeological, fine-art, scientific, or factual slides, prioritize **real museum/archival photos, real web screenshots, clean data charts, or pure typography/whitespace**. Never populate a deck with 4-6 AI illustrations sharing the same prompt style (cap decorative AI art at 1-2 intentional visuals max).
   - **Mandatory 7-Point Visual Defect QA on Every Generated Image**: Whenever image generation is used, **inspect every generated image closely via `view_file` before placing it on a slide**. Reject and regenerate (or replace with a real photo/clean typography) if ANY flaw exists:
     - **Anatomy/hands** (extra/fused fingers, missing joints, blurry ear-hair junctions);
     - **Geometry/closure** (floating handles, double rims, mismatched lids, missing jar bases or floating tables);
     - **Physics/lighting** (shadow contradicting light source, light leaking through opaque ceramic/stone walls, static curly smoke);
     - **Period/pattern accuracy** (broken geometric patterns, anachronistic Baroque ornament on ancient artifacts, wrong botanical/historical objects);
     - **Copy-paste artifacts** (cloned flames, tiled textures, melted reliefs).
   - **Human Editorial Framing**: Never drop raw generator aspect ratios (`1376x768`, `896x1200`) unchanged onto slides. Apply deliberate editorial cropping via **native shape-fill framing** (`p:sp` + `prstGeom`/`custGeom` + `blipFill`), add grounded captions/annotations where helpful, and enforce the `custGeom` 3-point check, z-order, and brightness gate (`<140` warning on light-bg/dark-text slides) per `references/frontend-slides-motion.md`.
6. **Density-Driven Font Floor & Incremental Inspection (2-3 Slides per Batch)**:
   - Cap typography at 4 disciplined size tiers and enforce the **density-split body floor**: **Speaking-led low-density `>= 20pt`** (if text overflows, cut words or expand the container — never shrink font size); **Reading-led high-density `14-16pt`** (never drop body below `14pt`; captions `10-11pt`, section `16-20pt`, page title `24-32pt`).
   - Render each 2-3 slide batch via `soffice` + `pdftoppm` and visually inspect every single slide before proceeding.
 
## Style-Driven & Intra-Page Font Pairing Module
 
Never lock an entire deck into a single font family, and never confine choices solely to `微软雅黑` / `思源黑体`. **First confirm the PPT visual style in `grill-me`**, then map that style to a **2-3 font intra-page combination** (3-4 curated families across the deck, always writing all 3 XML font slots cleanly so CJK never falls into Latin-only fonts; never scatter 5+ random fonts on body text, and continue banning `Noto` series):
 
1. **华丽发布会 / 视觉冲击风 (Ornate / Launch)**: Hero titles in `庞门正道标题体` / `站酷高端黑` / `Cinzel`; pull-quotes or accent callouts in `方正大标宋` / `华文行楷` / `LXGW WenKai (霞鹜文楷)`; KPI numerals in `DIN` / `Bahnschrift`; body in `苹方 PingFang SC` / `方正兰亭黑` / `HarmonyOS Sans SC`.
2. **典雅杂志 / 人文叙事风 (Elegant / Editorial)**: Titles in `思源宋体` / `方正书宋` / `Playfair Display`; verbatim quotes or policy excerpts in `LXGW WenKai (霞鹜文楷)` / `楷体` / `仿宋`; page tags/metadata in `Consolas`; body in `冬青黑体` / `思源黑体` / `MiSans`.
3. **瑞士极简 / 硬核数据风 (Swiss / Data-Driven)**: Bold headers in `方正兰亭粗黑` / `思源黑体 Bold` / `MiSans`; oversized KPIs & axis labels in `DIN` / `JetBrains Mono` (numbers/Latin only); captions in light sans; body in `苹方` / `HarmonyOS Sans SC` / `微软雅黑`.
4. **商务稳重 / 政企汇报风 (Executive / Formal)**: Titles in `华文中宋` / `方正小标宋`; numeric highlights in `Bahnschrift` / `Cambria`; policy citations in `仿宋` / `LXGW WenKai (霞鹜文楷)`; body in `微软雅黑` / `苹方` / `HarmonyOS Sans SC`.
 
- **Royalty-Free Open-Source CJK Alternatives**: When commercial-license safety is needed on external machines, substitute with `LXGW WenKai (霞鹜文楷)` (quotes/editorial), `MiSans`, or `HarmonyOS Sans SC` (clean sans) — still avoid `Noto Sans SC` / `Noto Serif SC`.
 
### English / Latin Font Pairing Extension (International Decks)
 
For English-only or English-primary decks, apply the **same discipline** as the CJK pairings above: style is confirmed in `grill-me` first, then each slide combines **2-3 Latin families by role** (display / body / numerals-or-meta), **3-4 families across the whole deck**, one family per role deck-wide. Never set an entire English deck in a single default sans, and never use `Arial`, `Inter`, `Roboto`, or `Helvetica` defaults (banned as overused in `references/de-ai-pptx-v3.1.md`).
 
1. **Ornate / Launch**: Hero titles in `Cinzel` / `Playfair Display` / `Bebas Neue` (condensed poster headlines); pull-quotes & accents in `Playfair Display Italic` / `Libre Baskerville Italic`; KPI numerals in `Oswald` / `Archivo` / `Bahnschrift`; body in `Source Sans 3` / `Lato` / `Work Sans`.
2. **Elegant / Editorial**: Titles in `Playfair Display` / `Libre Baskerville` / `Lora`; verbatim quotes in `EB Garamond Italic` / `Lora Italic`; page tags & metadata in `IBM Plex Mono` / `Space Mono`; body in `Source Serif 4` / `Lora`, captions in `Source Sans 3`.
3. **Swiss / Data-Driven**: Bold headers in `Archivo` (incl. Expanded weights) / `Libre Franklin` / `Space Grotesk`; oversized KPIs & axis labels in `IBM Plex Mono` / `JetBrains Mono` with tabular figures; captions in light-weight `Archivo` / `Libre Franklin`; body in `Archivo` / `Public Sans`.
4. **Executive / Formal**: Titles in `Libre Baskerville` / `Source Serif 4` / `Spectral`; numeric highlights in `Cambria` / `Source Serif 4`; citations & footnotes in `EB Garamond Italic`; body in `Source Sans 3` / `Public Sans` / `Lato`.
 
**Availability tiers (check before committing a deck):**
- **Tier A — preinstalled, safe on unknown machines**: `Georgia` (elegant serif, Win+Mac), `Garamond` / `Palatino Linotype` / `Book Antiqua` (formal serif), `Cambria` / `Constantia` (Office serif), `Century Gothic` (geometric display). Prefer Tier A when the deck will be presented from someone else's computer.
- **Tier B — free Google Fonts** (all families named above not in Tier A): use freely, but a `.pptx` does not reliably carry fonts with it — if the deck leaves the authoring machine, either embed the fonts on export or fall back to the Tier A partner for that role.
- **Mixed CJK + Latin decks**: write the Latin slot with the English family and the East Asian slot with its CJK partner from the pairings above (e.g. Latin `Playfair Display` + EA `思源宋体`; Latin `Archivo` + EA `思源黑体`; Latin `IBM Plex Mono` for numerals in both). Never let Latin text fall back to a CJK font's built-in Latin glyphs, and never let CJK fall into a Latin-only font.
 
## Anti-AI-Slop Hard Bans & Plain-Copy Rules (All Routes, Including Route 3)
 
1. **Plain Fact Titles, No Rhetorical/Negation Clichés & Zero Buzzwords**:
   - **Ban "Clause A + comma + Clause B" symmetrical titles**, **repeated rhetorical-question titles** (`"为什么...?"` / `"究竟是 A，还是 B?"`), and **"Data + Judgment + Metaphor" slogans**. Mix plain factual titles, dates/names, and concise noun phrases.
   - **Ban `"不是 A，而是 B"` / `"它不是 X。它是 Y"` negation-pivot clichés** (especially in narrative/cultural decks) and **standalone em-dash + quoted aphorism lines** (`——“...”`) repeated across slides.
   - **Blacklist AI & policy-jargon buzzwords**: Never use `锚点`, `压舱石`, `赋能`, `变现`, `闭环`, `卡位`, `打法`, `突围`, `红利`, `赛道`, `真金白银`, `双线并行`, `用脚投票`, `落地`, `下沉`, `分流`, `硬证据`, `领跑`, `退潮`, `精准`, `矩阵`, `工具箱`, `主战场`, `主通道`, `非共识`, `研判`, `霸榜`, `凸显`, `托底`, `攻坚`, `趋于理性`, `高质量发展`. Replace with plain wording.
   - **Ban 4-beat parallelism, empty divider pages, scare quotes & self-certifying claims**: Never build a divider slide out of 4-char parallelisms, never stack `靠A，靠B，靠C`, never put scare quotes around common terms (`"国家队"`, `"铁饭碗"`), and **never write `"编辑部核心研判"` or `"所有外部数据均已核验"`**.
2. **Sentence Burstiness, Nuanced Judgment & Asymmetric Structure**:
   - **Increase burstiness**: After every 2 long sentences, insert a short punchy sentence (<= 10 chars). Vary paragraph openings across data-first, narrative facts, and concessions; add nuanced judgment/caveats instead of robotic absolute conclusions.
   - **Break structural symmetry & drop half the closing takeaways**: Expand key slides and keep minor points brief; omit closing summary strips on at least half of the slides; ban `"三大xx / 四大xx"` grouping, all-caps English kickers (`SECTION 01 | ...`), and `｜`/`→`/`·`/`／` clustering. Vary layouts across 4-5 distinct structures (>= 2 real slide layouts; never > 30% identical coordinates).
3. **Clean Document Properties & OOXML Hygiene (Mandatory Post-Processing)**:
   - Keep `docProps/app.xml` and `docProps/core.xml` accurate and complete: sync `Slides` to actual slide count, `Words` to actual word count, `PresentationFormat` to `"On-screen Show (16:9)"`, and populate `author`, `last_modified_by`, and `title` consistently with user-confirmed values (never leave `"待填写"` or invented `"xx课题组"`).
   - Format OOXML declarations cleanly with standard double quotes (`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`).
   - **Human-Plausible Timestamps (Mandatory)**: Never ship a deck whose `docProps/core.xml` timestamps expose instant machine generation (`created` == `modified`, or a whole deck "made" in under a minute). Rewrite them to mimic a real human editing session:
     - `created`: a plausible start within the 3 days before delivery, during waking hours (09:00–23:00 local time), with non-round minutes/seconds (never `00:00:00`, never the exact delivery minute).
     - `modified`: `created` plus a realistic production span scaled to deck size — roughly **6–12 minutes per slide, minimum 40 minutes, maximum ~8 hours** (keep it same-day unless the deck is very large).
     - `revision`: scale with the span (about 1 save per 15–25 minutes of span, at least 3).
     - `TotalTime` in `docProps/app.xml` (minutes): a believable fraction of the created→modified span (about 40–70% — humans pause), never 0 and never equal to the full span.
     - Keep `created < modified` strictly, and keep both consistent with any dates stated inside the deck.
4. **Vary Source Citations & Natural Speaker Notes (NO Meta-Coaching Templates)**:
   - Never repeat an identical source line at the bottom of every slide; consolidate full citations on a dedicated Data & Sources slide or vary inline/footnote mentions.
   - **Ban stage-direction speaker notes**: Never write every slide's note as coaching instructions (`"这一页不需要念完每个神的名字，可以顺着视觉重点点出..."`) or `"第 N 页..."` templates. Write natural, uneven spoken cues (mix short 15-30 char reminders, factual bullets, and occasional 80-120 char talking notes; some simple visual slides can have brief or minimal notes).
5. **Final Proofread, Widow-Line & Collision Gate + Clean Charts/Links**:
   - **Mandatory Text & Visual Final Check**: Before delivery, run a text pass to catch duplicated words (e.g., `原本根本`, `的的`) and inspect **every single rendered PNG** for text-on-text overlap, clipped boxes, and 1-3 character widow lines (e.g., a lone `"沉默。"` on a centered line). Fix and re-render until zero collisions or widows remain.
   - Set real chart titles or `has_title = False`, set `minimum_scale = 0` on non-negative axes, style hyperlinks in high-contrast dark colors, merge same-URL runs into **one single run** matching the actual URL domain, update `theme1.xml`, disable shrink-to-fit auto-scaling, and use native bullets.
 
## Hard Gate: Alignment Before Output (`grill-me` Protocol + Style & Motion Gate)
 
**STOP AND ASK FIRST. Never generate or edit any `.pptx` or `.html` file on Turn 1** (except Route 2 Audit-Only, which outputs a score report without editing files), unless the user explicitly says "skip questions / generate directly".
 
1. Before writing any output file in **Route 1 (Refactor)**, **Route 3 (Author from Scratch)**, or **Route 4 (Visual Reconstruction)**, read `references/de-ai-pptx-v3.1.md` and `references/frontend-slides-motion.md` in full.
2. Execute the **`grill-me` interrogation protocol** from Chapter 7 of `references/de-ai-pptx-v3.1.md` on Turn 1:
   - Never ask facts already present in the prompt or attached files (investigate/extract those yourself first).
   - Group unblocked decisions into a frontier round: in **Route 1**, follow B01-B09; in **Route 3**, ask author/attribution, 7-Question Intake, density tier (Speaking `>= 20pt` vs. Reading `14-16pt`), real screenshots/photos, and 1-2 specific personal/local examples using the exact `grill-me` format.
   - **Mandatory Style Question (Must Ask in `grill-me`)**: Explicitly ask which **PPT visual style** the user prefers (e.g., 华丽发布会风 / 典雅杂志风 / 瑞士极简风 / 商务稳重风), and explain that the chosen style will drive the multi-font pairing within each slide.
   - **Mandatory Motion Question (Must Ask in `grill-me`)**: Explicitly ask whether the user wants **dynamic motion/animation effects** (e.g., native `.pptx` **Morph 平滑切换** with fade fallback, in-slide **auto-chained `afterEffect` animations**, or `frontend-slides` animated HTML per `references/frontend-slides-motion.md`), and ask **which pages (selected key pages like cover/section/hero slides, OR all pages, OR none for now)**.
   - Format every question with: **Number + Title**, **Question** (phrased so "Yes" accepts the recommended answer), **Recommended Answer**, and **Rationale**, separated by dividers.
   - **Stop immediately after asking the round** and wait for the user's reply before building or modifying any slides.
3. **Mandatory Post-Delivery Motion Note**: If the user chose **"不需要动态效果" (no animations for now)** during `grill-me`, **when delivering the final `.pptx` file, explicitly tell the user** that animations are currently omitted and that they can pick specific slide numbers (or all slides) and desired motion style (Morph transitions, in-slide auto-chained animations, or animated HTML) anytime to add dynamic effects.
 
## Task Routing
 
1. **Route 1 - Refactor Existing Deck (改造)**: Read `references/de-ai-pptx-v3.1.md` and `references/frontend-slides-motion.md`. If input is legacy `.ppt`, convert to `.pptx` via `soffice --headless --convert-to pptx` first. Run Phase 0 Compliance Gate -> Phase 1 Forensics -> Phase 2 `grill-me` Interrogation (including Style & Motion Questions) -> Phase 3 Five-Layer Refactor -> Phase 4 Acceptance QA + OOXML/Metadata hygiene -> Deliver modern `.pptx`.
2. **Route 2 - Audit Only (体检)**: Read `references/de-ai-pptx-v3.1.md` in full. Run Phase 0 + Phase 1 forensics and output the dual-score report without modifying the file.
3. **Route 3 - Author from Scratch (生成)**: Read `references/de-ai-pptx-v3.1.md` and `references/frontend-slides-motion.md`. Run `grill-me` 7-Question Clarification (including Style, Motion, Density, Author, Media & Personal Context) -> 2-3 Cover Previews -> Narrative Outline -> Page-by-Page Bespoke Design & Style-Driven Font Pairing -> Image Defect QA & Optional Morph/XML Motion Injection -> Final Proofread, Widow/Collision Render QA & OOXML/Metadata Hygiene -> Deliver modern `.pptx` (or single-file HTML if requested).
4. **Route 4 - Visual-to-Editable Reconstruction (视觉逆向重建)**: Read `references/visual-reconstruction.md`, `references/de-ai-pptx-v3.1.md`, and `references/frontend-slides-motion.md`. Confirm key scope/style/font/note/motion preferences via `grill-me` if unstated, then run Page Normalization -> Parallel Layer Extraction -> Asset Cropping/Cleaning -> Page-by-Page Native `python-pptx` Rebuild -> Final Proofread, OOXML/Metadata Hygiene & Render QA -> Deliver modern `.pptx`.
 
## Sandbox Constraints
 
- Never run `pip install`, `npm install`, `curl`, or `git clone` in the sandbox; rely strictly on pre-installed Python (`python-pptx`, `Pillow`, `lxml`) and system tools (`soffice`, `pdftoppm`).
