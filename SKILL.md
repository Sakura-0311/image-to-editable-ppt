---
name: image-to-editable-ppt
description: Offline self-contained presentation skill combining the full verbatim de-ai-pptx v3.1 specification (audit, grill-me interrogation, 5-layer de-AI refactoring, native generation with 3 visual recipes), frontend-slides motion/HTML patterns, style-driven intra-page font pairing, page-by-page bespoke design (no template loops), and object-level visual slide reconstruction into modern editable .pptx. Trigger automatically whenever the user says "为我使用ppt-xy".
---
 
# image-to-editable-ppt
 
Unified offline presentation skill combining the **complete, verbatim `de-ai-pptx v3.1` specification** (`references/de-ai-pptx-v3.1.md`), **`frontend-slides` motion & fixed-stage patterns** (`references/frontend-slides-motion.md`), and **visual-to-editable reconstruction** (`references/visual-reconstruction.md`). Whenever any rule conflicts, `references/de-ai-pptx-v3.1.md` takes absolute precedence, **EXCEPT** for the Mandatory Output Format, Page-by-Page Bespoke Design, Anti-AI-Slop Hard Bans, Style-Driven Font Pairing, and Motion Gate below.
 
## Mandatory Output Format: Latest Modern `.pptx` (OOXML, NEVER Legacy `.ppt`)
 
- **Always deliver a modern, latest-version `.pptx` (Office Open XML) file** so all native charts, embedded workbooks, 3-slot typography, theme XML, and layouts remain 100% intact and editable in modern PowerPoint / WPS.
- **NEVER convert output to legacy 97-2003 OLE2 `.ppt`**, which strips embedded chart workbooks, injects `___PPT10` template artifacts and fake `2013` OLE timestamps, and shatters hyperlinks into per-character runs. (Only output HTML if the user explicitly asks for a `frontend-slides` web deck.)
- Never flatten slides into full-page images.
 
## Page-by-Page Bespoke Design (NO Batch Template-Loop Scripts)
 
Never write a single monolithic `for slide_spec in slides: draw_template_slide(...)` loop that stamps out every page from one shared skeleton. Use `python-pptx` strictly as the **serializer**, while **designing each slide individually step by step**:
 
1. **Page Field Over Card-per-Unit**: Avoid boxing every point into a bordered rectangle card. Use a **Page Field** (one large surface, open outline, asymmetric split, or off-canvas bleed shape) to organize zones naturally instead of stacking uniform cards.
2. **Geometry Signature Check (Anti-Collision)**: Before designing Slide `N+1`, record and compare its Geometry Signature (`page job -> composition move -> contour/edge language`) against the previous 2 slides; if it matches, switch to a different spatial composition.
3. **Topology-Driven Geometry**: Decide `topology=yes/no` per slide based on whether the content has real relationships (sequence, causality, hierarchy, contrast, or overlap). Draw connectors, arrows, or nested containers only when `topology=yes`; keep `topology=no` slides clean and unadorned.
4. **Expanded Visual Expression Range & Layer-Accurate Fixes**: Go beyond plain rectangles using full-bleed image fields, editorial crops, scrim/spotlight overlays, freeform contours, and a subtle cross-page motif varied by page role. Fix defects at the owning layer (page-local fix for layout issues, outline fix for structure) without silently downgrading required assets.
5. **Build & Inspect Incrementally (2-3 Slides per Batch)**: Cap typography at 4 disciplined size tiers (Page Title 24-28pt, Section 16-18pt, Body 12-14pt, Caption 10-11pt). Render each batch via `soffice` + `pdftoppm` and visually inspect before proceeding. Always populate `author`, `last_modified_by`, and `title` consistently, set `created < modified`, and set `revision >= 2`.
 
## Style-Driven & Intra-Page Font Pairing Module
 
Never lock an entire deck into a single font family, and never confine choices solely to `微软雅黑` / `思源黑体`. **First confirm the PPT visual style in `grill-me`**, then map that style to a **2-3 font intra-page combination** (3-4 curated families across the deck, always writing all 3 XML font slots cleanly so CJK never falls into Latin-only fonts; never scatter 5+ random fonts on body text):
 
1. **华丽发布会 / 视觉冲击风 (Ornate / Launch)**: Hero titles in `庞门正道标题体` / `站酷高端黑` / `Cinzel`; pull-quotes or accent callouts in `方正大标宋` / `华文行楷`; KPI numerals in `DIN` / `Bahnschrift`; body in `苹方 PingFang SC` / `方正兰亭黑`.
2. **典雅杂志 / 人文叙事风 (Elegant / Editorial)**: Titles in `思源宋体` / `方正书宋` / `Playfair Display`; verbatim quotes or policy excerpts in `楷体` / `仿宋`; page tags/metadata in `Consolas`; body in `冬青黑体` / `思源黑体`.
3. **瑞士极简 / 硬核数据风 (Swiss / Data-Driven)**: Bold headers in `方正兰亭粗黑` / `思源黑体 Bold`; oversized KPIs & axis labels in `DIN` / `JetBrains Mono` (numbers/Latin only); captions in light sans; body in `苹方` / `微软雅黑`.
4. **商务稳重 / 政企汇报风 (Executive / Formal)**: Titles in `华文中宋` / `方正小标宋`; numeric highlights in `Bahnschrift` / `Cambria`; policy citations in `仿宋`; body in `微软雅黑` / `苹方`.
 
## Anti-AI-Slop Hard Bans & Plain-Copy Rules (All Routes, Including Route 3)
 
1. **Plain Fact Titles, No Comma-Paired Antithesis & Zero Buzzwords**:
   - **Ban "Clause A + comma + Clause B" symmetrical titles/subtitles** (e.g., 13+13 or 17+17 char balanced halves) and **"Data + Judgment + Metaphor" slogans**. Use single-sentence factual statements, questions, or natural spoken lines (`"AI岗位一年多了三成，但招人的不只是互联网"`, `"国考、国央企与基层岗位招录情况"`).
   - **Blacklist AI & policy-jargon buzzwords**: Never use `锚点`, `压舱石`, `赋能`, `变现`, `闭环`, `卡位`, `打法`, `突围`, `红利`, `赛道`, `真金白银`, `双线并行`, `用脚投票`, `落地`, `下沉`, `分流`, `硬证据`, `领跑`, `退潮`, `精准`, `矩阵`, `工具箱`, `主战场`, `主通道`, `非共识`, `研判`, `霸榜`, `凸显`, `托底`, `攻坚`, `趋于理性`, `高质量发展`. Replace with plain wording (`重点`, `保底/退路`, `集中投递/补录`, `不再盲目跟风`, `能转正的实习`, `主要渠道`).
   - **Concrete Rewrite Standard**:
     - Bad (AI): `"应届生身份保护：2026国考13,822个职位限应届生，且年龄放宽至38周岁。选调生与省考：各省定向选调多在10-12月启动..."`
     - Good (Human): `"应届身份真的很值钱——2026年国考1.38万个职位只招应届生，年龄还放宽到了38岁。但注意，这个身份基本上就用这一年。选调是另一条路..."`
   - **Ban 4-beat parallelism, empty divider pages, scare quotes & self-certifying claims**: Never build a divider slide out of 4-char parallelisms, never stack `靠A，靠B，靠C`, never put scare quotes around common terms (`"国家队"`, `"铁饭碗"`), and **never write `"编辑部核心研判"` or `"所有外部数据均已核验"`**.
2. **Sentence Burstiness, Nuanced Judgment & Asymmetric Structure**:
   - **Increase burstiness**: After every 2 long sentences, insert a short punchy sentence (<= 10 chars, e.g. `"竞争比98:1。什么概念？一个岗位近百人抢。"`). Vary paragraph openings across data-first, rhetorical questions, and concessions; add nuanced judgment/caveats (`"我倾向于理解为挤水分，而不是学历没用了"`) instead of robotic absolute conclusions.
   - **Break structural symmetry & drop half the closing takeaways**: Expand key slides and keep minor points brief; omit closing summary strips on at least half of the slides; ban `"三大xx / 四大xx"` grouping, all-caps English kickers (`SECTION 01 | ...`), and `｜`/`→`/`·`/`／` clustering. Vary layouts across 4-5 distinct structures (>= 2 real slide layouts; never > 30% identical coordinates).
3. **Erase `python-pptx` XML & `app.xml` Fingerprints (Mandatory Post-Processing)**:
   - Rewrite all `*.xml` and `*.rels` declarations inside the `.pptx` zip from single quotes (`'1.0'`) to **double quotes** (`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`).
   - Sync `docProps/app.xml`: set `Slides` to actual count, `Words` to actual word count, `PresentationFormat` to `"On-screen Show (16:9)"`, and realistic `TotalTime`.
4. **Vary Source Citations & Speaker Notes (Never Uniform Templates)**:
   - Never repeat an identical source line at the bottom of every slide; consolidate full citations on a dedicated Data & Sources slide or vary inline/footnote mentions.
   - Write natural, uneven speaker notes (mix 15-30 char reminders with 80-120 char spoken cues; never start every note with `"第 N 页..."`).
5. **Ban `main-title`, `___PPT10`, Negative Axes, Broken Hyperlinks & 0-Media Walls**:
   - Set real chart titles or `has_title = False`, set `minimum_scale = 0` on non-negative axes, style hyperlinks in high-contrast dark colors, merge same-URL runs into **one single run** with clean source names matching the actual URL domain.
   - Write all 3 font slots explicitly (`theme1.xml` updated), disable shrink-to-fit auto-scaling, use native bullets, and ask in `grill-me` for real screenshots/personal case details (1-2 specific examples or local context) so the deck has real human grounding.
 
## Hard Gate: Alignment Before Output (`grill-me` Protocol + Style & Motion Gate)
 
**STOP AND ASK FIRST. Never generate or edit any `.pptx` or `.html` file on Turn 1** (except Route 2 Audit-Only, which outputs a score report without editing files), unless the user explicitly says "skip questions / generate directly".
 
1. Before writing any output file in **Route 1 (Refactor)**, **Route 3 (Author from Scratch)**, or **Route 4 (Visual Reconstruction)**, read `references/de-ai-pptx-v3.1.md` and `references/frontend-slides-motion.md` in full.
2. Execute the **`grill-me` interrogation protocol** from Chapter 7 of `references/de-ai-pptx-v3.1.md` on Turn 1:
   - Never ask facts already present in the prompt or attached files (investigate/extract those yourself first).
   - Group unblocked decisions into a frontier round: in **Route 1**, follow B01-B09; in **Route 3**, ask author/attribution, 7-Question Intake, density tier, real screenshots, and 1-2 specific personal/local examples using the exact `grill-me` format.
   - **Mandatory Style Question (Must Ask in `grill-me`)**: Explicitly ask which **PPT visual style** the user prefers (e.g., 华丽发布会风 / 典雅杂志风 / 瑞士极简风 / 商务稳重风), and explain that the chosen style will drive the multi-font pairing within each slide.
   - **Mandatory Motion Question (Must Ask in `grill-me`)**: Explicitly ask whether the user wants **dynamic motion/animation effects** (e.g., subtle 200-300ms corporate fades, staggered editorial reveals, or dramatic cover/section transitions per `references/frontend-slides-motion.md`), and ask **which pages (selected key pages like cover/section/hero slides, OR all pages, OR none for now)**.
   - Format every question with: **Number + Title**, **Question** (phrased so "Yes" accepts the recommended answer), **Recommended Answer**, and **Rationale**, separated by dividers.
   - **Stop immediately after asking the round** and wait for the user's reply before building or modifying any slides.
3. **Mandatory Post-Delivery Motion Note**: If the user chose **"不需要动态效果" (no animations for now)** during `grill-me`, **when delivering the final `.pptx` file, explicitly tell the user** that animations are currently omitted and that they can pick specific slide numbers (or all slides) and desired motion style anytime to add dynamic effects (either inside `.pptx` or as a `frontend-slides` animated HTML deck).
 
## Task Routing
 
1. **Route 1 - Refactor Existing Deck (改造)**: Read `references/de-ai-pptx-v3.1.md` and `references/frontend-slides-motion.md`. If input is legacy `.ppt`, convert to `.pptx` via `soffice --headless --convert-to pptx` first. Run Phase 0 Compliance Gate -> Phase 1 Forensics -> Phase 2 `grill-me` Interrogation (including Style & Motion Questions) -> Phase 3 Five-Layer Refactor -> Phase 4 Acceptance QA + XML/app.xml hygiene -> Deliver modern `.pptx`.
2. **Route 2 - Audit Only (体检)**: Read `references/de-ai-pptx-v3.1.md` in full. Run Phase 0 + Phase 1 forensics and output the dual-score report without modifying the file.
3. **Route 3 - Author from Scratch (生成)**: Read `references/de-ai-pptx-v3.1.md` and `references/frontend-slides-motion.md`. Run `grill-me` 7-Question Clarification (including Style, Motion, Author, Media & Personal Context) -> 2-3 Cover Previews -> Narrative Outline -> Page-by-Page Bespoke Design & Style-Driven Font Pairing -> Incremental Render QA & L0/L1/XML Hygiene -> Deliver modern `.pptx` (or single-file HTML if requested).
4. **Route 4 - Visual-to-Editable Reconstruction (视觉逆向重建)**: Read `references/visual-reconstruction.md`, `references/de-ai-pptx-v3.1.md`, and `references/frontend-slides-motion.md`. Confirm key scope/style/font/note/motion preferences via `grill-me` if unstated, then run Page Normalization -> Parallel Layer Extraction -> Asset Cropping/Cleaning -> Page-by-Page Native `python-pptx` Rebuild -> XML Hygiene & Render QA -> Deliver modern `.pptx`.
 
## Sandbox Constraints
 
- Never run `pip install`, `npm install`, `curl`, or `git clone` in the sandbox; rely strictly on pre-installed Python (`python-pptx`, `Pillow`) and system tools (`soffice`, `pdftoppm`).
- Always convert finished decks to PDF (`soffice --headless --convert-to pdf`) and PNGs (`pdftoppm -png -r 150`) to visually inspect every slide before delivery.
