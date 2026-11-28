#!/usr/bin/env python3
"""
tools/audit.py — Read-only delivery audit for .pptx decks (BLOCK / WARN grading).

Standard library + lxml only (zero network, zero pip install).
Usage:
  python3 tools/audit.py deck.pptx [--brief] [--json]
Exit codes:
  0 = pass (0 BLOCK findings)
  1 = fail (>= 1 BLOCK findings, or unreadable .pptx)
  2 = CLI argument error
"""

import argparse
from collections import Counter
from datetime import datetime
import json
import os
import re
import sys
import zipfile
from lxml import etree

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "p159": "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
}

A_NS = NS["a"]
P_NS = NS["p"]
P159_NS = NS["p159"]
CP_NS = NS["cp"]
DC_NS = NS["dc"]
DCTERMS_NS = NS["dcterms"]
EP_NS = NS["ep"]

CJK_CHAR_RE = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbf]")
LATIN_WORD_RE = re.compile(r"[A-Za-z0-9]+(?:['’-][A-Za-z0-9]+)*")
SINGLE_QUOTE_DECL_RE = re.compile(rb"^\s*<\?xml[^>]*'1\.0'", re.DOTALL)

KNOWN_CJK_FONT_TOKENS = (
    "yahei", "微软雅黑", "思源", "source han", "simsun", "simhei", "宋体", "黑体",
    "楷体", "仿宋", "kaiti", "fangsong", "pingfang", "苹方", "冬青", "hiragino",
    "方正", "华文", "汉仪", "站酷", "庞门", "霞鹜", "lxgw", "misans", "harmonyos",
    "stsong", "stheiti", "stkaiti", "stfangsong", "dengxian", "等线", "noto sans cjk",
    "noto serif cjk", "noto sans sc", "noto serif sc",
)

SERIF_OR_DISPLAY_TOKENS = (
    "serif", "宋体", "书宋", "标宋", "中宋", "楷", "仿宋", "行楷", "wenkai", "霞鹜",
    "georgia", "garamond", "cambria", "playfair", "baskerville", "lora", "cinzel",
    "bodoni", "palatino", "spectral", "constantia", "mono", "consolas", "din",
    "bahnschrift", "bebas", "oswald", "庞门", "站酷",
)

BAD_META_TOKENS = (
    "待填写", "待补", "课题组", "python-pptx", "unknown", "template", "placeholder",
)

# Fonts preinstalled on typical end-user machines (Windows + Office, macOS).
# Anything used by a deck but NOT in this set will be silently substituted by
# PowerPoint/WPS on machines that lack it — which can wreck the design
# (real-machine test 2026-11-29: a Playfair Display / Source Sans 3 /
# IBM Plex Mono deck rendered entirely in a handwriting fallback font).
PREINSTALLED_SAFE_FONTS = {
    # Latin — Windows / Office staples
    "arial", "arial black", "arial narrow", "book antiqua", "calibri",
    "cambria", "candara", "century gothic", "comic sans ms", "consolas",
    "constantia", "corbel", "courier new", "franklin gothic medium",
    "garamond", "georgia", "gill sans mt", "impact", "lucida console",
    "lucida sans unicode", "microsoft sans serif", "palatino linotype",
    "perpetua", "segoe print", "segoe script", "segoe ui", "tahoma",
    "times new roman", "trebuchet ms", "tw cen mt", "verdana",
    # Latin — macOS staples
    "avenir", "avenir next", "helvetica", "helvetica neue", "menlo",
    "monaco", "optima", "futura", "didot", "american typewriter",
    # CJK — Windows
    "microsoft yahei", "microsoft yahei ui", "微软雅黑", "simsun", "宋体",
    "nsimsun", "新宋体", "simhei", "黑体", "kaiti", "楷体", "fangsong",
    "仿宋", "dengxian", "等线", "microsoft jhenghei", "微軟正黑體",
    "pmingliu", "新細明體", "mingliu", "細明體",
    # CJK — macOS
    "pingfang sc", "pingfang tc", "pingfang hk", "苹方", "hiragino sans",
    "hiragino sans gb", "冬青黑体", "stsong", "华文宋体", "stheiti",
    "华文黑体", "stkaiti", "华文楷体", "stfangsong", "华文仿宋",
    "songti sc", "宋体-简",
    # CJK — Japanese / Korean preinstalled
    "yu gothic", "yu mincho", "meiryo", "ms gothic", "ms mincho",
    "malgun gothic",
}

_FONT_SUFFIX_TOKENS = {
    "bold", "italic", "oblique", "light", "semilight", "medium",
    "semibold", "demibold", "black", "thin", "regular", "book",
    "condensed", "extended", "roman",
}


def normalize_font_name(name: str) -> str:
    parts = name.strip().lower().split()
    while len(parts) > 1 and parts[-1] in _FONT_SUFFIX_TOKENS:
        parts.pop()
    return " ".join(parts)


def is_preinstalled_safe_font(font_name: str) -> bool:
    return normalize_font_name(font_name) in PREINSTALLED_SAFE_FONTS


def is_cjk_font(font_name: str) -> bool:
    if not font_name:
        return False
    low = font_name.strip().lower()
    if CJK_CHAR_RE.search(font_name):
        return True
    return any(tok in low for tok in KNOWN_CJK_FONT_TOKENS)


def is_noto_font(font_name: str) -> bool:
    return bool(font_name) and "noto" in font_name.strip().lower()


def is_serif_or_display(font_name: str) -> bool:
    low = font_name.strip().lower()
    return any(tok in low for tok in SERIF_OR_DISPLAY_TOKENS)


def count_words(text: str) -> int:
    if not text:
        return 0
    return len(CJK_CHAR_RE.findall(text)) + len(LATIN_WORD_RE.findall(text))


def parse_iso_dt(raw: str | None) -> datetime | None:
    if not raw:
        return None
    s = raw.strip().rstrip("Z")
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def detect_expected_format(pres_bytes: bytes) -> str:
    root = etree.fromstring(pres_bytes)
    sld_sz = root.find(f"{{{P_NS}}}sldSz")
    if sld_sz is None:
        return "On-screen Show (16:9)"
    cx = int(sld_sz.get("cx", "0") or 0)
    cy = int(sld_sz.get("cy", "0") or 0)
    if cx <= 0 or cy <= 0:
        return "On-screen Show (16:9)"
    ratio = cx / cy
    if abs(ratio - 16.0 / 9.0) < 0.08:
        return "On-screen Show (16:9)"
    if abs(ratio - 4.0 / 3.0) < 0.08:
        return "On-screen Show (4:3)"
    if abs(ratio - 16.0 / 10.0) < 0.08:
        return "On-screen Show (16:10)"
    return "Custom"


def audit_pptx(path: str) -> tuple[int, dict]:
    findings: list[dict] = []

    def add(severity: str, category: str, code: str, message: str) -> None:
        findings.append({
            "severity": severity,
            "category": category,
            "code": code,
            "message": message,
        })

    if not os.path.isfile(path):
        add("BLOCK", "file", "FILE_NOT_FOUND", f"File not found: {path}")
        return 1, {"file": path, "status": "BLOCK", "block_count": 1, "warn_count": 0, "findings": findings}

    try:
        with zipfile.ZipFile(path, "r") as zin:
            entries = {info.filename: zin.read(info.filename) for info in zin.infolist()}
    except zipfile.BadZipFile:
        add("BLOCK", "file", "INVALID_ZIP", f"Not a valid .pptx zip archive: {path}")
        return 1, {"file": path, "status": "BLOCK", "block_count": 1, "warn_count": 0, "findings": findings}

    if "ppt/presentation.xml" not in entries:
        add("BLOCK", "file", "MISSING_PRESENTATION_XML", "Missing ppt/presentation.xml")
        return 1, {"file": path, "status": "BLOCK", "block_count": 1, "warn_count": 0, "findings": findings}

    slide_paths = sorted(
        [k for k in entries if re.match(r"^ppt/slides/slide\d+\.xml$", k)],
        key=lambda p: int(re.search(r"slide(\d+)\.xml$", p).group(1)),
    )
    actual_slides = len(slide_paths)
    expected_fmt = detect_expected_format(entries["ppt/presentation.xml"])
    pres_cx = pres_cy = 0
    try:
        _proot = etree.fromstring(entries["ppt/presentation.xml"])
        _sz = _proot.find(f"{{{P_NS}}}sldSz")
        if _sz is not None:
            pres_cx, pres_cy = int(_sz.get("cx", "0")), int(_sz.get("cy", "0"))
    except Exception:
        pass

    # Count actual words & inspect slides
    actual_words = 0
    missing_3slot_runs = 0
    cjk_ea_latin_runs = 0
    noto_refs: set[str] = set()
    used_fonts: set[str] = set()
    script_shape_names: list[str] = []
    fullpage_image_slides: list[str] = []
    bad_morph_slides: list[str] = []
    ppt10_hits = 0
    transition_types: list[str] = []
    timing_slides = 0

    for spath in slide_paths:
        raw = entries[spath]
        if b"___PPT10" in raw:
            ppt10_hits += 1
        root = etree.fromstring(raw)

        # Motion scan: transition type + in-slide animation presence
        trans_el = root.find(f"{{{P_NS}}}transition")
        if trans_el is not None and len(trans_el) > 0:
            transition_types.append(etree.QName(trans_el[0]).localname)
        if root.find(f"{{{P_NS}}}timing") is not None:
            timing_slides += 1

        # Check invalid Morph option="object"
        for m_el in root.iter(f"{{{P159_NS}}}morph"):
            opt = m_el.get("option", "")
            if opt and opt != "byObject":
                bad_morph_slides.append(f"{spath}(option={opt})")

        # Check script-style shape names (e.g. !!hero_jar, card_01)
        for cnv in root.iter(f"{{{P_NS}}}cNvPr"):
            sname = cnv.get("name", "")
            if sname.startswith("!!") or re.fullmatch(r"[a-z][a-z0-9]*(_[a-z0-9]+)+", sname or ""):
                script_shape_names.append(f"{os.path.basename(spath)}:{sname}")

        # Editability heuristic: slide is a single full-page picture with no text shapes
        pics = root.findall(f".//{{{P_NS}}}pic")
        has_text = any((t.text or '').strip() for t in root.iter(f"{{{A_NS}}}t"))
        if len(pics) == 1 and not has_text:
            ext = pics[0].find(f".//{{{A_NS}}}ext")
            if ext is not None and pres_cx and pres_cy:
                try:
                    cov = (int(ext.get('cx', '0')) * int(ext.get('cy', '0'))) / (pres_cx * pres_cy)
                except ValueError:
                    cov = 0
                if cov >= 0.9:
                    fullpage_image_slides.append(os.path.basename(spath))

        # Inspect runs
        for r_el in root.iter(f"{{{A_NS}}}r"):
            t_el = r_el.find(f"{{{A_NS}}}t")
            txt = (t_el.text or "") if t_el is not None else ""
            if not txt.strip():
                continue
            actual_words += count_words(txt)
            has_cjk = bool(CJK_CHAR_RE.search(txt))

            rpr = r_el.find(f"{{{A_NS}}}rPr")
            if rpr is None:
                missing_3slot_runs += 1
                continue

            lat_el = rpr.find(f"{{{A_NS}}}latin")
            ea_el = rpr.find(f"{{{A_NS}}}ea")
            cs_el = rpr.find(f"{{{A_NS}}}cs")

            lat_tf = (lat_el.get("typeface") or "").strip() if lat_el is not None else ""
            ea_tf = (ea_el.get("typeface") or "").strip() if ea_el is not None else ""
            cs_tf = (cs_el.get("typeface") or "").strip() if cs_el is not None else ""

            if not lat_tf or not ea_tf or not cs_tf:
                missing_3slot_runs += 1

            for tf in (lat_tf, ea_tf, cs_tf):
                if tf and not tf.startswith("+"):
                    used_fonts.add(tf)
                    if is_noto_font(tf):
                        noto_refs.add(tf)

            if has_cjk and ea_tf and not ea_tf.startswith("+") and not is_cjk_font(ea_tf):
                cjk_ea_latin_runs += 1

    # Check theme XMLs for Noto
    for k, data in entries.items():
        if re.match(r"^ppt/theme/theme\d+\.xml$", k):
            troot = etree.fromstring(data)
            for el in troot.iter():
                tf = el.get("typeface", "")
                if is_noto_font(tf):
                    noto_refs.add(tf)

    # Check chart XMLs for main-title placeholder
    main_title_charts = []
    for k, data in entries.items():
        if re.match(r"^ppt/charts/chart\d+\.xml$", k):
            if b">main-title<" in data:
                main_title_charts.append(k)

    # Check XML declarations for single quotes
    single_quote_xmls = sum(
        1 for k, data in entries.items()
        if (k.endswith(".xml") or k.endswith(".rels")) and SINGLE_QUOTE_DECL_RE.search(data)
    )

    # ── 1. Timestamps & Core Metadata ──
    core_bytes = entries.get("docProps/core.xml")
    app_bytes = entries.get("docProps/app.xml")

    author = ""
    last_mod = ""
    title = ""
    revision_val = None
    created_dt = None
    modified_dt = None

    if core_bytes:
        croot = etree.fromstring(core_bytes)
        el = croot.find(f"{{{DC_NS}}}creator")
        author = (el.text or "").strip() if el is not None else ""
        el = croot.find(f"{{{CP_NS}}}lastModifiedBy")
        last_mod = (el.text or "").strip() if el is not None else ""
        el = croot.find(f"{{{DC_NS}}}title")
        title = (el.text or "").strip() if el is not None else ""
        el = croot.find(f"{{{CP_NS}}}revision")
        if el is not None and (el.text or "").strip().isdigit():
            revision_val = int(el.text.strip())
        el = croot.find(f"{{{DCTERMS_NS}}}created")
        created_dt = parse_iso_dt(el.text if el is not None else None)
        el = croot.find(f"{{{DCTERMS_NS}}}modified")
        modified_dt = parse_iso_dt(el.text if el is not None else None)
    else:
        add("BLOCK", "metadata", "MISSING_CORE_XML", "Missing docProps/core.xml")

    app_slides = None
    app_words = None
    app_fmt = ""
    total_time = None
    if app_bytes:
        aroot = etree.fromstring(app_bytes)
        el = aroot.find(f"{{{EP_NS}}}Slides")
        if el is not None and (el.text or "").strip().isdigit():
            app_slides = int(el.text.strip())
        el = aroot.find(f"{{{EP_NS}}}Words")
        if el is not None and (el.text or "").strip().isdigit():
            app_words = int(el.text.strip())
        el = aroot.find(f"{{{EP_NS}}}PresentationFormat")
        app_fmt = (el.text or "").strip() if el is not None else ""
        el = aroot.find(f"{{{EP_NS}}}TotalTime")
        if el is not None and (el.text or "").strip().isdigit():
            total_time = int(el.text.strip())
    else:
        add("BLOCK", "metadata", "MISSING_APP_XML", "Missing docProps/app.xml")

    # 1. Timestamp rules
    span_minutes = None
    if created_dt is None or modified_dt is None:
        add("BLOCK", "timestamp", "MISSING_TIMESTAMPS", "Missing created or modified timestamp in core.xml")
    elif not (created_dt < modified_dt):
        add("BLOCK", "timestamp", "CREATED_NOT_BEFORE_MODIFIED", f"created ({created_dt}) must be strictly earlier than modified ({modified_dt})")
    else:
        span_minutes = round((modified_dt - created_dt).total_seconds() / 60.0, 2)
        if not (40 <= span_minutes <= 480):
            add("WARN", "timestamp", "SPAN_OUT_OF_RANGE", f"Editing span {span_minutes}m outside 40-480m human range")
        if not (9 <= created_dt.hour < 23):
            add("WARN", "timestamp", "CREATED_OFF_HOURS", f"created hour {created_dt.strftime('%H:%M')} outside 09:00-23:00")
        if created_dt.second == 0 or modified_dt.second == 0 or (created_dt.minute == 0 and created_dt.second == 0):
            add("WARN", "timestamp", "ROUNDED_TIMESTAMP", f"Timestamp has round :00 seconds/minutes (created={created_dt}, modified={modified_dt})")
        if total_time is None or total_time <= 0:
            add("WARN", "timestamp", "INVALID_TOTAL_TIME", f"app.xml TotalTime is missing or 0 ({total_time})")
        elif span_minutes > 0:
            ratio = total_time / span_minutes
            if not (0.30 <= ratio <= 0.80) or total_time == int(round(span_minutes)):
                add("WARN", "timestamp", "TOTAL_TIME_RATIO", f"TotalTime ({total_time}m) is {ratio*100:.1f}% of span ({span_minutes}m), expected 30%-80%")

    if revision_val is None or revision_val < 3:
        add("WARN", "timestamp", "LOW_REVISION", f"cp:revision is {revision_val}, expected >= 3")

    # 2. Metadata rules
    for field_name, val in (("author", author), ("last_modified_by", last_mod), ("title", title)):
        low = val.lower()
        if not val:
            add("BLOCK", "metadata", f"EMPTY_{field_name.upper()}", f"core.xml {field_name} is empty")
        elif any(bad in low for bad in BAD_META_TOKENS):
            add("BLOCK", "metadata", f"PLACEHOLDER_{field_name.upper()}", f"core.xml {field_name} contains placeholder/tool string: {val!r}")

    if app_slides != actual_slides:
        add("BLOCK", "metadata", "SLIDES_MISMATCH", f"app.xml Slides={app_slides} != actual {actual_slides}")
    if app_words != actual_words:
        add("BLOCK", "metadata", "WORDS_MISMATCH", f"app.xml Words={app_words} != actual {actual_words}")
    if app_fmt != expected_fmt:
        add("WARN", "metadata", "FORMAT_MISMATCH", f"app.xml PresentationFormat={app_fmt!r} != actual {expected_fmt!r}")

    # 3. Font rules
    if missing_3slot_runs > 0:
        add("WARN", "font", "MISSING_3SLOT", f"{missing_3slot_runs} run(s) missing one or more of a:latin/a:ea/a:cs")
    if cjk_ea_latin_runs > 0:
        add("BLOCK", "font", "CJK_EA_LATIN_BINDING", f"{cjk_ea_latin_runs} CJK run(s) have Latin-only font in a:ea slot")
    if noto_refs:
        add("BLOCK", "font", "NOTO_FONT_FOUND", f"Forbidden Noto font reference(s): {', '.join(sorted(noto_refs))}")
    if len(used_fonts) == 1 and not any(is_serif_or_display(f) for f in used_fonts):
        add("WARN", "font", "SINGLE_SANS_ONLY", f"Entire deck uses a single sans-serif font: {next(iter(used_fonts))}")
    nonportable_fonts = sorted(f for f in used_fonts if not is_preinstalled_safe_font(f))
    if nonportable_fonts:
        add("WARN", "font", "FONT_PORTABILITY",
            f"Font(s) not preinstalled on typical Windows/macOS machines: {', '.join(nonportable_fonts)}. "
            "On a machine without them, PowerPoint/WPS silently substitutes another font and the layout can break "
            "(verified on real hardware 2026-11-29: a deck in Playfair Display / Source Sans 3 / IBM Plex Mono "
            "rendered entirely in a handwriting fallback). Install these fonts on the target machine, or switch "
            "the deck to preinstalled-safe fonts (Tier A) before delivery.")

    # 4. Structural / OOXML / Morph hygiene
    if single_quote_xmls > 0:
        add("WARN", "xml", "SINGLE_QUOTE_XML_DECL", f"{single_quote_xmls} XML/rels file(s) use single-quoted <?xml version='1.0'?>")
    if main_title_charts:
        add("BLOCK", "chart", "MAIN_TITLE_PLACEHOLDER", f"Chart placeholder 'main-title' found in: {', '.join(main_title_charts)}")
    if ppt10_hits > 0:
        add("BLOCK", "xml", "LEGACY_PPT10_STUB", f"Legacy ___PPT10 placeholder found in {ppt10_hits} slide(s)")
    if bad_morph_slides:
        add("BLOCK", "motion", "INVALID_MORPH_OPTION", f"Invalid p159:morph option (must be byObject): {', '.join(bad_morph_slides)}")
    if script_shape_names:
        add("WARN", "naming", "SCRIPT_SHAPE_NAME", f"Developer-style shape names found: {', '.join(script_shape_names[:5])}")
    if fullpage_image_slides:
        add("WARN", "editability", "FULLPAGE_IMAGE_SLIDE", f"Slide(s) consisting of a single full-page picture with no editable text: {', '.join(fullpage_image_slides)}")
    if len(transition_types) >= 4:
        top_type, top_count = Counter(transition_types).most_common(1)[0]
        if top_count / len(transition_types) >= 0.75:
            add("WARN", "motion", "MOTION_MONOTONY",
                f"{top_count}/{len(transition_types)} transitioned slides use a single transition type ({top_type}). "
                "Known failure mode: defaulting to plain Morph/fade everywhere instead of the effect library. "
                "Confirm a per-page Motion Plan with named library effects was followed, or diversify.")

    block_count = sum(1 for f in findings if f["severity"] == "BLOCK")
    warn_count = sum(1 for f in findings if f["severity"] == "WARN")
    status = "BLOCK" if block_count > 0 else ("WARN" if warn_count > 0 else "PASS")

    report = {
        "file": path,
        "status": status,
        "block_count": block_count,
        "warn_count": warn_count,
        "summary": {
            "slides": actual_slides,
            "words": actual_words,
            "expected_format": expected_fmt,
            "app_format": app_fmt,
            "author": author,
            "title": title,
            "created": created_dt.strftime("%Y-%m-%dT%H:%M:%SZ") if created_dt else None,
            "modified": modified_dt.strftime("%Y-%m-%dT%H:%M:%SZ") if modified_dt else None,
            "span_minutes": span_minutes,
            "revision": revision_val,
            "total_time": total_time,
            "unique_fonts": sorted(used_fonts),
            "motion": {
                "transitioned_slides": len(transition_types),
                "transition_types": dict(Counter(transition_types)),
                "slides_with_inslide_animation": timing_slides,
            },
        },
        "findings": findings,
    }
    return (1 if block_count > 0 else 0), report


def main() -> int:
    p = argparse.ArgumentParser(description="Read-only delivery audit for .pptx decks.")
    p.add_argument("input", help="Target .pptx file path")
    mode = p.add_mutually_exclusive_group()
    mode.add_argument("--brief", action="store_true", help="Output a single verdict line")
    mode.add_argument("--json", action="store_true", help="Output machine-readable JSON")
    args = p.parse_args()

    code, report = audit_pptx(args.input)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    elif args.brief:
        print(
            f"{report['status']} | BLOCK={report['block_count']} WARN={report['warn_count']} | "
            f"slides={report.get('summary', {}).get('slides', 0)} words={report.get('summary', {}).get('words', 0)} | {args.input}"
        )
    else:
        print(f"Audit Result: {report['status']} (BLOCK={report['block_count']}, WARN={report['warn_count']}) — {args.input}")
        if not report["findings"]:
            print("  All checks passed cleanly.")
        else:
            for f in report["findings"]:
                print(f"  [{f['severity']}] ({f['category']}/{f['code']}) {f['message']}")
    return code


if __name__ == "__main__":
    sys.exit(main())
