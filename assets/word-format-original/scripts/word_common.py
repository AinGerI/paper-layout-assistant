from __future__ import annotations

import hashlib
import json
import os
import platform
import re
from collections import Counter
from pathlib import Path
from typing import Any, Dict, Iterable, Optional

import yaml


def require_windows() -> None:
    if platform.system().lower() != "windows":
        raise RuntimeError(
            "word-format requires Windows because it controls the installed Microsoft Word through COM."
        )


def load_profile(path: str | os.PathLike[str]) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    if not isinstance(data, dict):
        raise ValueError("Profile root must be a mapping/object.")
    version = data.get("schema_version", 1)
    if version != 1:
        raise ValueError(f"Unsupported profile schema_version={version}; expected 1.")
    return data


def dump_yaml(data: Dict[str, Any], path: str | os.PathLike[str]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False, width=100)


def connect_word(visible: bool = False):
    require_windows()
    import win32com.client  # type: ignore

    app = win32com.client.gencache.EnsureDispatch("Word.Application")
    app.Visible = bool(visible)
    app.DisplayAlerts = 0
    return app, win32com.client.constants


def clean_text(text: str) -> str:
    return (
        text.replace("\r", "")
        .replace("\x07", "")
        .replace("\v", " ")
        .strip()
    )


def range_text_without_mark(rng) -> str:
    return clean_text(rng.Text)


def sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="replace")).hexdigest()


def safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except Exception:
        return default


def safe_float(value: Any) -> Optional[float]:
    try:
        val = float(value)
        # Word commonly uses large negative sentinels such as wdUndefined.
        if val < -100000:
            return None
        return val
    except Exception:
        return None


def style_name(para) -> str:
    try:
        st = para.Range.Style
        name = getattr(st, "NameLocal", None) or getattr(st, "Name", None)
        return str(name or "")
    except Exception:
        try:
            return str(para.Range.Style)
        except Exception:
            return ""


def paragraph_outline_level(para) -> Optional[int]:
    try:
        return int(para.OutlineLevel)
    except Exception:
        try:
            return int(para.Format.OutlineLevel)
        except Exception:
            return None


def is_in_table(rng, constants) -> bool:
    try:
        return bool(rng.Information(constants.wdWithInTable))
    except Exception:
        return False


def paragraph_math_count(para) -> int:
    try:
        return int(para.Range.OMaths.Count)
    except Exception:
        return 0


def paragraph_inline_shape_count(para) -> int:
    try:
        return int(para.Range.InlineShapes.Count)
    except Exception:
        return 0


def document_snapshot(doc) -> Dict[str, Any]:
    text = str(doc.Content.Text)
    return {
        "text_sha256": sha256_text(text),
        "text_length": len(text),
        "paragraphs": safe_int(doc.Paragraphs.Count),
        "tables": safe_int(doc.Tables.Count),
        "inline_shapes": safe_int(doc.InlineShapes.Count),
        "floating_shapes": safe_int(doc.Shapes.Count),
        "equations": safe_int(doc.OMaths.Count),
        "fields": safe_int(doc.Fields.Count),
        "sections": safe_int(doc.Sections.Count),
    }


def normalize_style_name(name: str) -> str:
    return re.sub(r"\s+", "", name.strip().lower())


HEADING_STYLE_PATTERNS = (
    re.compile(r"^heading\s*([1-9])$", re.I),
    re.compile(r"^标题\s*([1-9])$"),
    re.compile(r"^標題\s*([1-9])$"),
)


def heading_level_from_style(name: str) -> Optional[int]:
    s = name.strip()
    for pattern in HEADING_STYLE_PATTERNS:
        m = pattern.match(s)
        if m:
            return int(m.group(1))
    return None


def heading_level_from_text(text: str) -> Optional[int]:
    # Conservative structural-numbering heuristic. It intentionally avoids treating bare years/values as headings.
    t = text.strip()
    if not t or len(t) > 120:
        return None
    m = re.match(r"^(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:\.(\d+))?[\s、\.]+\S+", t)
    if m:
        depth = 1 + sum(1 for g in m.groups()[1:] if g is not None)
        return min(depth, 4)
    m2 = re.match(r"^[一二三四五六七八九十]+[、\.．]\s*\S+", t)
    if m2:
        return 1
    return None


def matches_any(patterns: Iterable[str], text: str) -> bool:
    for pattern in patterns:
        try:
            if re.search(pattern, text, re.I):
                return True
        except re.error:
            if pattern.lower() in text.lower():
                return True
    return False


def detect_role(para, index: int, doc, constants, profile: Optional[Dict[str, Any]] = None) -> str:
    profile = profile or {}
    detection = profile.get("detection", {}) or {}
    text = clean_text(para.Range.Text)
    st_name = style_name(para)

    # Explicit paragraph-index overrides are the safest override mechanism.
    overrides = detection.get("paragraph_overrides", {}) or {}
    if str(index) in overrides:
        return str(overrides[str(index)])
    if index in overrides:
        return str(overrides[index])

    # Regex overrides are evaluated before heuristics.
    for item in detection.get("regex_overrides", []) or []:
        pattern = item.get("pattern")
        role = item.get("role")
        if pattern and role and re.search(str(pattern), text, re.I):
            return str(role)

    if is_in_table(para.Range, constants):
        return "table_text"

    if not text and paragraph_inline_shape_count(para) == 0:
        return "empty"

    math_count = paragraph_math_count(para)
    if math_count:
        # A paragraph containing equations plus little non-equation text is treated as a block equation.
        # We do not mutate the equation itself; this role controls paragraph spacing/alignment only.
        visible = re.sub(r"[\s\(\)\[\]（）\d\-–—.,，。:：;；]", "", text)
        if len(visible) <= 20:
            return "equation_block"
        return "body"

    if paragraph_inline_shape_count(para) > 0:
        return "figure_paragraph"

    # Built-in/explicit styles first.
    normalized_style = normalize_style_name(st_name)
    if normalized_style in {"title", "标题", "標題"}:
        return "title"

    style_level = heading_level_from_style(st_name)
    if style_level:
        return f"heading_{style_level}"

    outline = paragraph_outline_level(para)
    if outline is not None and 1 <= outline <= 9:
        return f"heading_{outline}"

    # Caption rules.
    fig_patterns = detection.get("figure_caption_patterns") or [
        r"^\s*图\s*\d+",
        r"^\s*圖\s*\d+",
        r"^\s*Fig(?:ure)?\.?\s*\d+",
    ]
    table_patterns = detection.get("table_caption_patterns") or [
        r"^\s*表\s*\d+",
        r"^\s*Table\s*\d+",
    ]
    if matches_any(fig_patterns, text):
        return "figure_caption"
    if matches_any(table_patterns, text):
        return "table_caption"

    # Common front/back matter labels.
    if re.fullmatch(r"\s*(摘要|摘\s*要|abstract)\s*[:：]?\s*", text, re.I):
        return "abstract_heading"
    if re.match(r"^\s*(关键词|關鍵詞|关键字|keywords?)\s*[:：]", text, re.I):
        return "keywords"
    if re.fullmatch(r"\s*(参考文献|參考文獻|references|bibliography)\s*", text, re.I):
        return "references_heading"

    text_level = heading_level_from_text(text)
    if text_level:
        return f"heading_{text_level}"

    # Conservative title fallback: first substantive paragraph, short and centered/large.
    if index <= int(detection.get("title_scan_paragraphs", 5)) and len(text) <= int(
        detection.get("title_max_chars", 80)
    ):
        try:
            alignment = int(para.Format.Alignment)
            center = alignment == int(constants.wdAlignParagraphCenter)
        except Exception:
            center = False
        try:
            size = safe_float(para.Range.Font.Size)
        except Exception:
            size = None
        if center and (size is None or size >= float(detection.get("title_min_pt", 15))):
            return "title"

    return "body"


def role_inventory(doc, constants, profile: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    counts: Counter[str] = Counter()
    items = []
    for i in range(1, safe_int(doc.Paragraphs.Count) + 1):
        para = doc.Paragraphs(i)
        role = detect_role(para, i, doc, constants, profile)
        counts[role] += 1
        if role != "empty":
            items.append(
                {
                    "index": i,
                    "role": role,
                    "style": style_name(para),
                    "outline_level": paragraph_outline_level(para),
                    "equations": paragraph_math_count(para),
                    "inline_shapes": paragraph_inline_shape_count(para),
                    "text": clean_text(para.Range.Text)[:220],
                }
            )
    return {"counts": dict(counts), "paragraphs": items}


def json_dump(data: Any, path: str | os.PathLike[str]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def enum_name_map(constants, names: Iterable[str]) -> Dict[int, str]:
    result = {}
    for name in names:
        try:
            result[int(getattr(constants, name))] = name
        except Exception:
            pass
    return result


def resolve_path(path: str | os.PathLike[str]) -> str:
    return str(Path(path).expanduser().resolve())
