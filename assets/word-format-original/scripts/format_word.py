from __future__ import annotations

import argparse
import json
import math
import os
import shutil
from pathlib import Path
from typing import Any, Dict, Optional

from word_common import (
    clean_text,
    connect_word,
    detect_role,
    document_snapshot,
    json_dump,
    load_profile,
    resolve_path,
    safe_float,
    safe_int,
)


def get_constant(constants, name: str, fallback: Optional[int] = None):
    try:
        return getattr(constants, name)
    except Exception:
        if fallback is not None:
            return fallback
        raise


def set_if_present(obj, attr: str, value: Any) -> None:
    if value is None:
        return
    try:
        setattr(obj, attr, value)
    except Exception:
        pass


def text_range_without_para_mark(para):
    rng = para.Range.Duplicate
    try:
        if rng.End > rng.Start:
            rng.End = rng.End - 1
    except Exception:
        pass
    return rng


def apply_font(rng, spec: Dict[str, Any]) -> None:
    font = rng.Font
    east_asia = spec.get("east_asia") or spec.get("chinese")
    latin = spec.get("latin") or spec.get("western")
    all_font = spec.get("name")
    size_pt = spec.get("size_pt")

    if all_font:
        set_if_present(font, "Name", str(all_font))
    if east_asia:
        set_if_present(font, "NameFarEast", str(east_asia))
    if latin:
        set_if_present(font, "NameAscii", str(latin))
        set_if_present(font, "NameOther", str(latin))
        set_if_present(font, "NameBi", str(latin))
    if size_pt is not None:
        set_if_present(font, "Size", float(size_pt))
    if "bold" in spec:
        set_if_present(font, "Bold", bool(spec.get("bold")))
    if "italic" in spec:
        set_if_present(font, "Italic", bool(spec.get("italic")))
    if "underline" in spec:
        # Word accepts a WdUnderline enum; for a simple boolean use single/none.
        set_if_present(font, "Underline", 1 if spec.get("underline") else 0)


def alignment_constant(value: str, constants):
    mapping = {
        "left": "wdAlignParagraphLeft",
        "center": "wdAlignParagraphCenter",
        "right": "wdAlignParagraphRight",
        "justify": "wdAlignParagraphJustify",
        "distributed": "wdAlignParagraphDistribute",
    }
    name = mapping.get(str(value).lower())
    return getattr(constants, name) if name else None


def apply_line_spacing(fmt, spec: Dict[str, Any], app, constants) -> None:
    spacing = spec.get("line_spacing")
    if spacing is None:
        return

    if isinstance(spacing, (int, float)):
        # Numeric values are interpreted as multiple line spacing, e.g. 1.25.
        fmt.LineSpacingRule = constants.wdLineSpaceMultiple
        fmt.LineSpacing = app.LinesToPoints(float(spacing))
        return

    if isinstance(spacing, str):
        key = spacing.strip().lower()
        if key in {"single", "1", "1.0"}:
            fmt.LineSpacingRule = constants.wdLineSpaceSingle
            return
        if key in {"1.5", "one_point_five"}:
            fmt.LineSpacingRule = constants.wdLineSpace1pt5
            return
        if key in {"double", "2", "2.0"}:
            fmt.LineSpacingRule = constants.wdLineSpaceDouble
            return

    if isinstance(spacing, dict):
        rule = str(spacing.get("rule", "")).lower()
        value = spacing.get("value")
        if rule == "exact":
            fmt.LineSpacingRule = constants.wdLineSpaceExactly
            if value is not None:
                fmt.LineSpacing = float(value)
        elif rule in {"at_least", "atleast"}:
            fmt.LineSpacingRule = constants.wdLineSpaceAtLeast
            if value is not None:
                fmt.LineSpacing = float(value)
        elif rule in {"multiple", "multi"}:
            fmt.LineSpacingRule = constants.wdLineSpaceMultiple
            if value is not None:
                fmt.LineSpacing = app.LinesToPoints(float(value))
        elif rule == "single":
            fmt.LineSpacingRule = constants.wdLineSpaceSingle
        elif rule in {"1.5", "one_point_five"}:
            fmt.LineSpacingRule = constants.wdLineSpace1pt5
        elif rule == "double":
            fmt.LineSpacingRule = constants.wdLineSpaceDouble


def apply_paragraph(para, spec: Dict[str, Any], app, constants) -> None:
    fmt = para.Format
    para_spec = spec.get("paragraph", spec)

    if para_spec.get("alignment"):
        val = alignment_constant(str(para_spec["alignment"]), constants)
        if val is not None:
            fmt.Alignment = val
    if para_spec.get("space_before_pt") is not None:
        fmt.SpaceBefore = float(para_spec["space_before_pt"])
    if para_spec.get("space_after_pt") is not None:
        fmt.SpaceAfter = float(para_spec["space_after_pt"])
    if para_spec.get("left_indent_pt") is not None:
        fmt.LeftIndent = float(para_spec["left_indent_pt"])
    if para_spec.get("right_indent_pt") is not None:
        fmt.RightIndent = float(para_spec["right_indent_pt"])
    if para_spec.get("first_line_indent_chars") is not None:
        try:
            fmt.CharacterUnitFirstLineIndent = float(para_spec["first_line_indent_chars"])
        except Exception:
            pass
    elif para_spec.get("first_line_indent_pt") is not None:
        fmt.FirstLineIndent = float(para_spec["first_line_indent_pt"])

    apply_line_spacing(fmt, para_spec, app, constants)

    if "keep_with_next" in para_spec:
        fmt.KeepWithNext = bool(para_spec["keep_with_next"])
    if "keep_together" in para_spec:
        fmt.KeepTogether = bool(para_spec["keep_together"])
    if "page_break_before" in para_spec:
        fmt.PageBreakBefore = bool(para_spec["page_break_before"])
    if "widow_control" in para_spec:
        fmt.WidowControl = bool(para_spec["widow_control"])

    font_spec = spec.get("font") or {}
    if font_spec:
        apply_font(text_range_without_para_mark(para), font_spec)


def apply_page_setup(doc, profile: Dict[str, Any], app, constants) -> None:
    page = profile.get("page") or {}
    if not page or not page.get("enabled", True):
        return
    for i in range(1, safe_int(doc.Sections.Count) + 1):
        section = doc.Sections(i)
        setup = section.PageSetup
        if page.get("orientation"):
            ori = str(page["orientation"]).lower()
            if ori == "portrait":
                setup.Orientation = constants.wdOrientPortrait
            elif ori == "landscape":
                setup.Orientation = constants.wdOrientLandscape

        # Units in cm are friendlier for Chinese academic formatting.
        cm_map = {
            "top_margin_cm": "TopMargin",
            "bottom_margin_cm": "BottomMargin",
            "left_margin_cm": "LeftMargin",
            "right_margin_cm": "RightMargin",
            "header_distance_cm": "HeaderDistance",
            "footer_distance_cm": "FooterDistance",
            "gutter_cm": "Gutter",
        }
        for key, attr in cm_map.items():
            if page.get(key) is not None:
                setattr(setup, attr, app.CentimetersToPoints(float(page[key])))

        if page.get("paper_width_cm") is not None:
            setup.PageWidth = app.CentimetersToPoints(float(page["paper_width_cm"]))
        if page.get("paper_height_cm") is not None:
            setup.PageHeight = app.CentimetersToPoints(float(page["paper_height_cm"]))


def line_width_constant(points: float, constants):
    candidates = [
        (0.25, "wdLineWidth025pt"),
        (0.5, "wdLineWidth050pt"),
        (0.75, "wdLineWidth075pt"),
        (1.0, "wdLineWidth100pt"),
        (1.5, "wdLineWidth150pt"),
        (2.25, "wdLineWidth225pt"),
        (3.0, "wdLineWidth300pt"),
        (4.5, "wdLineWidth450pt"),
        (6.0, "wdLineWidth600pt"),
    ]
    nearest = min(candidates, key=lambda x: abs(x[0] - float(points)))
    return getattr(constants, nearest[1])


def set_border(border, constants, style: str = "single", width_pt: float = 0.75) -> None:
    style_map = {
        "single": "wdLineStyleSingle",
        "double": "wdLineStyleDouble",
        "none": "wdLineStyleNone",
    }
    border.LineStyle = getattr(constants, style_map.get(style, "wdLineStyleSingle"))
    if style != "none":
        border.LineWidth = line_width_constant(width_pt, constants)


def clear_table_borders(table, constants) -> None:
    for name in [
        "wdBorderTop",
        "wdBorderLeft",
        "wdBorderBottom",
        "wdBorderRight",
        "wdBorderHorizontal",
        "wdBorderVertical",
        "wdBorderDiagonalDown",
        "wdBorderDiagonalUp",
    ]:
        try:
            b = table.Borders(getattr(constants, name))
            b.LineStyle = constants.wdLineStyleNone
        except Exception:
            pass


def apply_table_borders(table, border_spec: Dict[str, Any], constants) -> None:
    border_type = str(border_spec.get("type", "retain")).lower()
    if border_type == "retain":
        return
    if border_type == "none":
        clear_table_borders(table, constants)
        return
    if border_type in {"grid", "full_grid"}:
        width = float(border_spec.get("width_pt", 0.5))
        style = str(border_spec.get("style", "single"))
        for name in [
            "wdBorderTop",
            "wdBorderLeft",
            "wdBorderBottom",
            "wdBorderRight",
            "wdBorderHorizontal",
            "wdBorderVertical",
        ]:
            try:
                set_border(table.Borders(getattr(constants, name)), constants, style, width)
            except Exception:
                pass
        return
    if border_type == "three_line":
        clear_table_borders(table, constants)
        top = float(border_spec.get("top_pt", 1.5))
        header = float(border_spec.get("header_bottom_pt", 0.75))
        bottom = float(border_spec.get("bottom_pt", 1.5))
        set_border(table.Borders(constants.wdBorderTop), constants, "single", top)
        set_border(table.Borders(constants.wdBorderBottom), constants, "single", bottom)
        if safe_int(table.Rows.Count) >= 1:
            try:
                set_border(table.Rows(1).Borders(constants.wdBorderBottom), constants, "single", header)
            except Exception:
                pass


def apply_table(table, spec: Dict[str, Any], app, constants) -> None:
    if not spec or not spec.get("enabled", True):
        return

    # Table positioning.
    align = str(spec.get("alignment", "")).lower()
    table_align = {
        "left": "wdAlignRowLeft",
        "center": "wdAlignRowCenter",
        "right": "wdAlignRowRight",
    }
    if align in table_align:
        try:
            table.Rows.Alignment = getattr(constants, table_align[align])
        except Exception:
            pass

    if spec.get("autofit") is not None:
        try:
            table.AllowAutoFit = bool(spec.get("autofit"))
        except Exception:
            pass

    width_cm = spec.get("preferred_width_cm")
    if width_cm is not None:
        try:
            table.PreferredWidthType = constants.wdPreferredWidthPoints
            table.PreferredWidth = app.CentimetersToPoints(float(width_cm))
        except Exception:
            pass

    # Cell padding.
    cell = spec.get("cell") or {}
    for key, attr in [
        ("margin_top_pt", "TopPadding"),
        ("margin_bottom_pt", "BottomPadding"),
        ("margin_left_pt", "LeftPadding"),
        ("margin_right_pt", "RightPadding"),
    ]:
        if cell.get(key) is not None:
            try:
                setattr(table, attr, float(cell[key]))
            except Exception:
                pass

    valign = str(cell.get("vertical_alignment", "")).lower()
    valign_map = {
        "top": "wdCellAlignVerticalTop",
        "center": "wdCellAlignVerticalCenter",
        "bottom": "wdCellAlignVerticalBottom",
    }
    if valign in valign_map:
        try:
            table.Range.Cells.VerticalAlignment = getattr(constants, valign_map[valign])
        except Exception:
            pass

    # Row pagination and height. Auto is safest for formula-heavy tables.
    row = spec.get("row") or {}
    if row.get("allow_break_across_pages") is not None:
        try:
            table.Rows.AllowBreakAcrossPages = bool(row["allow_break_across_pages"])
        except Exception:
            pass
    if row.get("repeat_header") and safe_int(table.Rows.Count) >= 1:
        try:
            table.Rows(1).HeadingFormat = True
        except Exception:
            pass

    rule = str(row.get("height_rule", "auto")).lower()
    try:
        if rule == "auto":
            table.Rows.HeightRule = constants.wdRowHeightAuto
        elif rule in {"at_least", "atleast"}:
            table.Rows.HeightRule = constants.wdRowHeightAtLeast
            if row.get("height_pt") is not None:
                table.Rows.Height = float(row["height_pt"])
        elif rule == "exact":
            # Exact height is supported, but use it only when the user explicitly requests it.
            table.Rows.HeightRule = constants.wdRowHeightExactly
            if row.get("height_pt") is not None:
                table.Rows.Height = float(row["height_pt"])
    except Exception:
        pass

    # Apply text formatting to the complete table range.
    if spec.get("font"):
        apply_font(table.Range, spec["font"])
    paragraph_spec = spec.get("paragraph") or {}
    if paragraph_spec:
        # Range.ParagraphFormat exposes the same formatting properties.
        fmt = table.Range.ParagraphFormat
        if paragraph_spec.get("alignment"):
            val = alignment_constant(str(paragraph_spec["alignment"]), constants)
            if val is not None:
                fmt.Alignment = val
        if paragraph_spec.get("space_before_pt") is not None:
            fmt.SpaceBefore = float(paragraph_spec["space_before_pt"])
        if paragraph_spec.get("space_after_pt") is not None:
            fmt.SpaceAfter = float(paragraph_spec["space_after_pt"])
        if paragraph_spec.get("first_line_indent_chars") is not None:
            try:
                fmt.CharacterUnitFirstLineIndent = float(paragraph_spec["first_line_indent_chars"])
            except Exception:
                pass
        apply_line_spacing(fmt, paragraph_spec, app, constants)

    apply_table_borders(table, spec.get("border") or {}, constants)


def apply_figures(doc, spec: Dict[str, Any], app, constants) -> Dict[str, int]:
    stats = {"inline_shapes_resized": 0, "floating_shapes_resized": 0}
    if not spec or not spec.get("enabled", True):
        return stats
    max_width_cm = spec.get("max_width_cm")
    exact_width_cm = spec.get("width_cm")

    for i in range(1, safe_int(doc.InlineShapes.Count) + 1):
        shape = doc.InlineShapes(i)
        try:
            if spec.get("lock_aspect_ratio", True):
                shape.LockAspectRatio = True
            target = None
            if exact_width_cm is not None:
                target = app.CentimetersToPoints(float(exact_width_cm))
            elif max_width_cm is not None and float(shape.Width) > app.CentimetersToPoints(float(max_width_cm)):
                target = app.CentimetersToPoints(float(max_width_cm))
            if target is not None:
                shape.Width = target
                stats["inline_shapes_resized"] += 1
            if str(spec.get("paragraph_alignment", "")).lower() == "center":
                shape.Range.ParagraphFormat.Alignment = constants.wdAlignParagraphCenter
        except Exception:
            pass

    # Floating shapes are more layout-sensitive. Only resize when explicitly requested.
    if spec.get("resize_floating_shapes"):
        for i in range(1, safe_int(doc.Shapes.Count) + 1):
            shape = doc.Shapes(i)
            try:
                if spec.get("lock_aspect_ratio", True):
                    shape.LockAspectRatio = True
                target = None
                if exact_width_cm is not None:
                    target = app.CentimetersToPoints(float(exact_width_cm))
                elif max_width_cm is not None and float(shape.Width) > app.CentimetersToPoints(float(max_width_cm)):
                    target = app.CentimetersToPoints(float(max_width_cm))
                if target is not None:
                    shape.Width = target
                    stats["floating_shapes_resized"] += 1
            except Exception:
                pass
    return stats


def compare_snapshots(before: Dict[str, Any], after: Dict[str, Any]) -> Dict[str, Any]:
    critical_keys = [
        "text_sha256",
        "text_length",
        "tables",
        "inline_shapes",
        "floating_shapes",
        "equations",
        "fields",
        "sections",
    ]
    diffs = {}
    for key in critical_keys:
        if before.get(key) != after.get(key):
            diffs[key] = {"before": before.get(key), "after": after.get(key)}
    return {"passed": not diffs, "differences": diffs}


def build_md_report(report: Dict[str, Any]) -> str:
    lines = [
        "# Word 自动排版报告",
        "",
        f"- 源文件：`{report['source']}`",
        f"- 输出文件：`{report['output']}`",
        f"- Profile：`{report['profile']}`",
        "",
        "## 本次处理",
        "",
    ]
    for role, count in sorted(report.get("formatted_roles", {}).items()):
        lines.append(f"- `{role}`：{count} 个段落")
    lines.append(f"- 表格：{report.get('tables_formatted', 0)}")
    fig = report.get("figures", {})
    lines.append(f"- 调整行内图片：{fig.get('inline_shapes_resized', 0)}")
    lines.append(f"- 调整浮动图片：{fig.get('floating_shapes_resized', 0)}")

    lines += ["", "## 内容完整性检查", ""]
    safety = report["integrity_check"]
    if safety["passed"]:
        lines.append("- ✅ 文本哈希、公式/表格/图片/域数量在排版前后保持一致。")
    else:
        lines.append("- ⚠️ 检测到内容结构发生变化，请不要直接提交该文件。")
        for key, diff in safety["differences"].items():
            lines.append(f"  - `{key}`：{diff['before']} → {diff['after']}")

    warnings = report.get("warnings") or []
    lines += ["", "## 需要人工检查", ""]
    if warnings:
        lines.extend([f"- {w}" for w in warnings])
    else:
        lines.append("- 未发现脚本层面的明显异常；仍建议在 Word 中快速检查分页、浮动图形和复杂合并表格。")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="Format a .docx using Microsoft Word COM and a YAML profile.")
    parser.add_argument("document", help="Input .docx path")
    parser.add_argument("--profile", required=True, help="YAML format profile")
    parser.add_argument("--output", help="Output .docx path (default: *_formatted.docx)")
    parser.add_argument("--report", help="Markdown report path")
    parser.add_argument("--json-report", help="JSON report path")
    parser.add_argument("--visible", action="store_true", help="Show Word while formatting")
    parser.add_argument("--force", action="store_true", help="Allow replacing an existing output file; never overwrites source")
    args = parser.parse_args()

    source = resolve_path(args.document)
    profile_path = resolve_path(args.profile)
    profile = load_profile(profile_path)
    source_path = Path(source)
    output = resolve_path(args.output or str(source_path.with_name(source_path.stem + "_formatted.docx")))
    if os.path.normcase(source) == os.path.normcase(output):
        raise SystemExit("Refusing to overwrite the source document. Choose a different --output path.")
    if Path(output).exists() and not args.force:
        raise SystemExit(f"Output already exists: {output}. Use --force or choose another path.")

    report_path = resolve_path(args.report or str(Path(output).with_suffix("")) + "_report.md")
    json_report_path = resolve_path(args.json_report or str(Path(output).with_suffix("")) + "_report.json")

    Path(output).parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, output)

    app = None
    doc = None
    report: Dict[str, Any] = {
        "source": source,
        "output": output,
        "profile": profile_path,
        "formatted_roles": {},
        "tables_formatted": 0,
        "figures": {},
        "warnings": [],
    }

    try:
        app, constants = connect_word(visible=args.visible)
        doc = app.Documents.Open(output, ReadOnly=False, AddToRecentFiles=False)
        before = document_snapshot(doc)

        apply_page_setup(doc, profile, app, constants)

        roles = profile.get("roles") or {}
        role_counts: Dict[str, int] = {}
        for i in range(1, safe_int(doc.Paragraphs.Count) + 1):
            para = doc.Paragraphs(i)
            role = detect_role(para, i, doc, constants, profile)
            spec = roles.get(role)
            if spec and spec.get("enabled", True):
                apply_paragraph(para, spec, app, constants)
                role_counts[role] = role_counts.get(role, 0) + 1
        report["formatted_roles"] = role_counts

        table_spec = profile.get("table") or {}
        if table_spec and table_spec.get("enabled", True):
            for i in range(1, safe_int(doc.Tables.Count) + 1):
                apply_table(doc.Tables(i), table_spec, app, constants)
                report["tables_formatted"] += 1

        report["figures"] = apply_figures(doc, profile.get("figure") or {}, app, constants)

        # Repaginate only; do not update fields/cross-references automatically.
        if (profile.get("document") or {}).get("repaginate", True):
            try:
                doc.Repaginate()
            except Exception:
                pass

        after = document_snapshot(doc)
        integrity = compare_snapshots(before, after)
        report["before"] = before
        report["after"] = after
        report["integrity_check"] = integrity

        if not integrity["passed"]:
            report["warnings"].append(
                "排版前后内容结构不一致。脚本没有自动回滚输出文件，请以原文件为准并检查报告中的差异。"
            )

        # Flag layout-sensitive objects.
        if after.get("floating_shapes", 0):
            report["warnings"].append(
                f"检测到 {after['floating_shapes']} 个浮动图形；浮动对象受锚点和分页影响，建议人工检查位置。"
            )
        if after.get("tables", 0):
            try:
                merged_or_complex = 0
                for ti in range(1, safe_int(doc.Tables.Count) + 1):
                    tbl = doc.Tables(ti)
                    # A simple rectangular table has Rows*Columns cells. Merged tables often differ or raise.
                    try:
                        expected = safe_int(tbl.Rows.Count) * safe_int(tbl.Columns.Count)
                        actual = safe_int(tbl.Range.Cells.Count)
                        if expected != actual:
                            merged_or_complex += 1
                    except Exception:
                        merged_or_complex += 1
                if merged_or_complex:
                    report["warnings"].append(
                        f"检测到约 {merged_or_complex} 个包含合并单元格或复杂网格的表格，建议人工检查边框与分页。"
                    )
            except Exception:
                pass

        doc.Save()
        Path(report_path).write_text(build_md_report(report), encoding="utf-8")
        json_dump(report, json_report_path)
        print(output)
        print(report_path)
        print(json_report_path)
        return 0 if integrity["passed"] else 2
    except Exception:
        # The source is safe because formatting occurs only on a copied output file.
        raise
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=True)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
