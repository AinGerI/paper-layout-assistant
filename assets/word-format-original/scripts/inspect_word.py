from __future__ import annotations

import argparse
from collections import Counter
from pathlib import Path

from word_common import (
    connect_word,
    document_snapshot,
    json_dump,
    load_profile,
    resolve_path,
    role_inventory,
)


def build_markdown(data: dict) -> str:
    snap = data["snapshot"]
    counts = data["roles"]["counts"]
    lines = [
        "# Word 文档结构分析报告",
        "",
        f"- 文件：`{data['document']}`",
        f"- 段落：{snap['paragraphs']}",
        f"- 表格：{snap['tables']}",
        f"- 公式：{snap['equations']}",
        f"- 行内图片：{snap['inline_shapes']}",
        f"- 浮动图形：{snap['floating_shapes']}",
        f"- 域：{snap['fields']}",
        f"- 分节：{snap['sections']}",
        "",
        "## 语义角色识别",
        "",
    ]
    for role, count in sorted(counts.items()):
        lines.append(f"- `{role}`：{count}")

    warnings = []
    if counts.get("title", 0) == 0:
        warnings.append("未可靠识别到标题；建议检查模板样式或增加 paragraph_overrides。")
    if counts.get("title", 0) > 2:
        warnings.append("识别到多个 title，可能存在标题误判。")
    if snap["tables"] and counts.get("table_caption", 0) == 0:
        warnings.append("存在表格但未识别到表题。")
    if (snap["inline_shapes"] + snap["floating_shapes"]) and counts.get("figure_caption", 0) == 0:
        warnings.append("存在图片/图形但未识别到图题。")

    lines += ["", "## 建议人工检查", ""]
    if warnings:
        lines.extend([f"- {w}" for w in warnings])
    else:
        lines.append("- 未发现明显的结构识别异常。")

    lines += ["", "## 段落预览", ""]
    for item in data["roles"]["paragraphs"][:120]:
        text = item["text"].replace("`", "'")
        lines.append(
            f"- P{item['index']} · `{item['role']}` · `{item['style']}` · {text}"
        )
    if len(data["roles"]["paragraphs"]) > 120:
        lines.append("- ……（完整列表见 JSON 报告）")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect a .docx using Microsoft Word COM without modifying it.")
    parser.add_argument("document", help="Input .docx path")
    parser.add_argument("--profile", help="Optional YAML profile to improve semantic detection")
    parser.add_argument("--json", dest="json_path", help="Output JSON path")
    parser.add_argument("--md", dest="md_path", help="Output Markdown report path")
    parser.add_argument("--visible", action="store_true", help="Show Microsoft Word while inspecting")
    args = parser.parse_args()

    document = resolve_path(args.document)
    profile = load_profile(args.profile) if args.profile else None
    json_path = resolve_path(args.json_path or (str(Path(document).with_suffix("")) + "_inspection.json"))
    md_path = resolve_path(args.md_path or (str(Path(document).with_suffix("")) + "_inspection.md"))

    app = None
    doc = None
    try:
        app, constants = connect_word(visible=args.visible)
        doc = app.Documents.Open(document, ReadOnly=True, AddToRecentFiles=False)
        data = {
            "document": document,
            "snapshot": document_snapshot(doc),
            "roles": role_inventory(doc, constants, profile),
        }
        json_dump(data, json_path)
        Path(md_path).write_text(build_markdown(data), encoding="utf-8")
        print(json_path)
        print(md_path)
        return 0
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
