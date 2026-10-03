"""Real Word COM end-to-end test; all artifacts remain under --output-dir."""
import argparse
import json
import subprocess
import sys
from pathlib import Path
from word_common import connect_word, document_snapshot


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--output-dir', required=True)
    a = p.parse_args()
    out = Path(a.output_dir).resolve()
    out.mkdir(parents=True, exist_ok=False)
    source = out / '中文 测试原稿.docx'
    scripts = Path(__file__).parent
    profile = scripts.parent / 'profiles' / 'demo-chinese-academic.yaml'
    app, c = connect_word()
    doc = None
    try:
        print('Word version:', app.Version, flush=True)
        doc = app.Documents.Add()
        lines = ['通用学术排版测试', '第一段正文，用于验证中文字体和内容保护。',
                 '第二段正文包含 English 和数字 123，排版不能改变内容。',
                 '1 一级标题', '1.1 二级标题', '表 1 示例数据', '',
                 '图 1 图题占位段落', 'x^2+y^2=z^2', '域与图形测试', '']
        doc.Content.Text = '\r'.join(lines)
        doc.Content.Font.NameFarEast = '微软雅黑'
        doc.Content.Font.NameAscii = 'Arial'
        doc.Content.Font.Size = 11
        doc.Paragraphs(1).Range.Style = -63
        doc.Paragraphs(4).Range.Style = -2
        doc.Paragraphs(5).Range.Style = -3
        equation = doc.Paragraphs(9).Range.Duplicate
        equation.End -= 1
        doc.OMaths.Add(equation)
        doc.OMaths(1).BuildUp()
        table = doc.Tables.Add(doc.Paragraphs(7).Range, 2, 2)
        for row, values in enumerate([['项目', '数值'], ['样本 A', '123']], 1):
            for col, value in enumerate(values, 1):
                table.Cell(row,col).Range.Text = value
        table.Rows.Alignment = 0
        table.Range.ParagraphFormat.Alignment = 0
        table.Range.Cells.VerticalAlignment = 0
        table.Rows.HeightRule = 0
        # Include nonzero fields, inline pictures and floating shapes.
        import struct
        bmp = out / 'image.bmp'
        pixels = b'\x00\x00\xff\x00' * 2
        bmp.write_bytes(b'BM' + struct.pack('<IHHI', 62,0,0,54) + struct.pack('<IiiHHIIiiII',40,1,2,1,24,0,8,0,0,0,0) + pixels)
        anchor = doc.Range(doc.Content.End-1, doc.Content.End-1)
        doc.InlineShapes.AddPicture(str(bmp), False, True, anchor)
        anchor = doc.Range(doc.Content.End-1, doc.Content.End-1)
        doc.Fields.Add(anchor, -1, 'QUOTE "Stable field"', False)
        doc.Shapes.AddShape(1, 50,50,30,20,doc.Paragraphs(2).Range)
        doc.SaveAs2(str(source), FileFormat=16)
        doc.Close(False)
        doc = None
    finally:
        if doc is not None:
            doc.Close(False)
        app.Quit(SaveChanges=0)
    original_bytes = source.read_bytes()
    commands = [
        ['inspect_word.py', str(source), '--profile', str(profile)],
        ['format_word.py', str(source), '--profile', str(profile)],
        ['validate_word.py', str(source), str(source.with_stem(source.stem+'_formatted'))],
        ['extract_profile.py', str(source), '--output', str(out/'提取 Profile.yaml')],
    ]
    for args in commands:
        result = subprocess.run([sys.executable, str(scripts/args[0]), *args[1:]], capture_output=True, text=True, encoding='utf-8', timeout=180)
        (out/(args[0]+'.log')).write_text(result.stdout+'\n'+result.stderr, encoding='utf-8')
        print(args[0], result.returncode, flush=True)
        if result.returncode:
            raise RuntimeError(result.stderr or result.stdout)
    formatted = source.with_stem(source.stem+'_formatted')
    app, c = connect_word()
    snapshots = []
    formats = []
    try:
        for path in (source, formatted):
            doc = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
            snapshots.append(document_snapshot(doc))
            table = doc.Tables(1)
            formats.append({
                'body_east_asia': doc.Paragraphs(2).Range.Font.NameFarEast,
                'body_latin': doc.Paragraphs(2).Range.Font.NameAscii,
                'body_size': doc.Paragraphs(2).Range.Font.Size,
                'table_alignment': table.Rows.Alignment,
                'cell_horizontal': table.Cell(1,1).Range.ParagraphFormat.Alignment,
                'cell_vertical': table.Cell(1,1).VerticalAlignment,
                'cell_space_before': table.Cell(1,1).Range.ParagraphFormat.SpaceBefore,
                'cell_space_after': table.Cell(1,1).Range.ParagraphFormat.SpaceAfter,
                'cell_line_spacing': table.Cell(1,1).Range.ParagraphFormat.LineSpacingRule,
                'row_height_rule': table.Rows.HeightRule,
                'padding_top': table.TopPadding,
                'border_top': table.Borders(-1).LineStyle,
                'border_left': table.Borders(-2).LineStyle,
                'equation_font': doc.OMaths(1).Range.Font.Name,
                'equation_line_rule': doc.OMaths(1).Range.ParagraphFormat.LineSpacingRule,
            })
            doc.Close(False)
            doc = None
    finally:
        if doc is not None:
            doc.Close(False)
        app.Quit(SaveChanges=0)
    assert source.read_bytes() == original_bytes, 'Source bytes changed'
    assert snapshots[0] == snapshots[1], 'Content or objects changed'
    for key in ('tables','equations','inline_shapes','floating_shapes','fields'):
        assert snapshots[1][key] > 0, key
    assert formats[0] != formats[1], 'No format change'
    expected = dict(body_east_asia='宋体', body_latin='Times New Roman', body_size=12,
                    table_alignment=1, cell_horizontal=1, cell_vertical=1,
                    cell_space_before=0, cell_space_after=0, cell_line_spacing=0,
                    row_height_rule=0, padding_top=2, border_top=1, border_left=0,
                    equation_line_rule=3)
    for key, value in expected.items():
        assert formats[1][key] == value, (key, formats[1][key], value)
    assert formats[0]['equation_font'] == formats[1]['equation_font']
    result = {'passed': True, 'source_unchanged': True, 'source': str(source), 'output': str(formatted),
              'counts': {k:v for k,v in snapshots[1].items() if k != 'body_text'},
              'before_format': formats[0], 'after_format': formats[1]}
    (out/'test_result.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False,indent=2))


if __name__ == '__main__':
    main()
