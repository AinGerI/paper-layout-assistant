"""Portable structural regression tests; run with Python unittest."""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
import zipfile
from lxml import etree
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('converter', ROOT/'scripts/markdown_to_word.py')
c = importlib.util.module_from_spec(spec); spec.loader.exec_module(c)
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

class ConverterTests(unittest.TestCase):
    def test_role_format_and_content(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp); Image.new('RGB', (100, 40), 'white').save(d/'image.png')
            text = '# 标题\n\n## 1. 第一章\n\n正文 \\(x=1\\) 与 $y=2$。\n\n\\begin{equation}\\left\\{\\begin{array}{l}x+y=3\\\\x\\geq0\\end{array}\\right.\\end{equation}\n\n表1 数据\n\n<table><tr><td>变量</td><td>值</td></tr><tr><td>\\(x\\)</td><td>1</td></tr></table>\n\n![图](image.png)\n\n图1 示例\n'
            (d/'source.md').write_text(text)
            result = c.convert(d/'source.md', d/'out.docx', d/'report', ROOT/'profiles/chinese-paper.json')
            self.assertEqual((result['formulas'], result['tables'], result['images']), (4, 1, 1))
            with zipfile.ZipFile(d/'out.docx') as z:
                document = etree.fromstring(z.read('word/document.xml')); styles = etree.fromstring(z.read('word/styles.xml'))
            normal = styles.xpath('//w:style[@w:styleId="Normal"]', namespaces=NS)[0]
            self.assertEqual(normal.xpath('string(w:pPr/w:ind/@w:firstLineChars)', namespaces=NS), '200')
            self.assertEqual(normal.xpath('string(w:pPr/w:spacing/@w:line)', namespaces=NS), '360')
            for role, name in c.ROLES.items():
                style = styles.xpath('//w:style[w:name/@w:val=$name]', namespaces=NS, name=name.lower() if name.startswith('Heading ') else name)
                self.assertTrue(style, name)
                self.assertFalse(style[0].xpath('w:pPr/w:pBdr|w:pPr/w:shd|w:rPr/w:spacing', namespaces=NS))
                self.assertEqual(style[0].xpath('string(w:rPr/w:color/@w:val)', namespaces=NS), '000000')
                self.assertFalse(any('theme' in a.lower() for el in style[0].xpath('w:rPr/*', namespaces=NS) for a in el.attrib))
                if role != 'body': self.assertEqual(style[0].xpath('string(w:pPr/w:ind/@w:firstLineChars)', namespaces=NS), '0')
            marks = document.xpath('//w:bookmarkStart/@w:name', namespaces=NS)
            self.assertEqual(marks, [f'PWD_EQ_{i:06d}' for i in range(1, 5)])
            for par in document.xpath('//w:p[w:bookmarkStart]', namespaces=NS):
                for prop in ('snapToGrid', 'adjustRightInd'):
                    self.assertEqual(par.xpath('string(w:pPr/w:'+prop+'/@w:val)', namespaces=NS), '0')
            self.assertFalse(document.xpath('//m:oMath', namespaces={'m':'http://schemas.openxmlformats.org/officeDocument/2006/math'}))
            self.assertIn('array}{l}', ''.join(document.xpath('//w:t/text()', namespaces=NS)))
            with self.assertRaises(ValueError): c.convert(d/'source.md', d/'out.docx', d/'second', ROOT/'profiles/chinese-paper.json')

    def test_alternate_profile(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp); p = json.loads((ROOT/'profiles/chinese-paper.json').read_text())
            p['fonts']['east_asia'] = '仿宋'; p['roles']['body'].update(size_pt=14, line_spacing=2, first_line_chars=3)
            p['roles']['caption']['line_spacing'] = 2
            p['page']['left_cm'] = 3; p['table'].update(border='grid', header_repeat=False, allow_row_split=True)
            (d/'profile.json').write_text(json.dumps(p)); (d/'s.md').write_text('# 标题\n\n正文。\n\n表1 内容\n\n<table><tr><td>a</td></tr></table>')
            c.convert(d/'s.md', d/'o.docx', d/'report', d/'profile.json')
            with zipfile.ZipFile(d/'o.docx') as z:
                st = etree.fromstring(z.read('word/styles.xml')); doc = etree.fromstring(z.read('word/document.xml'))
            for name in ('Normal', 'PWDCaption'):
                node = st.xpath('//w:style[@w:styleId=$name]', namespaces=NS, name=name)[0]
                self.assertEqual(node.xpath('string(w:rPr/w:sz/@w:val)', namespaces=NS), '28')
                self.assertEqual(node.xpath('string(w:rPr/w:rFonts/@w:eastAsia)', namespaces=NS), '仿宋')
            self.assertEqual(st.xpath('string(//w:style[@w:styleId="Normal"]/w:pPr/w:ind/@w:firstLineChars)', namespaces=NS), '300')
            self.assertFalse(doc.xpath('//w:tblHeader|//w:cantSplit', namespaces=NS))
            self.assertEqual(doc.xpath('string(//w:tcBorders/w:left/@w:val)', namespaces=NS), 'single')

    def test_explicit_role_overrides(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp); p = json.loads((ROOT/'profiles/chinese-paper.json').read_text())
            p['roles']['caption'].update(size_pt=10, east_asia='楷体')
            p['roles']['title']['first_line_chars'] = 1
            (d/'p.json').write_text(json.dumps(p)); (d/'s.md').write_text('# 标题\n\n正文。')
            c.convert(d/'s.md', d/'o.docx', d/'report', d/'p.json')
            with zipfile.ZipFile(d/'o.docx') as z: st = etree.fromstring(z.read('word/styles.xml'))
            self.assertEqual(st.xpath('string(//w:style[@w:styleId="Title"]/w:pPr/w:ind/@w:firstLineChars)', namespaces=NS), '100')
            self.assertEqual(st.xpath('string(//w:style[@w:styleId="PWDCaption"]/w:rPr/w:sz/@w:val)', namespaces=NS), '20')

    def test_reject_unsupported(self):
        for source in ('正文 $broken', '- 列表', '|a|b|', '```\nx\n```', '正文 **粗体**', '<table><tr><td colspan="2">x</td></tr></table>', '\\begin{equation}x', '# 标题\n正文', '正文\n## 章节', '正文 \\(x\\) 和 \\(bad'):
            with self.subTest(source=source), self.assertRaises(ValueError): c.parse(source)

    def test_compatibility_narrow(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp); (d/'s.md').write_text(r'正文 \(\widetilde{\omega }+\widetilde{x}+\widetilde{\omega_i}\)。')
            r = c.convert(d/'s.md', d/'o.docx', d/'report', ROOT/'profiles/chinese-paper.json')
            eq = json.loads((d/'report/equations.json').read_text())[0]['code']
            self.assertIn(r'\tilde{\omega }', eq); self.assertIn(r'\widetilde{x}', eq); self.assertIn(r'\widetilde{\omega_i}', eq)
            self.assertEqual(r['compatibility_changes'][0]['occurrences'], 1)

if __name__ == '__main__': unittest.main()
