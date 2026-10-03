"""Audit effective paragraph properties in the generated main story, including cells."""
import argparse
import json
from pathlib import Path
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from markdown_to_word import load_profile, ROLES


def effective(par, element, attr):
    nodes = [par._p.pPr]
    style = par.style
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        nodes.append(style.element.pPr)
        style = style.base_style
    for pr in nodes:
        el = pr.find(qn('w:' + element)) if pr is not None else None
        if el is not None and el.get(qn('w:' + attr)) is not None:
            return el.get(qn('w:' + attr))
    return None


def audit(path, profile):
    _, roles = load_profile(profile)
    doc = Document(path)
    names = {name: role for role, name in ROLES.items()}
    errors = []
    counts = {}
    for i, node in enumerate(doc.element.body.iter(qn('w:p'))):
        par = Paragraph(node, doc)
        role = names.get(par.style.name)
        if not role:
            if par.style.name != 'PWD Image':
                errors.append({'paragraph': i, 'error': 'Unmapped paragraph role', 'style': par.style.name})
            continue
        counts[role] = counts.get(role, 0) + 1
        r = roles[role]
        expected = {'center': 'center', 'justify': 'both', 'left': 'left', 'right': 'right'}[r['alignment']]
        tests = [('jc', 'val', expected), ('ind', 'firstLineChars', str(round(r['first_line_chars'] * 100))),
                 ('spacing', 'line', str(round(r['line_spacing'] * 240))), ('spacing', 'lineRule', 'auto'),
                 ('spacing', 'before', str(round(r['before_pt'] * 20))),
                 ('spacing', 'after', str(round(r['after_pt'] * 20)))]
        for el, attr, want in tests:
            got = effective(par, el, attr)
            if got != want:
                errors.append({'paragraph': i, 'role': role, 'property': f'{el}.{attr}', 'expected': want, 'actual': got})
        for attr in ('beforeAutospacing', 'afterAutospacing', 'beforeLines', 'afterLines'):
            got = effective(par, 'spacing', attr)
            if got not in (None, '0', 'false', 'off'):
                errors.append({'paragraph': i, 'role': role, 'property': 'spacing.'+attr, 'actual': got, 'error': 'Overrides point spacing; explicit review required'})
        for attr in ('hanging', 'hangingChars', 'left', 'leftChars', 'right', 'rightChars'):
            got = effective(par, 'ind', attr)
            if got not in (None, '0'):
                errors.append({'paragraph': i, 'role': role, 'property': 'ind.'+attr, 'actual': got})
        if par._p.xpath('./w:pPr/w:numPr'):
            errors.append({'paragraph': i, 'error': 'Numbering indentation requires separate audit'})
    return {'scope': 'main story including table cells; not rendered Windows layout', 'paragraph_roles': counts, 'errors': errors, 'passed': not errors}


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('docx', type=Path)
    ap.add_argument('--profile', required=True, type=Path)
    a = ap.parse_args()
    result = audit(a.docx, a.profile)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result['passed'] else 1)
