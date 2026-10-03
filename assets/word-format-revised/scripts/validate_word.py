from __future__ import annotations
import argparse
from pathlib import Path
from word_common import connect_word, document_snapshot, resolve_path, json_dump, reserve_paths


def main():
    p=argparse.ArgumentParser(description='Compare source and formatted Word documents for content integrity.')
    p.add_argument('source'); p.add_argument('formatted'); p.add_argument('--json'); p.add_argument('--md'); p.add_argument('--visible',action='store_true')
    a=p.parse_args(); src=resolve_path(a.source); out=resolve_path(a.formatted)
    jp=resolve_path(a.json or str(Path(out).with_suffix(''))+'_validation.json')
    mp=resolve_path(a.md or str(Path(out).with_suffix(''))+'_validation.md')
    reserve_paths([src,out],[jp,mp])
    app=None; d1=d2=None
    try:
        app,c=connect_word(a.visible)
        d1=app.Documents.Open(src,ReadOnly=True,AddToRecentFiles=False); s1=document_snapshot(d1); d1.Close(False); d1=None
        d2=app.Documents.Open(out,ReadOnly=True,AddToRecentFiles=False); s2=document_snapshot(d2)
        keys=['body_text','text_length','tables','inline_shapes','floating_shapes','equations','fields','sections']
        diffs={k:{'source':s1[k],'formatted':s2[k]} for k in keys if s1[k]!=s2[k]}
        data={'source':src,'formatted':out,'passed':not diffs,'differences':diffs,'source_snapshot':s1,'formatted_snapshot':s2}
        jp=resolve_path(a.json or str(Path(out).with_suffix(''))+'_validation.json')
        mp=resolve_path(a.md or str(Path(out).with_suffix(''))+'_validation.md')
        json_dump(data,jp)
        lines=['# Word 内容完整性检查','',f"- 结果：{'✅ 通过' if not diffs else '⚠️ 存在差异'}",'', '## 差异','']
        if diffs:
            for k,v in diffs.items(): lines.append(f"- `{k}`：{v['source']} → {v['formatted']}")
        else: lines.append('- 正文文本以及公式、表格、图片、域和分节数量一致。')
        Path(mp).write_text('\n'.join(lines)+'\n',encoding='utf-8')
        print(jp); print(mp); return 0 if not diffs else 2
    finally:
        if d1 is not None:
            try:d1.Close(False)
            except Exception as exc: print(f"Word property or cleanup warning: {exc}", file=__import__("sys").stderr)
        if d2 is not None:
            try:d2.Close(False)
            except Exception as exc: print(f"Word property or cleanup warning: {exc}", file=__import__("sys").stderr)
        if app is not None:
            try:app.Quit()
            except Exception as exc: print(f"Word property or cleanup warning: {exc}", file=__import__("sys").stderr)
if __name__=='__main__': raise SystemExit(main())
