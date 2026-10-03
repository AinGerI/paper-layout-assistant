from __future__ import annotations
import argparse
from pathlib import Path
from word_common import connect_word, dump_yaml, resolve_path, safe_float


def alignment_name(v,c):
    m={int(c.wdAlignParagraphLeft):'left',int(c.wdAlignParagraphCenter):'center',int(c.wdAlignParagraphRight):'right',int(c.wdAlignParagraphJustify):'justify'}
    try:return m.get(int(v))
    except:return None

def style_spec(style,c):
    f=style.Font; p=style.ParagraphFormat
    spec={'font':{},'paragraph':{}}
    for key,attr in [('east_asia','NameFarEast'),('latin','NameAscii'),('size_pt','Size')]:
        try:
            val=getattr(f,attr)
            if key=='size_pt': val=safe_float(val)
            if val not in (None,''): spec['font'][key]=val
        except:pass
    for key,attr in [('space_before_pt','SpaceBefore'),('space_after_pt','SpaceAfter'),('first_line_indent_chars','CharacterUnitFirstLineIndent')]:
        try:
            val=safe_float(getattr(p,attr))
            if val is not None: spec['paragraph'][key]=round(val,3)
        except:pass
    try:
        a=alignment_name(p.Alignment,c)
        if a: spec['paragraph']['alignment']=a
    except:pass
    try:
        rule=int(p.LineSpacingRule)
        if rule==int(c.wdLineSpaceSingle): spec['paragraph']['line_spacing']='single'
        elif rule==int(c.wdLineSpace1pt5): spec['paragraph']['line_spacing']='1.5'
        elif rule==int(c.wdLineSpaceDouble): spec['paragraph']['line_spacing']='double'
        elif rule==int(c.wdLineSpaceExactly): spec['paragraph']['line_spacing']={'rule':'exact','value':safe_float(p.LineSpacing)}
        elif rule==int(c.wdLineSpaceAtLeast): spec['paragraph']['line_spacing']={'rule':'at_least','value':safe_float(p.LineSpacing)}
        elif rule==int(c.wdLineSpaceMultiple): spec['paragraph']['line_spacing']={'rule':'multiple','value':None}
    except:pass
    return spec

def main():
    ap=argparse.ArgumentParser(description='Extract a starter profile from a Word template/document styles.')
    ap.add_argument('template'); ap.add_argument('--output'); ap.add_argument('--visible',action='store_true'); a=ap.parse_args()
    path=resolve_path(a.template); out=resolve_path(a.output or str(Path(path).with_suffix(''))+'_profile.yaml')
    app=None;doc=None
    try:
        app,c=connect_word(a.visible); doc=app.Documents.Open(path,ReadOnly=True,AddToRecentFiles=False)
        sec=doc.Sections(1).PageSetup
        prof={'schema_version':1,'name':Path(path).stem+' extracted profile','source_template':path,
              'document':{'repaginate':True},
              'page':{'enabled':True,'top_margin_cm':round(app.PointsToCentimeters(sec.TopMargin),3),'bottom_margin_cm':round(app.PointsToCentimeters(sec.BottomMargin),3),'left_margin_cm':round(app.PointsToCentimeters(sec.LeftMargin),3),'right_margin_cm':round(app.PointsToCentimeters(sec.RightMargin),3)},
              'roles':{},'table':{'enabled':False},'figure':{'enabled':False},
              'detection':{},
              'extraction_notes':['标题/正文/题注优先从 Word 内置样式提取。表格和图片布局默认不自动推断，请结合官方格式要求补充。']}
        pairs=[('body','wdStyleNormal'),('title','wdStyleTitle'),('heading_1','wdStyleHeading1'),('heading_2','wdStyleHeading2'),('heading_3','wdStyleHeading3'),('figure_caption','wdStyleCaption'),('table_caption','wdStyleCaption')]
        for role,constname in pairs:
            try: prof['roles'][role]=style_spec(doc.Styles(getattr(c,constname)),c)
            except Exception: pass
        dump_yaml(prof,out); print(out); return 0
    finally:
        if doc is not None:
            try:doc.Close(False)
            except:pass
        if app is not None:
            try:app.Quit()
            except:pass
if __name__=='__main__': raise SystemExit(main())
