from pathlib import Path
import sys,json
sys.path.insert(0,r'C:\path\to\word-format\scripts')
from word_common import connect_word,style_name

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'04 排版结果'/'reference-paper版式测试'
OUT.mkdir(exist_ok=True)
REF=ROOT/'01 官方模板'/'reference-paper.docx'

def snapshot(r):
    p=r.ParagraphFormat;f=r.Font
    def value(obj,attr):
        try:return getattr(obj,attr)
        except Exception:return None
    return {'font':{k:value(f,k) for k in ['NameFarEast','NameAscii','NameOther','Size','Bold','Italic','Color','Underline','Scaling']},
            'paragraph':{k:value(p,k) for k in ['Alignment','SpaceBefore','SpaceAfter','LineSpacingRule','LineSpacing','FirstLineIndent','CharacterUnitFirstLineIndent','LeftIndent','RightIndent','KeepWithNext','KeepTogether','WidowControl','PageBreakBefore','DisableLineHeightGrid']}}

app,c=connect_word();doc=None
try:
    doc=app.Documents.Open(str(REF),ReadOnly=True,AddToRecentFiles=False)
    ps=doc.Sections(1).PageSetup
    data={'page':{k:getattr(ps,k) for k in ['PageWidth','PageHeight','TopMargin','BottomMargin','LeftMargin','RightMargin','Gutter','HeaderDistance','FooterDistance','Orientation','VerticalAlignment']},'samples':{},'tables':[]}
    for i in range(1, min(doc.Paragraphs.Count,650)+1):
        p=doc.Paragraphs(i);t=str(p.Range.Text).strip();st=style_name(p)
        if 'TOC' in st.upper() or '目录' in st:continue
        role=None
        if t.startswith('题 目：'):role='title'
        elif t.startswith('摘 要：'):role='abstract_heading'
        elif t.startswith('氢燃料电池能量转化效率高'):role='abstract_body'
        elif t.startswith('关键词：'):role='keywords'
        elif st in ('标题 1','标题1','Heading 1'):role='heading_1'
        elif st in ('标题 2','标题2','Heading 2'):role='heading_2'
        elif st in ('标题 3','标题3','Heading 3'):role='heading_3'
        elif t.startswith('质子交换膜燃料电池（'):role='body'
        elif t.startswith('图 ') and p.Range.InlineShapes.Count==0:role='figure_caption'
        elif t.startswith('表 ') and p.Range.Tables.Count==0:role='table_caption'
        elif p.Range.InlineShapes.Count and role is None:role='figure_paragraph'
        if role and role not in data['samples']:
            r=p.Range.Duplicate
            if r.End>r.Start:r.End-=1
            rec={'index':i,'text':t[:100],'style':st,**snapshot(r)}
            rec['first_character']=snapshot(doc.Range(r.Start,min(r.Start+1,r.End)))
            if role=='title':rec['title_content']=snapshot(doc.Range(r.Start+9,r.End))
            if role=='keywords':rec['content']=snapshot(doc.Range(r.Start+4,r.End))
            data['samples'][role]=rec
        if all(k in data['samples'] for k in ['title','abstract_heading','abstract_body','keywords','heading_1','heading_2','heading_3','body','figure_caption','table_caption','figure_paragraph']):break
    if doc.OMaths.Count:
        r=doc.OMaths(1).Range
        data['equation']=snapshot(r.Paragraphs(1).Range)
    for i in range(1,min(doc.Tables.Count,5)+1):
        tb=doc.Tables(i)
        rec={'index':i,'text':str(tb.Range.Text)[:100],**snapshot(tb.Range)}
        try:
            rec['alignment']=tb.Rows.Alignment
            rec['vertical_alignment']=tb.Cell(1,1).VerticalAlignment
            rec['borders']={str(k):{'style':tb.Borders(k).LineStyle,'width':tb.Borders(k).LineWidth} for k in (-1,-2,-3,-4,-5,-6)}
            rec['first_cell']=snapshot(tb.Cell(1,1).Range)
            rec['padding']={k:getattr(tb,k) for k in ['TopPadding','BottomPadding','LeftPadding','RightPadding']}
        except Exception as exc:rec['warning']=str(exc)
        data['tables'].append(rec)
    (OUT/'参考版式取样.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(data,ensure_ascii=False,indent=2))
finally:
    if doc is not None:doc.Close(False)
    app.Quit(SaveChanges=0)
