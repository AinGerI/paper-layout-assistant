from pathlib import Path
import sys,json,subprocess,zipfile,copy
from xml.etree import ElementTree as ET
import yaml
SKILL=Path(r'C:\path\to\word-format')
sys.path.insert(0,str(SKILL/'scripts'))
from word_common import connect_word,document_snapshot,load_profile,detect_role
from format_word import apply_font
ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'04 排版结果'/'reference-paper版式测试'
REF=ROOT/'01 官方模板'/'reference-paper.docx'
SOURCE=ROOT/'04 排版结果'/'文字规范测试'/'学术排版测试样稿_formatted.docx'
OUTPUT=OUT/'学术排版测试样稿_reference-paper版式.docx'
PROFILE=ROOT/'03 格式要求'/'reference-paper_参考版式.yaml'
DATA=json.loads((OUT/'参考版式取样.json').read_text(encoding='utf-8'))

def font(f):
    result={}
    for key,prop in [('east_asia','NameFarEast'),('latin','NameAscii'),('size_pt','Size')]:
        v=f.get(prop)
        if v not in (None,'',9999999):result[key]=v
    for key,prop in [('bold','Bold'),('italic','Italic'),('underline','Underline')]:
        if f.get(prop) not in (None,9999999):result[key]=bool(f[prop])
    if result.get('latin')=='隶书':result['latin']=f.get('NameOther') or '宋体'
    return result

def apply_reference_font(r,raw):
    apply_font(r,font(raw))
    if raw.get('NameAscii')=='隶书':
        r.Font.NameAscii='隶书'
        r.Font.NameOther=raw.get('NameOther') or '宋体'

def paragraph(p):
    result={'alignment':{0:'left',1:'center',2:'right',3:'justify',4:'distributed'}[p['Alignment']]}
    for key,prop in [('space_before_pt','SpaceBefore'),('space_after_pt','SpaceAfter'),('left_indent_pt','LeftIndent'),('right_indent_pt','RightIndent'),('first_line_indent_chars','CharacterUnitFirstLineIndent')]:result[key]=p[prop]
    for key,prop in [('keep_with_next','KeepWithNext'),('keep_together','KeepTogether'),('widow_control','WidowControl'),('page_break_before','PageBreakBefore')]:result[key]=bool(p[prop])
    rule=p['LineSpacingRule'];v=p['LineSpacing']
    result['line_spacing']={0:'single',1:'1.5',2:'double'}.get(rule)
    if rule in (3,4):result['line_spacing']={'rule':'at_least' if rule==3 else 'exact','value':v}
    if rule==5:result['line_spacing']={'rule':'multiple','value':v/12}
    return result

def build_profile():
    samples=DATA['samples'];roles={}
    for name in ('title','abstract_heading','keywords','heading_1','heading_2','heading_3','body','figure_caption','table_caption','figure_paragraph'):
        s=samples[name]
        if name=='title':s=s['title_content']
        elif name=='keywords':s=s['content']
        roles[name]={'paragraph':paragraph(s['paragraph'])}
        if s.get('font'):roles[name]['font']=font(s['font'])
    roles['equation_block']={'paragraph':{'alignment':'center','space_before_pt':6,'space_after_pt':6,'line_spacing':{'rule':'at_least','value':18},'first_line_indent_chars':0,'left_indent_pt':0,'right_indent_pt':0}}
    roles['references_heading']=copy.deepcopy(roles['heading_1'])
    p=DATA['page'];page={'enabled':True,'orientation':'portrait'}
    for key,prop in [('paper_width_cm','PageWidth'),('paper_height_cm','PageHeight'),('top_margin_cm','TopMargin'),('bottom_margin_cm','BottomMargin'),('left_margin_cm','LeftMargin'),('right_margin_cm','RightMargin'),('gutter_cm','Gutter'),('header_distance_cm','HeaderDistance'),('footer_distance_cm','FooterDistance')]:page[key]=p[prop]*2.54/72
    table=DATA['tables'][0]
    tf=font(table['first_cell']['font']);tf['bold']=False
    profile={'schema_version':1,'name':'reference-paper 实际版式参考（非官方规范声明）','source_template':str(REF),
             'extraction_notes':['从代表性段落实际格式提取；不以 Normal 样式替代正文取样。','封面、学校/队员信息、目录与页码不复制。','公式保留原生 OMath；参考文件多数公式为 MathType/OLE，不能照搬对象类型。公式采用已有安全最小行距并沿用中心/右制表位。','表格格式按正文第3张参数表取样（其他宽表有不同字号与行距）。','图片保持原尺寸，超出新正文宽度时才缩小。'],
             'page':page,'roles':roles,
             'table':{'enabled':True,'alignment':'center','font':tf,'paragraph':paragraph(table['first_cell']['paragraph']),
                      'cell':{'vertical_alignment':'top',**{key:table['padding'][prop] for key,prop in [('margin_top_pt','TopPadding'),('margin_bottom_pt','BottomPadding'),('margin_left_pt','LeftPadding'),('margin_right_pt','RightPadding')]}},
                      'border':{'type':'three_line','top_pt':1.5,'bottom_pt':1.5,'header_bottom_pt':table['header_border_width']/8}},
             'figure':{'enabled':True,'paragraph_alignment':'center','max_width_cm':(p['PageWidth']-p['LeftMargin']-p['RightMargin'])*2.54/72,'lock_aspect_ratio':True,'resize_floating_shapes':False},
             'detection':{'regex_overrides':[{'pattern':'^本测试的行内公式','role':'body'}]}}
    PROFILE.write_text(yaml.safe_dump(profile,allow_unicode=True,sort_keys=False),encoding='utf-8')
    load_profile(PROFILE)
    return profile

def run(name,*args):
    r=subprocess.run([sys.executable,str(SKILL/'scripts'/name),*map(str,args)],capture_output=True,text=True,encoding='utf-8',timeout=240)
    (OUT/(name+'.log')).write_text(r.stdout+'\n'+r.stderr,encoding='utf-8')
    print(name,r.returncode,flush=True)
    if r.returncode:raise RuntimeError(r.stderr or r.stdout)

def role_of(p,i,doc,c,profile):return detect_role(p,i,doc,c,profile)

def supplement(profile):
    app,c=connect_word();d=None
    try:
        d=app.Documents.Open(str(OUTPUT),ReadOnly=False,AddToRecentFiles=False)
        before=document_snapshot(d)
        for sec in d.Sections:sec.PageSetup.VerticalAlignment=DATA['page']['VerticalAlignment']
        for i in range(1,d.Paragraphs.Count+1):
            p=d.Paragraphs(i);r=role_of(p,i,d,c,profile);text=str(p.Range.Text)
            spec=profile['roles'].get(r)
            if r=='table_text':spec=profile['table']
            if spec:
                target=spec.get('paragraph',{})
                if target.get('first_line_indent_chars')==0:
                    p.Format.FirstLineIndent=0;p.Format.CharacterUnitFirstLineIndent=0
            source=DATA['samples'].get(r)
            if r=='references_heading':source=DATA['samples']['heading_1']
            if source:p.Format.DisableLineHeightGrid=bool(source['paragraph']['DisableLineHeightGrid'])
            if r=='table_text':p.Format.DisableLineHeightGrid=bool(DATA['tables'][0]['paragraph']['DisableLineHeightGrid'])
            if r=='keywords':
                label=p.Range.Duplicate;label.End=label.Start+4
                apply_reference_font(label,DATA['samples']['keywords']['first_character']['font'])
            if r=='abstract_heading':apply_reference_font(p.Range,DATA['samples']['abstract_heading']['font'])
            if r=='equation_block':
                p.Format.DisableLineHeightGrid=True
                p.Format.Alignment=0
                w=d.PageSetup.PageWidth-d.PageSetup.LeftMargin-d.PageSetup.RightMargin
                p.Format.TabStops.ClearAll();p.Format.TabStops.Add(w/2,1,0);p.Format.TabStops.Add(w,2,0)
        for table in d.Tables:apply_font(table.Rows(1).Range,font(DATA['tables'][0]['first_cell']['font']))
        assert document_snapshot(d)==before
        d.Save()
    finally:
        if d is not None:d.Close(False)
        app.Quit(SaveChanges=0)
    # Preserve exact reference document-grid semantics; supported by OOXML but not Profile v1.
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with zipfile.ZipFile(REF) as z:
        refxml=ET.fromstring(z.read('word/document.xml'))
        grid=refxml.find('.//w:sectPr/w:docGrid',ns)
    with zipfile.ZipFile(OUTPUT) as z:entries=[(x,z.read(x.filename)) for x in z.infolist()]
    xml=ET.fromstring(next(b for i,b in entries if i.filename=='word/document.xml'))
    for sec in xml.findall('.//w:sectPr',ns):
        old=sec.find('w:docGrid',ns)
        if old is not None:sec.remove(old)
        sec.append(copy.deepcopy(grid))
    # lxml preserves original namespace prefixes and ignorable-extension bindings.
    # The system stdlib ElementTree serialization would lose those declarations.
    import subprocess
    helper=OUT/'copy_grid.py'
    helper.write_text('''from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import copy,sys
ref,out=map(Path,sys.argv[1:])
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
with ZipFile(ref) as z: grid=E.fromstring(z.read('word/document.xml')).find('.//w:sectPr/w:docGrid',ns)
with ZipFile(out) as z: entries=[(i,z.read(i.filename)) for i in z.infolist()]
root=E.fromstring(next(b for i,b in entries if i.filename=='word/document.xml'))
for sec in root.findall('.//w:sectPr',ns):
 old=sec.find('w:docGrid',ns)
 if old is not None: sec.remove(old)
 sec.append(copy.deepcopy(grid))
with ZipFile(out,'w') as z:
 for i,b in entries: z.writestr(i,E.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True) if i.filename=='word/document.xml' else b)
''',encoding='utf-8')
    subprocess.run([r'C:\path\to\python.exe',str(helper),str(REF),str(OUTPUT)],check=True,timeout=30)

def verify(profile,source_bytes,ref_bytes):
    checks=[]
    def ck(name,ok):checks.append({'name':name,'passed':bool(ok)})
    def near(a,b):return abs(float(a)-float(b))<0.2
    app,c=connect_word();d=None;snaps=[]
    try:
        for path in (SOURCE,OUTPUT):
            d=app.Documents.Open(str(path),ReadOnly=True,AddToRecentFiles=False)
            snaps.append(document_snapshot(d))
            if path==OUTPUT:
                for key,val in DATA['page'].items():ck('Page '+key,near(getattr(d.PageSetup,key),val))
                for i in range(1,d.Paragraphs.Count+1):
                    p=d.Paragraphs(i);role=role_of(p,i,d,c,profile);spec=profile['roles'].get(role)
                    if not spec:continue
                    ps=spec['paragraph'];fmt=p.Format
                    ck(f'P{i} {role} spacing',near(fmt.SpaceBefore,ps['space_before_pt']) and near(fmt.SpaceAfter,ps['space_after_pt']))
                    ck(f'P{i} {role} indentation',near(fmt.CharacterUnitFirstLineIndent,ps['first_line_indent_chars']) and (ps['first_line_indent_chars']!=0 or near(fmt.FirstLineIndent,0)))
                    if 'font' in spec:
                        f=spec['font'];r=p.Range.Duplicate;r.End=r.Start+1
                        if role=='keywords':f=font(DATA['samples']['keywords']['first_character']['font'])
                        if role in ('keywords','abstract_heading'):f['latin']='隶书'
                        ck(f'P{i} {role} font',r.Font.NameFarEast==f['east_asia'] and r.Font.NameAscii==f['latin'] and near(r.Font.Size,f['size_pt']) and bool(r.Font.Bold)==f['bold'] and bool(r.Font.Underline)==f['underline'])
                    spacing=ps['line_spacing'];rule=fmt.LineSpacingRule
                    if isinstance(spacing,dict):ok=rule=={'exact':4,'at_least':3,'multiple':5}[spacing['rule']] and near(fmt.LineSpacing,spacing['value']*(12 if spacing['rule']=='multiple' else 1))
                    else:ok=rule=={'single':0,'1.5':1,'double':2}[spacing]
                    ck(f'P{i} {role} line spacing',ok)
                    desired={'left':0,'center':1,'right':2,'justify':3}[ps['alignment']]
                    if role=='equation_block':desired=0
                    ck(f'P{i} {role} alignment',fmt.Alignment==desired)
                    if role=='keywords':
                        r=p.Range.Duplicate;r.Start+=4;r.End-=1
                        ck('Keyword content font',r.Font.NameFarEast=='宋体' and r.Font.Size==12 and bool(r.Font.Bold))
                tb=d.Tables(1)
                ck('Three-line table',tb.Borders(-1).LineStyle==1 and tb.Borders(-3).LineStyle==1 and tb.Borders(-2).LineStyle==0 and tb.Borders(-4).LineStyle==0 and tb.Rows(1).Borders(-3).LineStyle==1)
                ck('Table body font and alignment',tb.Cell(2,1).Range.Font.Size==10 and tb.Cell(2,1).Range.Font.NameFarEast=='宋体' and tb.Rows.Alignment==1 and tb.Cell(2,1).Range.ParagraphFormat.Alignment==1)
                ck('Table header bold',bool(tb.Cell(1,1).Range.Font.Bold))
                d.Repaginate();d.ExportAsFixedFormat(str(OUT/'final-preview.pdf'),17)
            d.Close(False);d=None
    finally:
        if d is not None:d.Close(False)
        app.Quit(SaveChanges=0)
    ck('Source unchanged',SOURCE.read_bytes()==source_bytes)
    ck('Reference unchanged',REF.read_bytes()==ref_bytes)
    ck('Exact text and object counts',snaps[0]==snaps[1])
    # Reuse the existing independent OOXML integrity comparator, without running its workflow.
    import importlib.util
    sp=importlib.util.spec_from_file_location('previous_test',Path(__file__).with_name('run_spec_test.py'))
    module=importlib.util.module_from_spec(sp);sp.loader.exec_module(module)
    ck('Math text table data and media bytes',module.xml_content(SOURCE)==module.xml_content(OUTPUT))
    report={'passed':all(x['passed'] for x in checks),'checks':checks,'counts':{k:v for k,v in snaps[1].items() if k!='body_text'},'source':str(SOURCE),'reference':str(REF),'output':str(OUTPUT),'limits':profile['extraction_notes']}
    (OUT/'最终检查.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    (OUT/'最终检查.md').write_text('# 参考论文版式重排检查\n\n'+f"自动检查：{sum(x['passed'] for x in checks)}/{len(checks)} 项通过。\n\n"+'输入是已排版论文，不是独立文字规范；版式根据实际取样整理，不能视为官方格式标准。\n\n'+ '\n'.join('- '+s for s in profile['extraction_notes'])+'\n\n'+ '\n'.join(f"- {'PASS' if x['passed'] else 'FAIL'} {x['name']}" for x in checks),encoding='utf-8')
    print(json.dumps({'passed':report['passed'],'checks':len(checks),'failures':[x for x in checks if not x['passed']]},ensure_ascii=False))
    if not report['passed']:raise SystemExit(2)

if __name__=='__main__':
    sb=SOURCE.read_bytes();rb=REF.read_bytes();profile=build_profile()
    if '--resume' not in sys.argv and '--format-only' not in sys.argv:
        run('extract_profile.py',REF,'--output',OUT/'内置样式初步提取.yaml')
        run('inspect_word.py',SOURCE,'--profile',PROFILE,'--json',OUT/'样稿分析.json','--md',OUT/'样稿分析.md')
    if '--resume' not in sys.argv:
        run('format_word.py',SOURCE,'--profile',PROFILE,'--output',OUTPUT)
    supplement(profile)
    run('validate_word.py',SOURCE,OUTPUT)
    verify(profile,sb,rb)
