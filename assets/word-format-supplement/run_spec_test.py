from pathlib import Path
import json
import subprocess
import sys
import struct
import zipfile
import xml.etree.ElementTree as ET
import yaml

SKILL = Path(r'C:\path\to\word-format')
sys.path.insert(0, str(SKILL/'scripts'))
from word_common import connect_word, document_snapshot, load_profile
from format_word import apply_font

ROOT = Path(__file__).resolve().parents[1]
RULES = ROOT/'03 格式要求'
RESULT = ROOT/'04 排版结果'/'文字规范测试'
SOURCE = ROOT/'02 待排版论文'/'学术排版测试样稿_原稿.docx'
OUTPUT = RESULT/'学术排版测试样稿_formatted.docx'
PROFILE = RULES/'学术论文测试规范.yaml'

def para(align='left', before=0, after=0, spacing=1.5, indent=0, **extra):
    return dict(alignment=align, space_before_pt=before, space_after_pt=after,
                line_spacing=spacing, first_line_indent_chars=indent,
                left_indent_pt=0,right_indent_pt=0,**extra)

def role(size=12, bold=False, chinese='宋体', **paragraph):
    return {'font':dict(east_asia=chinese,latin='Times New Roman',size_pt=size,bold=bold,underline=False),
            'paragraph':paragraph}

def write_profiles():
    profile = {
        'schema_version':1,'name':'Word 学术论文自动排版测试规范（非正式标准）',
        'source_requirements':'本聊天用户提供的第1至12节；图片末句不完整，不推断。补充规则见同目录 supplemental.yaml。',
        'page':dict(enabled=True,orientation='portrait',paper_width_cm=21,paper_height_cm=29.7,
                    top_margin_cm=2.54,bottom_margin_cm=2.54,left_margin_cm=2.8,right_margin_cm=2.8,gutter_cm=0),
        'roles':{
            'title':role(18,True,'黑体',**para('center',0,18,'single',keep_with_next=True)),
            'heading_1':role(16,True,'黑体',**para(before=18,after=12,keep_with_next=True,keep_together=True)),
            'heading_2':role(14,True,'黑体',**para(before=12,after=6,keep_with_next=True)),
            'heading_3':role(12,True,'黑体',**para(before=6,after=3,keep_with_next=True)),
            'body':role(**para('justify',indent=2,widow_control=True)),
            'abstract_heading':role(14,True,'黑体',**para('center',12,6,'single')),
            'keywords':role(**para(before=6,after=12)),
            'equation_block':{'paragraph':para('center',6,6,{'rule':'at_least','value':18})},
            'figure_paragraph':{'paragraph':para('center',6,0,'single')},
        },
        'table':{'enabled':False},
        'figure':dict(enabled=True,paragraph_alignment='center',max_width_cm=14,lock_aspect_ratio=True,resize_floating_shapes=False),
        'detection':{'regex_overrides':[
            {'pattern':'^本测试的行内公式','role':'body'},
            {'pattern':'^\\s*（[12]）$','role':'equation_block'}]},
    }
    PROFILE.write_text(yaml.safe_dump(profile,allow_unicode=True,sort_keys=False),encoding='utf-8')
    supplemental = {
        'page_vertical_alignment':'top',
        'global_text':{'east_asia':'宋体','latin':'Times New Roman','color_rgb':'000000','underline':False,'shadow':False,'scaling_percent':100},
        'keyword_label':{'text':'关键词：','east_asia':'黑体','size_pt':12,'bold':True},
        'keyword_content':{'east_asia':'宋体','latin':'Times New Roman','size_pt':12,'bold':False},
        'math':{'font':'preserve native OMath font','size_pt':12,'display_tabs':'center and right within text width; existing tabs only'},
        'image_width':'min(14 cm, available text width)',
        'blank_paragraphs':'do not insert or delete',
        'inline_math_spacing':'Word multiple 1.5 expands with tall native math; visually verify',
    }
    (RULES/'supplemental.yaml').write_text(yaml.safe_dump(supplemental,allow_unicode=True,sort_keys=False),encoding='utf-8')
    load_profile(PROFILE)

def make_bitmap(path):
    width,height=900,150
    rows=[]
    for y in range(height):
        row=bytearray()
        for x in range(width):
            rgb=(246,248,250)
            if 35<y<115 and any(a<x<b for a,b in ((40,230),(350,540),(660,850))): rgb=(42,85,120)
            if 72<y<78 and (230<=x<=350 or 540<=x<=660): rgb=(85,120,145)
            row.extend(rgb[::-1])
        rows.append(bytes(row))
    pixels=b''.join(rows)
    path.write_bytes(b'BM'+struct.pack('<IHHI',54+len(pixels),0,0,54)+struct.pack('<IiiHHIIiiII',40,width,height,1,24,0,len(pixels),0,0,0,0)+pixels)

def create_source():
    if SOURCE.exists():
        return
    app,c=connect_word();doc=None
    try:
        doc=app.Documents.Add()
        lines=[
            '基于多目标优化的无人机运输与通信协同调度研究',
            '摘要',
            '本文为自动排版功能测试样稿，不代表真实研究结论。样稿同时包含中文、English words、数字 2026、标题、表格、图片和原生公式，用于检查字体设置与内容保护。测试仅比较格式变化，不修改任何文字、数据或编号。',
            '关键词：无人机；多目标优化；UAV scheduling；协同通信',
            '1 问题重述',
            '本测试模拟无人机运输与通信协同调度问题。系统包含起点、服务节点与终点，需在给定资源约束下安排服务顺序。这里的参数和数据仅用于展示排版效果，不应解释为实验结果。',
            '1.1 问题背景',
            '在测试场景中，运输任务与通信任务共享有限的时间和能量资源。中文与英文需要分别设置字体，例如 UAV、latency 和 energy；数值 12.50 与 2026 应保持原样。不同自然段之间不添加空白段落。',
            '1.1.1 无人机运输需求分析',
            '本测试的行内公式 x^2+y^2=z^2 嵌入正文中，前后的文字用于判断公式是否仍与正文连续排列。公式必须保留为 Word 原生数学对象，不能替换成图片或普通文本。',
            '本测试的行内公式 (a+b)/(c+d) 含有分子和分母，用于观察较高公式在一倍半行距下的显示情况。需要检查上下边界是否完整，且中文、English 与数字 123 的字体均正确。',
            '2 模型与测试结果',
            '以下独立公式用于验证公式居中、编号右对齐和最小行距。源文件已包含制表符与编号，排版过程只能调整制表位和段落属性，不能改变编号或公式内容。',
            '\tx^2+y^2=z^2\t（1）',
            '\t(a+b)/(c+d)\t（2）',
            '2.1 测试图表',
            '下图只是一幅测试示意图，不表达真实研究结论。它在原稿中的宽度故意超过目标上限，用于验证嵌入型图片能否等比例缩小到不超过 14 cm，并且保持图片内容不变。',
            '',
            '图 1 排版测试示意图',
            '表 1 示例参数（仅用于排版测试）',
            '',
            '表中数据必须保持原样。表格的边框与对齐方式在本次规范中没有单独要求，因此保留原稿设置，仅遵守明确的全文字体规则。',
            '3 内容保护检查',
            '排版前后对正文文字、表格数据、图片内容、原生公式、浮动对象和域进行比较。该样稿用于检查排版引擎的行为，不能据此认为任意复杂论文都已通过全部测试。',
            '参考文献',
            '[1] 排版测试资料条目。此条目不是实际出版物，仅用于验证参考文献文字保持不变。',
        ]
        doc.Content.Text='\r'.join(lines)
        doc.Content.Font.NameFarEast='微软雅黑';doc.Content.Font.NameAscii='Arial'
        doc.Content.Font.Size=11;doc.Content.Font.Color=0x804000
        doc.Content.ParagraphFormat.SpaceAfter=4
        doc.Content.ParagraphFormat.LineSpacingRule=0
        doc.Paragraphs(1).Range.Style=-63
        for i in (5,12,23): doc.Paragraphs(i).Range.Style=-2
        for i in (7,16): doc.Paragraphs(i).Range.Style=-3
        doc.Paragraphs(9).Range.Style=-4
        # Reverse order keeps ranges stable while Word builds the math structures.
        for index,expr in ((15,'(a+b)/(c+d)'),(14,'x^2+y^2=z^2'),(11,'(a+b)/(c+d)'),(10,'x^2+y^2=z^2')):
            p=doc.Paragraphs(index);offset=str(p.Range.Text).index(expr)
            r=doc.Range(p.Range.Start+offset,p.Range.Start+offset+len(expr))
            doc.OMaths.Add(r)
            r.OMaths(1).BuildUp()
        image=RESULT/'test-image.bmp';make_bitmap(image)
        r=doc.Paragraphs(18).Range.Duplicate;r.End-=1
        pic=doc.InlineShapes.AddPicture(str(image),False,True,r)
        pic.LockAspectRatio=True;pic.Width=16.5*72/2.54
        table=doc.Tables.Add(doc.Paragraphs(21).Range,3,2)
        for ri,values in enumerate([['参数','测试值'],['任务数量','12'],['阈值','0.85']],1):
            for ci,text in enumerate(values,1):table.Cell(ri,ci).Range.Text=text
        table.Borders.Enable=True
        doc.SaveAs2(str(SOURCE),FileFormat=16)
    finally:
        if doc is not None:doc.Close(False)
        app.Quit(SaveChanges=0)

def run_script(name,*args):
    result=subprocess.run([sys.executable,str(SKILL/'scripts'/name),*map(str,args)],capture_output=True,text=True,encoding='utf-8',timeout=180)
    (RESULT/(name+'.log')).write_text(result.stdout+'\n'+result.stderr,encoding='utf-8')
    print(name,result.returncode,flush=True)
    if result.returncode:raise RuntimeError(result.stderr or result.stdout)

def supplement():
    app,c=connect_word();doc=None
    try:
        doc=app.Documents.Open(str(OUTPUT),ReadOnly=False,AddToRecentFiles=False)
        before=document_snapshot(doc)
        doc.Content.Font.Color=0
        doc.Content.Font.Underline=0;doc.Content.Font.Shadow=False
        doc.Content.Font.Scaling=100;doc.Content.Font.Outline=False
        doc.Content.Font.Emboss=False;doc.Content.Font.Engrave=False
        # Base engine deliberately preserves OMath fonts, including Cambria Math.
        for i in range(1,doc.OMaths.Count+1):doc.OMaths(i).Range.Font.Size=12
        for section in doc.Sections:section.PageSetup.VerticalAlignment=0
        profile=load_profile(PROFILE)
        from word_common import detect_role
        for i in range(1,doc.Paragraphs.Count+1):
            p=doc.Paragraphs(i);text=str(p.Range.Text)
            role_name=detect_role(p,i,doc,c,profile)
            paragraph_spec=(profile['roles'].get(role_name) or {}).get('paragraph') or {}
            if paragraph_spec.get('first_line_indent_chars') == 0:
                # Chinese character indentation can coexist with a stale point value.
                p.Format.FirstLineIndent=0
                p.Format.CharacterUnitFirstLineIndent=0
            if role_name in ('table_text','figure_caption','table_caption','references_heading') or text.startswith('[1]'):
                apply_font(p.Range,{'east_asia':'宋体','latin':'Times New Roman'})
            p.Format.DisableLineHeightGrid=True
            if text.startswith('关键词：'):
                apply_font(p.Range,dict(east_asia='宋体',latin='Times New Roman',size_pt=12,bold=False))
                label=p.Range.Duplicate;label.End=label.Start+4
                apply_font(label,dict(east_asia='黑体',latin='Times New Roman',size_pt=12,bold=True))
            if p.Range.OMaths.Count and text.startswith('\t'):
                setup=p.Range.Sections(1).PageSetup
                width=setup.PageWidth-setup.LeftMargin-setup.RightMargin-setup.Gutter
                p.Format.Alignment=0
                p.Format.TabStops.ClearAll()
                p.Format.TabStops.Add(float(width/2),1,0)
                p.Format.TabStops.Add(float(width),2,0)
        for i in range(1,doc.InlineShapes.Count+1):
            shape=doc.InlineShapes(i);setup=shape.Range.Sections(1).PageSetup
            limit=min(14*72/2.54,setup.PageWidth-setup.LeftMargin-setup.RightMargin-setup.Gutter)
            if shape.Width>limit:shape.Width=limit
        assert document_snapshot(doc)==before,'Supplement changed content'
        doc.Save()
    finally:
        if doc is not None:doc.Close(False)
        app.Quit(SaveChanges=0)

def xml_content(path):
    ns={'m':'http://schemas.openxmlformats.org/officeDocument/2006/math','w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with zipfile.ZipFile(path) as z:
        xml=ET.fromstring(z.read('word/document.xml'))
        return {
            'math_text':[''.join(el.itertext()) for el in xml.findall('.//m:oMath',ns)],
            'tables':[[t.text for t in table.findall('.//w:t',ns)] for table in xml.findall('.//w:tbl',ns)],
            'media':sorted([z.read(n) for n in z.namelist() if n.startswith('word/media/')]),
        }

def verify_and_render(source_bytes):
    checks=[]
    def check(name,passed,detail=None):
        checks.append({'name':name,'passed':bool(passed),'detail':detail})
    def near(a,b):return abs(float(a)-b)<0.15
    app,c=connect_word();doc=None;snapshots=[];properties=[]
    try:
        for path in (SOURCE,OUTPUT):
            doc=app.Documents.Open(str(path),ReadOnly=True,AddToRecentFiles=False)
            snapshots.append(document_snapshot(doc))
            if path==OUTPUT:
                s=doc.Sections(1).PageSetup
                check('A4 and portrait',near(s.PageWidth,21*72/2.54) and near(s.PageHeight,29.7*72/2.54) and s.Orientation==0)
                check('Margins gutter vertical alignment',all(near(getattr(s,k),v*72/2.54) for k,v in dict(TopMargin=2.54,BottomMargin=2.54,LeftMargin=2.8,RightMargin=2.8,Gutter=0).items()) and s.VerticalAlignment==0)
                profile=load_profile(PROFILE)
                from word_common import detect_role
                for i in range(1,doc.Paragraphs.Count+1):
                    p=doc.Paragraphs(i);text=str(p.Range.Text);name=detect_role(p,i,doc,c,profile)
                    spec=profile['roles'].get(name)
                    if not spec:continue
                    pf=p.Format;target=spec['paragraph']
                    check(f'P{i} {name} spacing',near(pf.SpaceBefore,target['space_before_pt']) and near(pf.SpaceAfter,target['space_after_pt']))
                    check(f'P{i} {name} indentation',near(pf.CharacterUnitFirstLineIndent,target['first_line_indent_chars']) and near(pf.LeftIndent,0) and near(pf.RightIndent,0))
                    if target['first_line_indent_chars']==0:
                        check(f'P{i} zero physical first-line indent',near(pf.FirstLineIndent,0))
                    expected_alignment={'left':0,'center':1,'justify':3}[target['alignment']]
                    if name=='equation_block':expected_alignment=0
                    check(f'P{i} {name} alignment',pf.Alignment==expected_alignment)
                    for key,attr in [('keep_with_next','KeepWithNext'),('keep_together','KeepTogether'),('widow_control','WidowControl')]:
                        if key in target:check(f'P{i} {key}',bool(getattr(pf,attr))==target[key])
                    spacing=target['line_spacing']
                    if spacing=='single':ok=pf.LineSpacingRule==0
                    elif isinstance(spacing,dict):ok=pf.LineSpacingRule==3 and near(pf.LineSpacing,18)
                    else:ok=pf.LineSpacingRule in (1,5) and near(pf.LineSpacing,18)
                    check(f'P{i} {name} line spacing',ok)
                    if 'font' in spec:
                        r=p.Range.Duplicate;r.End=r.Start+1
                        fs=spec['font']
                        if name=='keywords':fs=dict(fs,east_asia='黑体',bold=True)
                        check(f'P{i} {name} font',r.Font.NameFarEast==fs['east_asia'] and r.Font.NameAscii==fs['latin'] and near(r.Font.Size,fs['size_pt']) and bool(r.Font.Bold)==fs['bold'])
                    if name=='keywords':
                        r=p.Range.Duplicate;r.Start+=4;r.End-=1
                        check('Keyword content regular Song/TNR',r.Font.NameFarEast=='宋体' and r.Font.NameAscii=='Times New Roman' and r.Font.Bold==0)
                    if name=='equation_block':
                        tabs=pf.TabStops;width=s.PageWidth-s.LeftMargin-s.RightMargin
                        check(f'P{i} equation tabs',tabs.Count==2 and tabs(1).Alignment==1 and tabs(2).Alignment==2 and near(tabs(1).Position,width/2) and near(tabs(2).Position,width))
                    properties.append({'index':i,'role':name,'text':text.strip(),'spacing':pf.LineSpacingRule,'line_pt':pf.LineSpacing})
                check('Global black/no underline/no shadow/no scaling',doc.Content.Font.Color==0 and doc.Content.Font.Underline==0 and doc.Content.Font.Shadow==0 and doc.Content.Font.Scaling==100)
                check('Native math count 4',doc.OMaths.Count==4)
                check('Native math font and size',all(doc.OMaths(i).Range.Font.Name=='Cambria Math' and near(doc.OMaths(i).Range.Font.Size,12) for i in range(1,doc.OMaths.Count+1)))
                pic=doc.InlineShapes(1)
                check('Image width and aspect',pic.Width<=14*72/2.54+0.15 and abs(pic.Width/pic.Height-6)<0.02 and bool(pic.LockAspectRatio))
                check('No added blank paragraphs',not any(not str(doc.Paragraphs(i).Range.Text).strip() and doc.Paragraphs(i).Range.InlineShapes.Count==0 for i in range(1,doc.Paragraphs.Count+1)))
            doc.Repaginate()
            pdf=RESULT/('after.pdf' if path==OUTPUT else 'before.pdf')
            doc.ExportAsFixedFormat(str(pdf),17)
            doc.Close(False);doc=None
    finally:
        if doc is not None:doc.Close(False)
        app.Quit(SaveChanges=0)
    check('Original file byte equality',SOURCE.read_bytes()==source_bytes)
    check('Text and object count equality',snapshots[0]==snapshots[1])
    check('OOXML math text tables and image bytes equality',xml_content(SOURCE)==xml_content(OUTPUT))
    report={'passed':all(x['passed'] for x in checks),'checks':checks,'paragraphs':properties,
            'counts':{k:v for k,v in snapshots[1].items() if k!='body_text'},
            'base_skill_only':False,'supplemental_step':'run_spec_test.py supplement(); see supplemental.yaml',
            'limits':['Synthetic sample, not a user thesis','Final image requirement is incomplete','Existing tabs used; arbitrary formula-number conversion not tested','No punctuation or blank-paragraph deletion performed']}
    (RESULT/'验收报告.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    (RESULT/'验收报告.md').write_text('# 文字规范排版测试\n\n'+f"自动检查：{'通过' if report['passed'] else '存在失败'}，{sum(x['passed'] for x in checks)}/{len(checks)} 项。\n\n"+
        '使用自动生成的测试样稿；原稿未覆盖。原生公式、正文、表格数据与图片内容逐项比较。\n\n'+
        '基础 Skill 处理页面、段落、字体和图片尺寸；任务内补充步骤处理关键词局部字体、黑色字体、顶端对齐和公式制表位。这不等于基础 Skill 已原生支持所有条款。\n\n'+
        '图片规范最后一句不完整，本次未推断。关键词分隔符和公式制表符在测试原稿中已存在，排版时未修改。\n\n'+
        '\n'.join(f"- {'PASS' if x['passed'] else 'FAIL'} {x['name']}" for x in checks)+'\n',encoding='utf-8')
    print(json.dumps({'passed':report['passed'],'checks':len(checks),'failed':[x for x in checks if not x['passed']]},ensure_ascii=False,indent=2))
    if not report['passed']:raise SystemExit(2)

if __name__=='__main__':
    RESULT.mkdir(parents=True,exist_ok=True)
    write_profiles()
    create_source()
    original=SOURCE.read_bytes()
    if '--resume' not in sys.argv:
        run_script('inspect_word.py',SOURCE,'--profile',PROFILE,'--json',RESULT/'原稿分析.json','--md',RESULT/'原稿分析.md')
        run_script('format_word.py',SOURCE,'--profile',PROFILE,'--output',OUTPUT,'--report',RESULT/'基础引擎报告.md','--json-report',RESULT/'基础引擎报告.json')
    supplement()
    run_script('validate_word.py',SOURCE,OUTPUT,'--json',RESULT/'最终完整性.json','--md',RESULT/'最终完整性.md')
    verify_and_render(original)
