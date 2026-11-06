"""OfficeCLI native layout specimens from five verified Bilibili short tutorials.
Images are caller-provided assets; no video artwork or slide bitmap is embedded.
"""
import argparse,json,subprocess,struct
from pathlib import Path

def jpeg_size(path):
    with open(path,'rb') as f:
        if f.read(2)!=b'\xff\xd8':raise ValueError('Provide JPEG assets')
        while True:
            b=f.read(1)
            while b!=b'\xff':b=f.read(1)
            while b==b'\xff':b=f.read(1)
            marker=b[0];length=struct.unpack('>H',f.read(2))[0]
            if marker in (0xC0,0xC1,0xC2):
                _,h,w=struct.unpack('>BHH',f.read(5));return w,h
            f.seek(length-2,1)

def build(out,cli,landscape,portrait):
    out=Path(out).resolve()
    if out.exists():raise FileExistsError('Choose a new file; refusing overwrite')
    out.parent.mkdir(parents=True,exist_ok=True)
    images=[Path(landscape).resolve(),Path(portrait).resolve()]
    dimensions={str(p):jpeg_size(p) for p in images}
    def run(*args):
        result=subprocess.run([cli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
        if result.returncode:raise RuntimeError(result.stdout+'\n'+result.stderr)
        return result.stdout
    run('create',out);run('set',out,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
    commands=[]
    def add(parent,kind,**props):commands.append({'command':'add','parent':parent,'type':kind,'props':{k:str(v) for k,v in props.items()}})
    def shape(s,name,x,y,w,h,text='',size=20,color='152C35',fill='none',**props):
        p=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',text=text,size=size,color=color,font='Microsoft YaHei',margin='0cm',line='none',fill=fill)
        p.update(props)
        if 'image' in p:p.pop('fill',None)
        add(f'/slide[{s}]','shape',**p)
    def crop(path,w,h,focus=.5):
        iw,ih=dimensions[str(path)];ar=iw/ih;target=w/h
        if ar>target:
            trim=1-target/ar;return f'{round(trim*focus*100000)},0,{round(trim*(1-focus)*100000)},0'
        trim=1-ar/target;return f'0,{round(trim*focus*100000)},0,{round(trim*(1-focus)*100000)}'
    def photo(s,name,path,x,y,w,h,focus=.5,**props):
        shape(s,name,x,y,w,h,image=str(path),srcRect=crop(path,w,h,focus),**props)
    def note(s,text):add(f'/slide[{s}]','notes',text=text)
    def heading(s,kicker,title,subtitle='',dark=False):
        c='FFFFFF' if dark else '152C35';muted='BACAD0' if dark else '50636B'
        shape(s,f'kicker-{s}',1.6,1.2,30,.65,kicker,12,muted,spacing=1.5)
        shape(s,f'title-{s}',1.6,2.5,30,2.5,title,44,c,bold='true')
        if subtitle:shape(s,f'desc-{s}',1.6,5.2,29,1.5,subtitle,20,muted)
    backgrounds=['F6F3ED','F6F3ED','132930','142C32','142C32','F6F3ED','FFFFFF','F6F3ED','F6F3ED','132930','F6F3ED','F6F3ED','132930']
    for c in backgrounds:add('/','slide',layout='blank',background=c)
    # 1: stable family, distinct weights, optically compatible Latin.
    heading(1,'TYPOGRAPHY / 01','字体先统一气质，再建立主次。','同一正文保持稳定笔画；标题与数据用字重区分。')
    for i,(head,body) in enumerate([('笔画粗细','正文笔画保持一致，\n不要频繁更换字重。'),('字形风格','中英文字形相近，\n才有统一的阅读气质。'),('视觉高度','字号不是字面高度，\n仍要实际并排观察。')]):
        x=1.6+i*10.3;shape(1,f'font-card-{i}',x,8,9.3,7.6,fill='FFFFFF')
        shape(1,f'font-rule-{i}',x+.5,8.7,8.3,1.3,head,26,bold='true')
        shape(1,f'font-copy-{i}',x+.5,10.7,8.3,2.6,body,20,lineSpacing='1.2x')
    shape(1,'font-latin',2.1,13.8,8.2,1.1,'PRODUCT / 2026',24,font='Arial',bold='true')
    shape(1,'font-sample-cn',12.4,13.8,8.2,1.1,'产品观察',24,bold='true')
    shape(1,'font-sample-en',22.7,13.8,8.2,1.1,'Product Notes',23,font='Arial')
    note(1,'BV1tz4y1C7PJ 00:49–02:51。三项核查：笔画/字形/视觉高度。使用本机字体替代源字体包，不复制作者商业字库或宣称字体许可。')
    # 2: alignment, paragraph separation, hierarchy, one accent.
    heading(2,'TYPE LAYOUT / 02','重要信息，提到第一眼。','先决定阅读顺序，再决定文字的形状。')
    shape(2,'text-main',2,8,20,2.7,'先讲结论，再展开依据。',36,bold='true')
    shape(2,'text-accent',2,11.5,1.4,.1,fill='BB663E')
    shape(2,'text-p1',2,12.5,20,2.6,'把长段文字拆成几个完整的意思。\n每段共享一条左对齐轴线。',21,lineSpacing='1.25x')
    shape(2,'text-p2',2,15.8,20,1.5,'段落之间留出空隙，重要词只用一个强调色。',21)
    shape(2,'text-side',25,8.5,6,7,'对齐\n留白\n对比',30,'BB663E',lineSpacing='1.6x',bold='true')
    note(2,'BV19V411U7hs 02:15–07:02。左对齐不用空格凑，段间留白，标题/正文以大小字重和主题色建立主次。')
    # 3: real glyph outlines are applied by COM, not rectangle borders.
    for row in range(6):
        for col in range(4):shape(3,f'outline-{row}-{col}',.5+col*8.4,.5+row*3.1,8.1,2.5,'READ',54,'47646B',bold='true',font='Arial',rotation=-8)
    shape(3,'repeat-panel',3.8,6.6,26.2,5.8,fill='132930')
    shape(3,'repeat-main',4.8,7.5,24.2,4.2,'让背景低调，\n让内容清晰。',44,'FFFFFF',bold='true',align='center')
    note(3,'BV19V411U7hs 07:03–09:19。装饰文字重复为纹理，主信息清晰。COM仅精修outline-*的字形填充/线条，其他文字保持可编辑。')
    # 4: closed visual unit for Chinese title.
    photo(4,'frame-photo',images[0],0,0,33.87,19.05)
    shape(4,'frame-dim',0,0,33.87,19.05,fill='071B24',opacity=.6)
    shape(4,'title-closure',4.3,5.4,25.2,7.2,fill='none',line='DBC796:1.2')
    shape(4,'frame-cn',5.3,6.8,23.2,3.0,'山河之间',58,'F8EEDC',font='SimSun',align='center')
    shape(4,'frame-small',5.3,10.5,23.2,1.1,'把独立的字，收成一个整体。',22,'F8EEDC',align='center',spacing=1)
    note(4,'BV1TohbzdEAL 01:04–01:39。闭合修饰把汉字连成视觉块；不是所有主题都必须画矩形框。本作使用普通宋体和不同风景，未复制源书法或背景。')
    # 5: separate text boxes provide rhythm without scaling glyph width.
    photo(5,'rhythm-photo',images[0],0,0,33.87,19.05)
    shape(5,'rhythm-dim',0,0,33.87,19.05,fill='071B24',opacity=.65)
    shape(5,'rhythm-cn-a',5,6,11,3.6,'数 字',62,'F6EBD4',font='SimSun',bold='true')
    shape(5,'rhythm-cn-b',15.5,7,12,3.6,'乡 村',62,'F6EBD4',font='SimSun',bold='true')
    shape(5,'rhythm-gap',6,10.4,8.7,1.1,'连接每一种可能',18,'F6EBD4',align='right')
    shape(5,'rhythm-bottom',5,14.8,25,1.1,'信息服务 · 社区建设 · 共同参与',20,'E3E9E7',align='center',spacing=1)
    note(5,'BV1TohbzdEAL 01:47–02:13。封面短标题拆为两个文字对象，轻微高低错落；小字穿插空隙。未使用栅格化字，也不将长正文随机错位。')
    # 6: dense Chinese reorganized by meaning.
    heading(6,'CHINESE DENSITY / 03','四个行动，让社区服务更清晰。','先分语义，再统一轴线与层级；字号不能靠缩小来挤满。')
    cards=[('看见需求','先收集居民真实问题，\n把相近诉求归到一起。'),('连接资源','明确每项服务由谁负责，\n公开可联络的入口。'),('缩短流程','减少重复填写与来回跑动，\n保留必要核验步骤。'),('持续复核','定期查看反馈，\n对照目标修正安排。')]
    for i,(h,b) in enumerate(cards):
        x=1.6+(i%2)*15.8;y=8+(i//2)*4.7
        shape(6,f'density-card-{i}',x,y,14.6,4.3,fill='FFFFFF')
        shape(6,f'density-tag-{i}',x+.5,y+.5,1.4,1.1,str(i+1).zfill(2),22,'BB663E',bold='true',font='Arial')
        shape(6,f'density-head-{i}',x+2.2,y+.5,11.8,1.4,h,26,bold='true')
        shape(6,f'density-body-{i}',x+2.2,y+1.8,11.8,2.3,b,20,lineSpacing='1.2x')
    note(6,'BV1TohbzdEAL 02:17–03:27。长中文换清晰黑体、增加留白、大小/字重、按轴线分组。社区文案为原创演示，不是源医疗数据。')
    # 7: true one-row/five-column native table, animated number and text.
    photo(7,'editorial-photo',images[0],0,0,33.87,19.05)
    add('/slide[7]','table',name='editorial-grid',rows=1,cols=5,x='0cm',y='0cm',width='33.87cm',height='19.05cm',style='none',headerFill='15212B',firstRow='false',**{'border.all':'none','border.vertical':'1pt solid ADB8C0'})
    for i,(h,b) in enumerate([('观察','先理解受众，\n再确定内容。'),('提炼','用一句结论\n组织一页信息。'),('分组','让相关内容\n靠近并成组。'),('对齐','以统一轴线\n降低阅读负担。'),('复核','回到放映环境\n检查实际效果。')]):
        x=i*6.774
        shape(7,f'ed-num-{i}',x+.55,3.4,5.67,3.2,str(i+1).zfill(2),58,'FFFFFF',font='Georgia',align='center')
        shape(7,f'ed-text-{i}',x+.55,8.4,5.67,7,h+'\n\n'+b,20,'FFFFFF',align='center',lineSpacing='1.2x')
        add(f'/slide[7]/shape[@name=ed-num-{i}]','animation',effect='float',**{'class':'entrance','duration':400,'trigger':'withPrevious' if i==0 else 'afterPrevious'})
        add(f'/slide[7]/shape[@name=ed-text-{i}]','animation',effect='fly',**{'class':'entrance','direction':'up','duration':400,'trigger':'afterPrevious'})
    note(7,'BV1T54y1c7EU 00:21–01:59。整图背景+1行5列表格；全表灰层，2/4列更透明，内边框；序号float、文字fly按afterPrevious逐项。本作深灰替代中灰以保证白色正文对比，COM设置45%/55%透明度。背景可替换，其他元素不动。')
    # 8: waterfall gallery, identical gutters, original images retained as fills.
    heading(8,'IMAGE ARRANGEMENT / 04','顺着列宽，让图片自然呼吸。','瀑布流固定列宽，图片高度随内容变化；标题保留独立区域。')
    gallery=[(0,0,6.9,5.9),(0,6.3,6.9,4.0),(7.4,0,6.9,3.8),(7.4,4.2,6.9,6.1),(14.8,0,6.9,5.0),(14.8,5.4,6.9,4.9)]
    for i,(x,y,w,h) in enumerate(gallery):photo(8,f'waterfall-{i}',images[i%2],10.4+x,7.2+y,w,h,focus=.2+.12*i)
    shape(8,'waterfall-note',1.6,9.8,7.4,5.5,'列宽相同\n间隙一致\n不压扁图片',25,'BB663E',lineSpacing='1.5x')
    note(8,'BV1SY411J732 00:07–00:14。瀑布流。示例用两张合法公开验证图片与不同裁切，非六个不同摄影作品。')
    # 9: tilt varies orientation; identity and meaningful title remain fixed.
    heading(9,'IMAGE ARRANGEMENT / 05','一点倾斜，让画面更灵动。','同一组照片保持白边与间距，角度只作轻微变化。')
    for i in range(6):
        x=2.1+(i%3)*10.1;y=8+(i//3)*4.7;a=[-5,3,-3,4,-4,5][i]
        shape(9,f'tilt-border-{i}',x-.13,y-.13,9.26,4.26,fill='FFFFFF',rotation=a)
        photo(9,f'tilt-photo-{i}',images[i%2],x,y,9,4,focus=.15+.14*i,rotation=a)
    note(9,'BV1SY411J732 00:14–00:18。轻微倾斜图片。本例最大5度，原视频更斜；未测试源平滑动画，只复现静态倾斜构图。')
    # 10: tiled picture wall and one native dark mask.
    for i in range(6):photo(10,f'tile-{i}',images[i%2],(i%3)*11.29,(i//3)*9.525,11.29,9.525,focus=.15+.14*i)
    shape(10,'gallery-mask',0,0,33.87,19.05,fill='081A24',opacity=.7)
    shape(10,'gallery-title',3,7,27.87,3.3,'山 河 之 间',56,'FFFFFF',font='SimSun',align='center')
    shape(10,'gallery-caption',3,11,27.87,1.4,'每一帧观察，汇成同一个主题。',22,'E8ECEC',align='center')
    note(10,'BV1SY411J732 00:18–00:23。图片平铺，覆一层蒙版承接主题；普通暗色形状，不冒称背景虚化。')
    # 11: real preset shape image clips.
    heading(11,'IMAGE WINDOW / 06','改变窗口，不改变图片比例。','矩形、六边形与三角形，都可以成为图片的原生形状窗口。')
    for i,g in enumerate(['rect','hexagon','triangle']):
        photo(11,f'clip-{g}',images[i%2],2+i*10.3,8,9.3,7.5,geometry=g)
        shape(11,f'clip-label-{i}',2+i*10.3,16,9.3,1.3,['矩形','六边形','三角形'][i],23,align='center')
    note(11,'BV1SY411J732 00:23–00:33。原生prstGeom+image fill+srcRect实现真裁切，非用背景色矩形假遮。')
    # 12: title is the center; pictures organize the perimeter.
    positions=[(1.5,1.5,8,4),(13,1.5,8,4),(24.5,1.5,8,4),(1.5,13.5,8,4),(13,13.5,8,4),(24.5,13.5,8,4)]
    for i,(x,y,w,h) in enumerate(positions):photo(12,f'around-{i}',images[i%2],x,y,w,h,focus=.15+.14*i)
    shape(12,'around-title',3,7.4,27.87,2.7,'山河之间',48,align='center',font='SimSun',bold='true')
    shape(12,'around-copy',3,10.4,27.87,1.2,'用周围的图像，支撑中心的主题。',22,'50636B',align='center')
    note(12,'BV1SY411J732 00:33–00:37。图片围绕主题，不遮挡核心文字；当前页纯静态。')
    # 13: separate letter glyph image fill, finalized via COM.
    for i,ch in enumerate('VIEW'):shape(13,f'glyph-{i}',1.5+i*7.8,4.3,7.5,8,ch,168,'FFFFFF',font='Arial Black',align='center')
    shape(13,'glyph-caption',3,14.8,27.87,1.3,'把图片藏进字形，保留可编辑的文字。',23,'FFFFFF',align='center')
    note(13,'BV1SY411J732 00:37–00:41。英文放大，每个字母填不同图片。COM Font.Fill.UserPicture填进真实字形，不用文字框背景图冒充；仍需原生渲染与结构检查。')
    batch=out.with_suffix('.batch.json');batch.write_text(json.dumps(commands,ensure_ascii=False,indent=2),encoding='utf-8')
    result=run('batch',out,'--input',batch,'--json');out.with_suffix('.build-log.json').write_text(result,encoding='utf-8')
    parsed=json.loads(result)
    if parsed.get('success') is False or parsed.get('ok') is False:raise RuntimeError(result)
    run('close',out)
    out.with_suffix('.validation.json').write_text(run('validate',out,'--json'),encoding='utf-8')
    print(out)

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--out',required=True);p.add_argument('--officecli',required=True);p.add_argument('--landscape',required=True);p.add_argument('--portrait',required=True);a=p.parse_args();build(a.out,a.officecli,a.landscape,a.portrait)
