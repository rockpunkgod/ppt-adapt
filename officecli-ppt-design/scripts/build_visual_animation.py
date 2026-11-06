"""Reproduce selected tutorial mechanisms with OfficeCLI, no PPT authoring library.

Usage: python build.py --officecli PATH --outdir PATH
Outputs are examples with original vector artwork, not copies of tutorial assets.
"""
from __future__ import annotations
import argparse, json, subprocess, struct, zlib, math
from pathlib import Path

ROOT=Path(__file__).resolve().parent
W,H=33.87,19.05
FONT='Microsoft YaHei'

class Deck:
    def __init__(self, exe: str, path: Path):
        self.exe,self.path=exe,path
        self.log=[]
        if path.exists():raise FileExistsError(f'Refusing to overwrite existing presentation: {path}. Choose a new --outdir.')
        self.call('create',str(path),'--locale','zh-CN')
        self.call('set',str(path),'/', '--prop',f'slideWidth={W}cm','--prop',f'slideHeight={H}cm')
        self.n=0
    def call(self,*args):
        proc=subprocess.run([self.exe,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace')
        self.log.append({'args':list(map(str,args)),'code':proc.returncode,'stdout':proc.stdout,'stderr':proc.stderr})
        if proc.returncode or 'UNSUPPORTED' in proc.stdout+proc.stderr:
            raise RuntimeError(str(args)+'\n'+proc.stdout+'\n'+proc.stderr)
        return proc.stdout
    def batch(self,ops):
        for j in range(0,len(ops),12):
            inp=self.path.with_suffix('.batch.json')
            inp.write_text(json.dumps(ops[j:j+12],ensure_ascii=False),encoding='utf-8')
            out=self.call('batch',str(self.path),'--input',str(inp),'--json')
            parsed=json.loads(out)
            if parsed.get('success') is False or 'atomicRolledBack": true' in out:
                raise RuntimeError(out)
    def add(self,typ,parent,**props):
        return {'command':'add','parent':parent,'type':typ,'props':{k:str(v).lower() if isinstance(v,bool) else str(v) for k,v in props.items()}}
    def slide(self,bg='151027',**extra):
        self.n+=1;self.batch([self.add('slide','/',layout='blank',background=bg,**extra)])
        return f'/slide[{self.n}]'
    def shape(self,s,name,x,y,w,h,text='',size=20,color='FFFFFF',fill='none',preset='rect',**extra):
        props=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',text=text,font=FONT,size=size,color=color,fill=fill,preset=preset,line='none',margin='0cm')
        props.update(extra)
        return self.add('shape',s,**props)
    def title(self,s,t,sub=''):
        ops=[self.shape(s,'title',1.5,1.1,30,2,t,size=36,bold=True)]
        if sub:ops.append(self.shape(s,'subtitle',1.5,3.1,29.5,1.2,sub,size=18,color='C7C0D5'))
        return ops
    def note(self,s,t):return self.add('notes',s,text=t)
    def motion_flags(self,s,reverse=False,repeat=False):
        # 1.0.159/160 help advertises these but Add warns unsupported and Set
        # reports success without persisting them. Explicit OOXML is verified.
        if reverse:self.call('raw-set',str(self.path),s,'--xpath',"//p:cTn[@presetClass='path']",'--action','setattr','--xml','autoRev=1')
        if repeat:self.call('raw-set',str(self.path),s,'--xpath',"//p:cTn[@presetClass='path']",'--action','setattr','--xml','repeatCount=indefinite')
    def finish(self):
        self.call('close',str(self.path))
        validation=self.call('validate',str(self.path))
        issues=self.call('view',str(self.path),'issues','--json')
        self.path.with_suffix('.validation.txt').write_text(validation+'\n'+issues,encoding='utf-8')
        self.call('view',str(self.path),'html','-o',str(self.path.with_suffix('.html')))
        self.path.with_suffix('.build-log.json').write_text(json.dumps(self.log,ensure_ascii=False,indent=2),encoding='utf-8')

def artwork(outdir):
    # Repo-native vector illustration, no photo claims and no external downloads.
    art=outdir/'landscape.svg'
    art.write_text('''<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 1600 900"><defs><linearGradient id="sky" x2="0" y2="1"><stop stop-color="#c4e3dd"/><stop offset="1" stop-color="#f5e5bc"/></linearGradient><linearGradient id="water" x2="0" y2="1"><stop stop-color="#477e78"/><stop offset="1" stop-color="#173e40"/></linearGradient></defs><rect width="1600" height="900" fill="url(#sky)"/><circle cx="1120" cy="190" r="96" fill="#f4dc9f"/><path d="M0 530 L150 170 390 510 590 220 800 500 1040 240 1260 540 1450 260 1600 490V900H0Z" fill="#84a99b"/><path d="M0 630 L210 340 430 610 670 370 900 620 1200 340 1430 610 1600 380V900H0Z" fill="#456e62"/><rect y="650" width="1600" height="250" fill="url(#water)"/><path d="M0 850Q400 690 760 760T1600 790V900H0Z" fill="#173c33"/><path d="M980 650v-90h210v90M950 560h270l-135-70Z" fill="#243e39"/><path d="M1010 585h150v18h-150Z" fill="#ecce91"/></svg>''',encoding='utf-8')
    # Native raster fallback authored here, using only stdlib. Also supply true
    # continuous crops for the three window shapes, not repeated whole images.
    pixels=[]
    for y in range(900):
        row=bytearray()
        for x in range(1600):
            t=y/900
            rgb=(int(190-60*t),int(222-55*t),int(216-50*t))
            if (x-1120)**2+(y-190)**2<96**2:rgb=(244,220,159)
            far=465+130*math.sin(x/112)+50*math.cos(x/61)
            near=585+105*math.sin(x/150+.8)+55*math.cos(x/93)
            if y>far:rgb=(119,155,140)
            if y>near:rgb=(60,101,88)
            if y>650:rgb=(int(57-(y-650)*.10),int(109-(y-650)*.17),int(103-(y-650)*.13))
            if y>835+35*math.sin(x/200):rgb=(23,60,51)
            if 980<x<1190 and 560<y<650:rgb=(36,62,57)
            if 950<x<1220 and 490<y<560 and abs(x-1085)<(y-490)*1.93:rgb=(36,62,57)
            row.extend(rgb)
        pixels.append(bytes(row))
    def png(path,rows,width):
        def chunk(kind,data):return struct.pack('!I',len(data))+kind+data+struct.pack('!I',zlib.crc32(kind+data)&0xffffffff)
        raw=b''.join(b'\0'+r for r in rows)
        path.write_bytes(b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('!IIBBBBB',width,900,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(raw,7))+chunk(b'IEND',b''))
    raster=outdir/'landscape.png';png(raster,pixels,1600)
    for i in range(3):
        left=round(i*1600/3);right=round((i+1)*1600/3)
        png(outdir/f'slice-{i}.png',[r[left*3:right*3] for r in pixels],right-left)
    return raster

def zoom_lab(exe,outdir):
    d=Deck(exe,outdir/'zoom-lab.pptx');s=d.slide()
    ops=d.title(s,'从总览进入主题','原生 Slide Zoom 验证作品 · 三个独立目标页')
    for i,(label,preset) in enumerate([('机会','diamond'),('市场','hexagon'),('方案','ellipse')]):
        x=3.0+i*10.0
        ops += [d.shape(s,f'hex-{i}',x,5.1,8,8,fill='FFFFFF',preset='hexagon',opacity=.10,line='DAD4FF:1',rotation=0),d.shape(s,f'label-{i}',x,14.9,8,1.5,label,size=28,align='center',bold=True)]
    ops += [d.note(s,'源教程使用Section Zoom。本作品验证的是OfficeCLI原生Slide Zoom，不冒充Section Zoom。点击缩放对象进入一个目标页并返回。')]
    d.batch(ops)
    for i,(label,value,desc) in enumerate([('机会','27%','识别值得投入的增长空间'),('市场','97M','用一项关键数据建立背景'),('方案','3 步','把判断转成可执行的行动')]):
        s=d.slide('201334')
        ops=d.title(s,label)
        ops += [d.shape(s,'value',3.0,6.0,27,4.2,value,size=90,bold=True,align='center'),d.shape(s,'explanation',3.0,11.4,27,2.0,desc,size=28,align='center'),d.note(s,'源机制：从Hub放大进入主题。此处内容为自创示例。')]
        d.batch(ops)
    ops=[]
    for i in range(3):ops.append(d.add('zoom','/slide[1]',target=i+2,x=f'{3+i*10.0}cm',y='7cm',width='8cm',height='4.5cm',name=f'zoom-{i+1}',returnToParent=True,transitionDur=1000))
    d.batch(ops);d.finish()

def visual_lab(exe,outdir,art):
    d=Deck(exe,outdir/'visual-lab.pptx')
    s=d.slide('F1F3F5');ops=[d.shape(s,'title',1.6,4.5,11.6,3.2,'山水之间',size=46,color='1C3540',bold=True),d.shape(s,'body',1.6,8.4,10.4,5.6,'同一幅主视觉\n分成连续的三个窗口\n让图形承担版式结构',size=22,color='2E4850',lineSpacing='1.5x')]
    for i in range(3):ops.append(d.shape(s,f'slice-{i}',14.2+i*5.8,4.0+i*.7,5.5,9.28,image=outdir/f'slice-{i}.png'))
    # Keep illustration and content moving inward independently, preserving source semantic direction.
    d.batch(ops+[d.note(s,'来源T3gf6MlkcbE 00:16–01:56。源图为三条斜矩形共享图片填充；本页改用三条直立连续裁片，属于明确标注的几何简化，不能当成精确斜切复现。')])
    for i in range(3):d.batch([d.add('animation',f'{s}/shape[@name=slice-{i}]',effect='fly',direction='right',duration=1000,trigger='withPrevious')])
    d.batch([d.add('animation',f'{s}/shape[@name=title]',effect='fly',direction='left',duration=1000,trigger='withPrevious'),d.add('animation',f'{s}/shape[@name=body]',effect='fly',direction='left',duration=1000,trigger='withPrevious')])
    # Parallax scene pair: identical actor names, depth separation, same persistent background.
    for phase in range(2):
        s=d.slide('173C33',**({'transition':'morph-1000'} if phase else {}))
        ops=[d.add('picture',s,src=art,name='!!landscape',x=f'{-phase*3}cm',y='0cm',width='37cm',height='20.8cm',alt='原创山水插画'),d.shape(s,'!!hero-title',2-phase*25,3.2,28,4.1,'山水之间',size=74,bold=True),d.shape(s,'!!reading-panel',31-phase*14,1.8,15.4 if phase else 1.4,15.4,fill='102C2B',opacity=.82),d.shape(s,'!!detail-title',36 if phase==0 else 19,4.0,11.2,1.8,'为内容让出空间',size=30,bold=True),d.shape(s,'!!detail-body',36 if phase==0 else 19,6.8,11.2,6.4,'背景轻移，标题退出。\n阅读面板展开承载说明。\n前景叶片保留在最上层。',size=22,lineSpacing='1.5x')]
        for i,(x,y,rot) in enumerate([(0.2,14.5,-35),(29,0.5,40),(31,3,65)]):ops.append(d.shape(s,f'!!leaf-{i}',x+(phase*.6 if i%2 else -phase*.8),y+(phase*.3),3.8,1.5,fill='0C2C26',preset='ellipse',rotation=rot))
        ops.append(d.note(s,'来源T3gf6MlkcbE 02:01–04:34。此页使用原创矢量山水与原生椭圆前景，机制为同名对象Morph视差；不是原视频叶片照片。'))
        d.batch(ops)
    s=d.slide('111A1D');ops=[d.add('picture',s,src=art,name='editorial-bg',x='0cm',y='0cm',width=f'{W}cm',height=f'{H}cm',opacity=.35,alt='低对比山水背景'),d.shape(s,'title',1.6,1.8,30,3,'用结构替代项目符号',size=44,bold=True)]
    for i,(head,body) in enumerate([('聚焦','一组内容只表达\n一个明确观点'),('分组','相同层级共享\n尺寸与基线'),('留白','间距帮助读者\n分辨信息关系'),('收束','用一句话指向\n下一步行动')]):
        x=1.7+i*7.9
        ops += [d.shape(s,f'head-{i}',x,8.0,6.6,1.7,head,size=30,bold=True),d.shape(s,f'body-{i}',x,10.0,6.6,3.3,body,size=22,lineSpacing='1.5x')]
        if i<3:ops.append(d.shape(s,f'divider-{i}',x+7.1,8.1,.025,6.3,fill='C6D6D2'))
    ops.append(d.note(s,'来源T3gf6MlkcbE 04:47–07:01。直接生成原生文本列组，替代GUI SmartArt快捷操作。背景推进的Grow110%未在本页写入，避免把默认Grow误认成可设110%。'))
    d.batch(ops);d.finish()

def animation_lab(exe,outdir):
    d=Deck(exe,outdir/'animation-lab.pptx')
    s=d.slide('171C23');ops=d.title(s,'一次点击，讲清一个步骤','独立原生对象 · 按顺序点击出现')
    for i,(h,b) in enumerate([('观察','先识别关系'),('判断','再选表达形式'),('验证','最后检查结果')]):
        x=2.2+i*10.5;ops.append(d.shape(s,f'step-{i}',x,6,8.5,6.5,f'{i+1:02d}  {h}\n\n{b}',size=26,bold=True,fill=['425B70','526F63','7A644B'][i],preset='roundRect',margin='.6cm'))
    d.batch(ops+[d.note(s,'来源Top10 00:48–02:38；Fade入口每次点击出现一个对象。')])
    d.batch([d.add('animation',f'{s}/shape[@name=step-{i}]',effect='fade',duration=500,trigger='onClick') for i in range(3)])
    s=d.slide('171C23');ops=d.title(s,'同一拍开始，依次到达','withPrevious + 0 / 100 / 200 / 300 ms 延迟')
    for i in range(4):ops.append(d.shape(s,f'part-{i}',3+i*7.8,6.8,4.5,5.5,str(i+1),size=58,bold=True,fill=['56736A','6C8974','91A27D','C5B787'][i],preset='roundRect',align='center',valign='center'))
    d.batch(ops+[d.note(s,'来源Top10 04:10–06:19。保留0.1秒递增延迟，省略未验证的Fly bounceEnd细节。')])
    d.batch([d.add('animation',f'{s}/shape[@name=part-{i}]',effect='fly',direction='up',duration=1200,trigger='withPrevious',delay=i*100) for i in range(4)])
    s=d.slide('171C23');ops=d.title(s,'去而复返的运动','自定义直线路径 · autoReverse · repeat indefinite')
    ops += [d.shape(s,'track',4,11,25,.06,fill='849293'),d.shape(s,'traveler',4,9.3,2.5,2.5,fill='D4B46C',preset='ellipse'),d.note(s,'来源Top10 06:19–07:43。轨迹长度=画布宽度的0.6，运动中心往返，不用两个不连续的单向动画拼接。')]
    d.batch(ops);d.batch([d.add('animation',f'{s}/shape[@name=traveler]',**{'class':'motion'},path='custom',d='M 0 0 L 0.6 0 E',duration=1500,trigger='withPrevious')]);d.motion_flags(s,reverse=True,repeat=True)
    s=d.slide('171C23');ops=d.title(s,'沿轨道持续运行','光点组合的简化实现 · 单个原生形状沿圆路径')
    ops += [d.shape(s,'orbit',12,5.5,10,10,fill='none',line='A7BBB5:1.5',preset='ellipse'),d.shape(s,'glow',16.4,4.9,1.2,1.2,fill='F2E7BC',preset='ellipse',glow='F2E7BC-12-65'),d.note(s,'来源SlideZoom 25:13–28:59。源视频光点为两圆组合；此处为单圆glow简化。圆路径尺寸需原生实播核验，不能仅凭path=circle认定贴合轨道。')]
    rx,ry,k=5/W,5/H,.55228475
    path=f'M 0 0 C {k*rx} 0 {rx} {(1-k)*ry} {rx} {ry} C {rx} {(1+k)*ry} {k*rx} {2*ry} 0 {2*ry} C {-k*rx} {2*ry} {-rx} {(1+k)*ry} {-rx} {ry} C {-rx} {(1-k)*ry} {-k*rx} 0 0 0 E'
    d.batch(ops);d.batch([d.add('animation',f'{s}/shape[@name=glow]',**{'class':'motion'},path='custom',d=path,duration=6000,trigger='withPrevious')]);d.motion_flags(s,repeat=True)
    # A rectangular tube mask is an exact instance of the same masking mechanism,
    # deliberately not labelled as the curved Erlenmeyer flask from the source.
    s=d.slide('171C23');ops=d.title(s,'让液面只出现在窗口内','前景遮挡 · 背景液面上移 · 原生形状保持可编辑')
    ops += [d.shape(s,'water',13,15.0,8,9.0,fill='4095AD'),d.shape(s,'mask-left',0,4.6,13,H-4.6,fill='171C23'),d.shape(s,'mask-right',21,4.6,W-21,H-4.6,fill='171C23'),d.shape(s,'mask-bottom',13,15.2,8,H-15.2,fill='171C23'),d.shape(s,'tube-outline',13,5.4,8,9.8,fill='none',line='E2ECE9:2'),d.note(s,'来源Top10 20:35–24:04。源为矩形减烧瓶真孔+波浪液面；本页采用矩形窗口的三块原生遮挡，机制相同但轮廓简化。水从窗口下方上移5.5cm。')]
    d.batch(ops);d.batch([d.add('animation',f'{s}/shape[@name=water]',**{'class':'motion'},path='custom',d=f'M 0 0 L 0 {-5.5/H:.8f} E',duration=2000,trigger='onClick')])
    d.finish()

def main():
    p=argparse.ArgumentParser();p.add_argument('--officecli',default='officecli');p.add_argument('--outdir',type=Path,required=True);args=p.parse_args()
    out=args.outdir.resolve()
    existing=[out/(n+'.pptx') for n in ['zoom-lab','visual-lab','animation-lab'] if (out/(n+'.pptx')).exists()]
    if existing:p.error('Refusing to overwrite existing output: '+str(existing[0])+'; choose a new --outdir.')
    out.mkdir(parents=True,exist_ok=True);art=artwork(out)
    zoom_lab(args.officecli,out);print('zoom-lab generated',flush=True)
    visual_lab(args.officecli,out,art);print('visual-lab generated',flush=True)
    animation_lab(args.officecli,out);print('animation-lab generated',flush=True)
if __name__=='__main__':main()
