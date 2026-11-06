"""Reproducible OfficeCLI adaptations of three mechanisms from the masterclass.

All slide objects are editable native shapes. No third-party Python libraries.
"""
import argparse,json,subprocess,shutil,sys
from pathlib import Path

def build(output,cli):
    output=Path(output).resolve()
    if output.exists(): raise FileExistsError(f'Choose a new output file: {output}')
    output.parent.mkdir(parents=True,exist_ok=True)
    def run(*args):
        p=subprocess.run([cli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
        if p.returncode: raise RuntimeError(p.stdout+'\n'+p.stderr)
        return p.stdout
    run('create',output)
    run('set',output,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
    commands=[]
    def add(parent,kind,**props):commands.append({'command':'add','parent':parent,'type':kind,'props':{k:str(v) for k,v in props.items()}})
    def shape(s,name,x,y,w,h,text='',size=20,color='F3F4F1',fill='none',**props):
        add(f'/slide[{s}]','shape',name=name,x=f'{x:.3f}cm',y=f'{y:.3f}cm',width=f'{w:.3f}cm',height=f'{h:.3f}cm',text=text,size=size,color=color,fill=fill,line='none',font='Microsoft YaHei',margin='0cm',**props)
    def title(s,kicker,heading,desc):
        shape(s,f's{s}-kicker',1.6,1.2,28,.6,kicker,13,'99B9B2',spacing=2)
        shape(s,f's{s}-title',1.6,2.2,30,2,heading,36,bold='true')
        shape(s,f's{s}-desc',1.6,4.35,28,1.1,desc,18,'B6C4C4')
    # Create targets before links, since OfficeCLI validates slide targets.
    for s in range(1,6):add('/','slide',layout='blank',background='122B2E')
    title(1,'MASTERCLASS / NATIVE OBJECTS','从目录，到细节','点击年份进入对应状态；点击右上角返回目录。')
    for j,(year,caption) in enumerate([('1980','移动通话'),('2000','随身通信'),('2020','触屏生态')]):
        x=1.6+j*10.2
        shape(1,f'home-card-{j}',x,7.0,9.2,8.7,fill='1D3B3E',preset='roundRect')
        shape(1,f'home-year-{j}',x+.65,7.9,7.9,2.4,year,48,'D4B57B',bold='true')
        shape(1,f'home-label-{j}',x+.65,10.4,7.9,1.3,caption,26)
        shape(1,f'home-go-{j}',x+.65,13.5,7.9,.85,'查看这一阶段  →',17,'99B9B2')
        shape(1,f'home-hit-{j}',x,7.0,9.2,8.7,fill='FFFFFF',opacity='0',preset='roundRect',link=f'slide[{j+2}]')
    shape(1,'home-rule',1.6,17.4,30.6,.035,fill='547273')
    shape(1,'home-foot',1.6,17.85,29,.6,'课程机制复现：内部链接 + 一致对象匹配 + 时间轴平移',12,'99B9B2')
    years=['1980','2000','2020']
    descriptions=['一部能带着走的电话','从通话工具到随身终端','从硬件按钮到软件界面']
    # Keep phone components identical across adjacent states and translate/scale them.
    for active in range(3):
        s=active+2
        add(f'/slide[{s}]','notes',text=f'来源 2:56:41–3:05:39。当前年份 {years[active]}。年份轨道左移，当前手机放大，其余缩小。素材为原创原生图形；不是原视频照片。')
        title(s,'EVOLUTION / ONE CONTINUOUS TIMELINE',descriptions[active],'以同一条时间线承载前后变化，让当前阶段成为唯一焦点。')
        shape(s,f's{s}-back',28.4,1.2,3.8,.7,'返回目录 ↗',14,'DCE6DA',align='right')
        shape(s,f's{s}-back-hit',28.4,1.2,3.8,.7,fill='FFFFFF',opacity='0',link='slide[1]')
        shape(s,'!!axis',1.6,15.8,30.6,.045,fill='547273')
        for j,year in enumerate(years):
            cx=6.1+(j-active)*7.9
            # off-canvas intentional states preserve object identity.
            shape(s,f'!!year-circle-{j}',cx-1.7,14.1,3.4,3.4,year,24,'122B2E' if j==active else 'DCE6DA','D4B57B' if j==active else '122B2E',preset='ellipse',align='center',valign='center')
            shape(s,f'!!year-hit-{j}',cx-1.7,14.1,3.4,3.4,fill='FFFFFF',opacity='0',preset='ellipse',link=f'slide[{j+2}]')
            # Hand-drawn phone family, centered over each timeline station.
            scale=1 if j==active else .52
            px=cx-1.6*scale;py=7.0+(1-scale)*4.5
            def part(k,dx,dy,w,h,**p):shape(s,f'!!phone-{j}-{k}',px+dx*scale,py+dy*scale,w*scale,h*scale,**p)
            part('case',0,0,3.2,6.1,fill=['34464A','587779','273F43'][j],preset='roundRect')
            part('screen',.28,.48,2.64,[1.45,2.1,5.1][j],fill=['99B9B2','B1C9B8','407F7D'][j],preset='roundRect')
            if j<2:
                for row in range(4):
                    for col in range(3):part(f'key-{row}-{col}',.45+col*.8,2.8+row*.62,.65,.37,fill='A7B8AD',preset='roundRect')
            else:
                for row in range(3):
                    for col in range(3):part(f'app-{row}-{col}',.5+col*.8,1.2+row*.9,.55,.55,fill=['D4B57B','99B9B2','DCE6DA'][(row+col)%3],preset='roundRect')
            if j==0:part('antenna',2.45,-1.5,.23,1.7,fill='99B9B2')
        commands.append({'command':'set','path':f'/slide[{s}]','props':{'transition':'morph-750'}})
    # Delay-trail mechanism: identical custom path with per-copy delays, no invented trail effect.
    title(5,'ANIMATION / DELAY AS A DESIGN VARIABLE','同一路径，不同时间','同一条路径上的短延迟，让多个对象形成连续的运动轨迹。')
    for j in range(6,-1,-1):
        color=['F3F4F1','D4B57B','AA996F','7C8770','527475','385A5F','25464B'][j]
        shape(5,f'trail-{j}',11.6,9.0,10.5,3.1,'FLOW',64,color,bold='true',align='center')
        add(f'/slide[5]/shape[@name=trail-{j}]','animation',**{'class':'motion','path':'custom','d':'M 0 0 C 0.09 -0.12 0.18 0.12 0.27 0 E','duration':3000,'delay':j*70,'trigger':'withPrevious'})
    shape(5,'trail-info',1.6,16.1,30.5,.9,'7 个原生文本对象 · 相同路径 · 每层相差 70 ms',16,'99B9B2')
    shape(5,'trail-back',28.4,1.2,3.8,.7,'返回目录 ↗',14,'DCE6DA',align='right')
    shape(5,'trail-back-hit',28.4,1.2,3.8,.7,fill='FFFFFF',opacity='0',link='slide[1]')
    add('/slide[5]','notes',text='来源 5:54:42–6:01:30。延迟形成拖尾。该样例用短贝塞尔路径替代源视频环形路径，机制相同但外观不是逐像素复刻。动态效果需实播。')
    batch=output.with_suffix('.batch.json');batch.write_text(json.dumps(commands,ensure_ascii=False,indent=2),encoding='utf-8')
    result=run('batch',output,'--input',batch,'--json');output.with_suffix('.build-log.json').write_text(result,encoding='utf-8')
    parsed=json.loads(result)
    if parsed.get('success') is False or parsed.get('ok') is False or 'atomicRolledBack": true' in result:raise RuntimeError('OfficeCLI batch rejected; inspect build log')
    run('close',output)
    validation=run('validate',output,'--json');output.with_suffix('.validation.json').write_text(validation,encoding='utf-8')
    html=output.with_suffix('.html');run('view',output,'html','-o',html)
    print(output)

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--out',required=True);p.add_argument('--officecli',default=shutil.which('officecli') or 'officecli');a=p.parse_args();build(a.out,a.officecli)
