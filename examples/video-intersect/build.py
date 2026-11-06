from pathlib import Path
import sys,argparse,subprocess
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'officecli-ppt-design/scripts'))
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--out',required=True);p.add_argument('--officecli',required=True);p.add_argument('--ffmpeg',required=True);a=p.parse_args()
out=Path(a.out).resolve();assets=out.parent/'assets'
if out.exists() or assets.exists():raise FileExistsError('Fresh output and assets directory required')
assets.mkdir(parents=True)
vf=r"drawtext=fontfile='C\:/Windows/Fonts/arial.ttf':text='SOURCE %{pts\:hms}':fontsize=36:fontcolor=white:box=1:boxcolor=black@0.8:x=(w-tw)/2:y=(h-th)/2"
subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error','-f','lavfi','-i','testsrc2=size=960x540:rate=30:duration=8','-vf',vf,'-c:v','libx264','-preset','fast','-pix_fmt','yuv420p',str(assets/'original-motion.mp4')],check=True)
subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error','-i',str(assets/'original-motion.mp4'),'-frames:v','1',str(assets/'poster.png')],check=True)
d=Deck(out,a.officecli)
for i in [1,2]:
 s=d.slide('16251D');ops=[]
 for n,t,x,y,w,h,size,c in [('kicker','TRACE / 原生视频',1.5,1.1,30,1,15,'DAEF78'),('title','让画面，\n拥有自己的轮廓。',1.5,4.5,12,4,35,'F2F3EC'),('body','真实媒体，保留时间。\n原生几何，重新定义边界。',1.5,10,12,3,20,'BCC7BB'),('caption','裸视频对照' if i==1 else '视频自身与组合形状相交',1.5,17.2,30,1,14,'BCC7BB')]:
  ops.append(dict(command='add',parent=f'/slide[{s}]',type='shape',props=dict(name=n,text=t,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',size=str(size),font='Noto Serif SC' if n=='title' else 'Noto Sans SC',color=c,fill='none',line='none',margin='0cm')))
 ops.append(dict(command='add',parent=f'/slide[{s}]',type='video',props=dict(name='video',src=str(assets/'original-motion.mp4'),poster=str(assets/'poster.png'),x='15cm',y='4.5cm',width='16cm',height='9cm',autoPlay=True,loop=False,volume=0)))
 d.batch(ops);d.notes(s,'原创8秒诊断媒体；非真实产品内容。第一页面为裸视频控制；第二页面使用原生媒体与组合几何相交。')
d.run('close');print(d.file)
