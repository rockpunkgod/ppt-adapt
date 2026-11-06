"""OfficeCLI foundations for a native single-textbox, by-letter Wave example."""
from pathlib import Path
import argparse,sys
sys.path.insert(0,str(Path(__file__).resolve().parent))
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',type=Path,required=True);a=p.parse_args()
d=Deck(a.out,a.officecli)
for i,(title,footer) in enumerate([
 ('让标题，起一点波澜。','单击开始 / 原生逐字Wave · 0.5秒 · 字母延迟10%'),
 ('慢下来，看清每一个字母。','单击开始 / 2秒慢放 · 纵向位移0.04页高'),
 ('改变幅度，不改变身份。','单击开始 / 2秒慢放 · 纵向位移0.08页高'),
 ('留白，也是一种节奏。','静态对照 / 同一个可编辑文本框，没有动画。')],1):
 s=d.slide('16251D');d.shape(s,'kicker',1.5,1.1,30,1,'TRACE / 逐字波浪',font='Noto Sans SC',size=15,color='DAEF78')
 d.shape(s,'title',1.5,3,30,2,title,font='Noto Serif SC',size=33,color='F2F3EC')
 d.shape(s,'wave-target',5.5,9.2,23,3,'DESIGN',font='Arial',size=72,color='DAEF78')
 d.run('set',f'/slide[{s}]/shape[@name=wave-target]',props={'spacing':'24','autoFit':'none'})
 d.shape(s,'baseline',5.5,12.1,23,.025,fill='637066')
 d.shape(s,'footer',1.5,17,31,1,footer,font='Noto Sans SC',size=14,color='BCC7BB')
 d.notes(s,'One editable textbox. Native Wave with letter iteration; no separate letter shapes or rendered animation. Fourth page is a static control.')
d.run('close');print(a.out)
