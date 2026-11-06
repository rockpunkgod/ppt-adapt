import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parent))
from build_components import Deck
import argparse
p=argparse.ArgumentParser()
p.add_argument('--out',required=True)
p.add_argument('--officecli',required=True)
a=p.parse_args()
out=Path(a.out).resolve()
d=Deck(out,a.officecli)
s=d.slide('F5F7FB')
d.heading(s,'连续状态 · 抽拉目录','数字与卡片先做原生形状联合，再组合文案；相邻页仅移动组合。')
colors=['173869','2856A8','4177DB','79A1F0']
for i,c in enumerate(colors,1):
 x=3+(i-1)*7
 d.shape(s,f'card-{i}',x,5.7,9,11.7,fill=c,geometry='roundRect')
 d.shape(s,f'digit-{i}',x-1.25,6.4,4,3.7,str(i),font='Arial',size=96,bold=True,color=c)
 d.shape(s,f'title-{i}',x+.65,10.7,5.7,1.2,['洞察','结构','表达','行动'][i-1],size=26,bold=True,color='FFFFFF')
 d.shape(s,f'body-{i}',x+.65,12.4,5.65,2.6,'先归类核心信息\n再用运动呈现关系',size=17,color='FFFFFF')
d.clone(1)
d.morph(2)
d.finish()
