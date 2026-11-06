"""True native target/actual series, shared zero axis and100% overlap."""
import argparse,sys
from pathlib import Path
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args()
d=Deck(a.out,a.officecli,static=True);s=d.slide('F6F8FC')
d.shape(s,'title',1.7,1.1,30,1.8,'ACTUAL AND TARGET',font='Arial',size=34,bold=True,color='17324E')
d.shape(s,'summary',1.8,3.1,29,1,'Target: 100 tasks per quarter',font='Arial',size=18,color='527086')
d.run('add',f'/slide[{s}]','--type','chart',props=dict(name='target-overlay',chartType='column',categories='Q1,Q2,Q3,Q4',data='Target:100,100,100,100;Actual:45,60,75,90',x='1.8cm',y='4.7cm',width='30cm',height='12cm',title='none',legend='bottom',legendFont='16:17324E:Arial',axisfont='16:527086:Arial',axismin='0',axismax='140',majorunit='20',chartareafill='none',chartborder='none',plotareafill='none',plotborder='none',gridlines='DCE4EE:0.5',colors='1C6561,1C6561',overlap='100',gapwidth='300',transparency='85',dataLabels='none'))
# The tested color setter replaces the second solid fill and clears its alpha.
d.run('set',f'/slide[{s}]/chart[1]/series[2]',props={'color':'1C6561'})
d.shape(s,'footer',1.8,17.4,30,1,'Original sample: actual45 /60 /75 /90 tasks.',font='Arial',size=14,color='527086')
v=d.clone(1)
d.run('set',f'/slide[{v}]/chart[1]/series[1]',props={'values':'120,120,120,120'})
d.set(v,'summary',text='Target: 120 tasks per quarter')
d.finish()
