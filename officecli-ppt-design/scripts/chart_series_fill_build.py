"""Native gradient-series chart fixture; pattern polish is a separate step."""
import argparse,sys
from pathlib import Path
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args()
d=Deck(a.out,a.officecli,static=True);s=d.slide('102338')
d.heading(s,'SERIES GRADIENT','A native chart, one scale, four observed values')
d.set(s,'heading-1',color='FFFFFF');d.set(s,'caption-1',color='A5BBD0')
d.run('add','/slide[1]','--type','chart',props=dict(name='styled-series',chartType='column',categories='Q1,Q2,Q3,Q4',data='Actual:45,60,75,90',x='2cm',y='5cm',width='29.8cm',height='11.5cm',title='none',legend='none',axisfont='16:A5BBD0:Arial',axismin=0,axismax=100,majorunit=20,chartareafill='none',chartborder='none',plotareafill='none',plotborder='none',gridlines='29465F:0.5',gapwidth=160,gradient='28D9E7-102338:90',dataLabels='none'))
d.finish()
