"""Real concentric pie/ring charts and complement cards; CLI-only authoring."""
import argparse,sys,math
from pathlib import Path
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True)
p.add_argument('--values',default='55,20,15,10');p.add_argument('--changed-values',default='40,30,20,10');p.add_argument('--cards',default='65,25,10');a=p.parse_args()
def percentages(raw,n):
 try:vals=[float(v) for v in raw.split(',')]
 except ValueError:p.error('Percentages must be numeric')
 if len(vals)!=n or any(not math.isfinite(v) or v<0 or v>100 for v in vals) or not math.isclose(sum(vals),100,abs_tol=1e-8):p.error(f'Expected {n} finite nonnegative percentages totalling 100')
 return vals
original=percentages(a.values,4);changed=percentages(a.changed_values,4);cards=percentages(a.cards,3)
d=Deck(a.out,a.officecli,static=True)
palette=['1C6561','607F9A','D5AD6E','9684AE']
d.run('set','/',props={f'theme.color.accent{i+1}':c for i,c in enumerate(palette)})
def chart(page,name,kind,x,y,size,values,categories,holes=85):
 props=dict(name=name,chartType=kind,categories=','.join(categories),data='Share:'+','.join(map(str,values)),x=f'{x}cm',y=f'{y}cm',width=f'{size}cm',height=f'{size}cm',title='none',legend='none',chartareafill='none',chartborder='none',plotareafill='none',plotborder='none',varyColors=True,firstSliceAngle=0,dataLabels='none')
 if kind=='doughnut':props['holeSize']=holes
 d.run('add',f'/slide[{page}]','--type','chart',props=props)
cats=['Research','Design','Build','Review']
for vals in [original]:
 s=d.slide('F6F8FC');d.heading(s,'ONE COMPOSITION, TWO LAYERS','The outer ring repeats the same four categories.')
 d.shape(s,'upper-band',2.5,4.7,13.8,5.8,fill='E6EFED')
 d.shape(s,'lower-band',2.5,10.5,13.8,6.1,fill='EEDFCB')
 chart(s,'outer-ring','doughnut',3.5,4.5,12,vals,cats)
 chart(s,'inner-pie','pie',5,6,9,vals,cats)
 for j,(cat,val,col) in enumerate(zip(cats,vals,palette)):
  y=6.4+j*2.25
  d.shape(s,f'swatch-{j}',19,y,.45,.45,fill=f'accent{j+1}')
  d.shape(s,f'category-{j}',19.8,y-.25,7,1,cat,font='Arial',size=22,color='17324E')
  d.shape(s,f'value-{j}',27.3,y-.25,4,1,f'{val:g}%',font='Arial',size=22,bold=True,color='17324E')
 d.shape(s,'total',19,15.7,12,1,'Total: 100% | illustrative data',font='Arial',size=14,color='527086')
clone=d.clone(1)
for n in [1,2]:d.run('set',f'/slide[{clone}]/chart[{n}]/series[1]',props={'values':','.join(map(str,changed))})
for j,val in enumerate(changed):d.set(clone,f'value-{j}',text=f'{val:g}%')
s=d.slide('F6F8FC');d.heading(s,'THREE PARTS OF ONE WHOLE','Separate example. Green: named share. Blue: remainder.')
for j,(name,v) in enumerate(zip(['Research','Design','Review'],cards)):
 x=1.8+j*10.6
 d.shape(s,f'card-{j}',x,5,9.4,11.2,fill='FFFFFF',geometry='roundRect')
 d.shape(s,f'label-{j}',x+.5,5.7,8.4,1,name,font='Arial',size=22,bold=True,color='17324E')
 chart(s,f'card-ring-{j}','doughnut',x+.7,7.4,8,[v,100-v],['Share','Remainder'],holes=68)
 d.shape(s,f'card-value-{j}',x+2.7,10.8,4,1.2,f'{v:g}%',font='Arial',size=28,bold=True,color='17324E',align='center')
d.finish()
