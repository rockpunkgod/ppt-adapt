"""Original four-card state navigation foundation, made only with OfficeCLI."""
import argparse,sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parent))
from build_components import Deck

def main():
 p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args()
 d=Deck(a.out,a.officecli,static=True)
 names=['ADA','LIN','MIRA','NOAH']
 roles=['RESEARCH','STRATEGY','DESIGN','DELIVERY']
 copy=['Find the question behind the brief. Compare evidence, record uncertainty, and turn observations into a clear direction.',
       'Choose the work that matters. Connect audience needs with a practical plan and make each decision easy to explain.',
       'Give ideas a visible structure. Use hierarchy, spacing, and consistent forms to help people understand the message.',
       'Bring the pieces together. Check the details, rehearse the experience, and keep the final result easy to maintain.']
 for active in range(5):
  s=d.slide('F3F1EB');ops=[]
  def shape(name,x,y,w,h,text,**kw):
   props=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',text=text,font='Arial',size=20,color='183D40',fill='none',line='none',margin='0.45cm')
   props.update(kw);ops.append({'command':'add','parent':f'/slide[{s}]','type':'shape','props':props})
  shape('title',1.5,1.1,30,1.8,'MEET THE TEAM',size=36,bold=True)
  shape('instruction',1.5,3.0,30,1.2,'Move over a card to explore the team.',size=18)
  for i in range(4):
   selected=active==i+1
   text=f'{names[i]}\n\n{roles[i]}'
   shape(f'hover-card-{i+1}',1.5+i*7.8,5.5,7.2,10.6,text,size=20,fill='1C6561' if selected else 'D5E2DB',color='FFFFFF' if selected else '183D40',geometry='roundRect')
   if selected:shape(f'bio-{i+1}',1.5+i*7.8,9.7,7.2,5.9,copy[i],size=17,color='FFFFFF')
  shape('footer',1.5,17.25,30,1,'Original sample profiles',size=14)
  d.batch(ops)
 d.finish();print(d.file)
if __name__=='__main__':main()
