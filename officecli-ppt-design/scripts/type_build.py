"""Stdlib orchestration of native OfficeCLI objects; optional COM polish is separate."""
import argparse,json,subprocess
from pathlib import Path
W,H=33.87,19.05
class Deck:
 def __init__(self,cli,path):self.cli=cli;self.path=path;self.ops=[];self.s=0
 def call(self,*args):
  r=subprocess.run([self.cli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace')
  if r.returncode:raise RuntimeError(r.stdout+r.stderr)
  return r.stdout
 def start(self):
  if self.path.exists():raise FileExistsError(self.path)
  self.call('create',self.path);self.call('open',self.path)
  self.call('set',self.path,'/','--prop',f'slideWidth={W}cm','--prop',f'slideHeight={H}cm')
 def add(self,parent,typ,**props):self.ops.append(dict(command='add',parent=parent,type=typ,props={k:str(v).lower() if isinstance(v,bool) else str(v) for k,v in props.items()}))
 def flush(self):
  if not self.ops:return
  f=self.path.with_suffix('.batch.json');f.write_text(json.dumps(self.ops),encoding='utf-8');r=self.call('batch',self.path,'--input',f,'--json')
  if '"success": false' in r:raise RuntimeError(r)
  self.ops=[]
 def new(self,bg='F7F5F0'):
  self.flush();self.s+=1;self.add('/','slide',layout='blank',background=bg);self.flush()
 def shape(self,name,x,y,w,h,text='',size=22,font='Arial',color='152F2B',fill='none',**props):
  p=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',text=text,size=size,font=font,color=color,fill=fill,line='none',margin='0cm',valign='top',autoFit='none');p.update(props);self.add(f'/slide[{self.s}]','shape',**p)
 def notes(self,s):self.add(f'/slide[{self.s}]','notes',text=s)
 def finish(self):
  self.flush();self.call('close',self.path)
  self.path.with_suffix('.validation.json').write_text(self.call('validate',self.path,'--json'),encoding='utf-8')
  self.call('view',self.path,'html','-o',self.path.with_suffix('.html'));print(self.path)
 def label(self,txt):self.shape('edition',2.82,1.7,27,1,text=txt,size=12,spacing=2,color='56766C')

def foundations(cli,out):
 d=Deck(cli,out/'foundations.pptx');d.start()
 d.call('set',d.path,'/theme','--prop','name=Forest Study','--prop','accent1=174C43','--prop','accent2=80A98D','--prop','accent3=CDDECC','--prop','headingFont=Arial','--prop','bodyFont=Calibri')
 d.new();d.label('FIELD NOTES / 01')
 # Grid: x=2 columns, y=2 rows; no author photo assets are copied.
 d.shape('hero',W/6,H/4,14,5.2,text='A quieter\nway to move.',size=48,bold=True,lineSpacing='.9x')
 d.shape('body',W/6,H/4+5.7,11.5,2.1,text='Start with a clear direction.\nMake room for what matters.',size=20,font='Calibri',lineSpacing='1.15x')
 d.shape('cta',W/6,H/4+8.7,6.0,1.35,fill='174C43',preset='roundRect',text='EXPLORE',size=14,color='FFFFFF',spacing=2,align='center',valign='center')
 d.shape('device',22.0,H/8,8.0,H*6/8,fill='174C43',preset='roundRect')
 d.shape('screen',22.45,H/8+.5,7.1,H*6/8-1,fill='CDDECC',preset='roundRect')
 d.shape('screen-title',23.05,6.6,5.9,3,text='Take the\nscenic route.',font='Georgia',size=27,lineSpacing='1.1x')
 d.shape('device-dot',25.45,3.2,1.0,.12,fill='174C43',preset='roundRect')
 d.notes('TeZ-y45Ei1c 02:28-05:33: 12-column / 8-row grid, title 2 columns and 2 rows from edges, body and button share left edge, phone has 1 empty row above/below. Grid computed, omitted from final artwork. Original phone screenshot replaced by native device illustration.')
 d.new();d.label('FIELD NOTES / 02');d.shape('title',2.82,3.0,28,2,text='Small steps. Clear direction.',size=37,bold=True)
 d.shape('axis',4,9.5,26,.04,fill='9CAF9D')
 for i,(year,txt) in enumerate(zip(['2026','2027','2028','2029','2030'],['Listen first','Set a course','Build together','Test the path','Move forward'])):
  cx=4.5+i*6.2
  d.shape(f'year-{i}',cx-2,6.9,4,1.2,text=year,size=25,bold=True,align='center',valign='bottom')
  d.shape(f'point-{i}',cx-.7,8.8,1.4,1.4,fill=['174C43','36745D','5D9274','80A98D','A4C4A5'][i],preset='ellipse')
  d.shape(f'copy-{i}',cx-2,11.0,4,2,text=txt,size=18,font='Calibri',align='center')
 d.notes('TeZ 06:14-08:57: year boxes share bottom, copy shares top, each event shares center, evenly distribute complete groups. Native shape coordinates implement the same constraints.')
 d.new();d.label('TYPE STUDY / 03')
 for j,(font,title) in enumerate([('Georgia','An editorial voice.'),('Arial','A modern voice.')]):
  x=2.82+j*15.5;d.shape(f'fontlabel-{j}',x,4,13,1,text=font.upper(),size=12,spacing=2)
  d.shape(f'fonttitle-{j}',x,5.6,13,4,text=title,font=font,size=40,bold=j==1,lineSpacing='.9x')
  d.shape(f'fontbody-{j}',x,11,12,3,text='The same message takes a different tone.\nKeep the reading experience clear.',font='Calibri',size=20,lineSpacing='1.2x')
 d.notes('TeZ 09:19-13:19: category changes mood, refine tracking/leading/weight, inspect full-screen legibility. Local Georgia/Arial substitute for missing Libre Baskerville/Montserrat; source 50/14/10pt is an example, not a universal rule.')
 d.new();d.label('COLOUR STUDY / 04');d.shape('title',2.82,3.4,27,2,text='One idea. A quieter palette.',font='Georgia',size=36)
 for i,(c,t) in enumerate(zip(['174C43','36745D','5D9274','80A98D','CDDECC'],['FOREST','LEAF','SAGE','MOSS','MIST'])):
  x=2.82+i*5.7;d.shape(f'colour-{i}',x,7,5.0,5.3,fill=c,preset='rect');d.shape(f'colourlabel-{i}',x,13,5,1,text=t,size=12,spacing=1.2)
 d.shape('closing',2.82,15.3,26,1.2,text='A shared hue family gives the page a consistent mood.',size=21,font='Calibri')
 d.notes('TeZ 14:31-18:12: preserve brand constraints; monochrome shares hue, varies saturation/brightness; save selected hex codes into theme accent slots. This green palette is an implementation choice, not an exact screenshot palette.')
 d.finish()

def pairing(cli,out):
 d=Deck(cli,out/'font-pairing.pptx');d.start()
 for i,(head,body) in enumerate([('Georgia','Calibri'),('Georgia','Arial'),('Times New Roman','Arial')]):
  d.new('FFFFFF');d.label(f'TYPE PAIR / 0{i+1}');d.shape('rule',2.82,3.2,1.2,.1,fill='174C43')
  d.shape('title',2.82,5.0,27,2.5,text='The shape of a good idea',font=head,size=42,bold=True)
  d.shape('subtitle',2.82,8,25,1.5,text='Make space for the story.',font=head,size=25,italic=True)
  d.shape('body',2.82,11.0,24,4,text='Good typography begins with a clear role for every line.\nA confident heading invites attention. A quieter body face\nkeeps the story easy to follow, one thought at a time.',font=body,size=22,lineSpacing='1.2x')
  d.shape('credit',2.82,17,25,1,text=f'{head} + {body}',font='Calibri',size=12,spacing=1,color='56766C')
  d.notes('fCHriW2sUJ4 06:20-09:58: form/contrast matrix is a shortlist heuristic, not a rule; decide title/body roles and refine by trial. Pair examples 10:19-10:54. Sizes adapted for projected reading; not copying arbitrary source small-print sizes.')
 d.finish()

def fontbasics(cli,out):
 d=Deck(cli,out/'font-hierarchy.pptx');d.start();d.new('FFFFFF');d.label('A SIMPLE INTRODUCTION')
 d.shape('heading',6.0,5.6,22,2.3,text='Who we are',font='Georgia',size=38)
 d.shape('body',6.0,9.2,21.5,5.7,text='We are a small team with a shared sense of direction.\nOur work brings people, places and ideas together.\n\nWe believe good decisions begin with a clear story\nand enough room to think.',font='Garamond',size=23,lineSpacing='1.15x')
 d.notes('XzIR8ep4vw0 01:30-03:11: make heading noticeably larger; change heading to Georgia and body to Garamond; keep small text readable. Both fonts are installed locally. Artwork is original and does not copy the paid Brusher template.')
 d.finish()

def typography(cli,out):
 d=Deck(cli,out/'typography-before.pptx');d.start();d.new('F6EDE8')
 d.shape('word-seed',4,5.3,26,8,text='Parallax',font='Script MT Bold',size=150,color='171620',valign='center')
 d.flush();d.call('set',d.path,"/slide[1]/shape[@name='word-seed']",'--prop','textWarp=textSlantUp')
 # 1.0.160 appends warp after noAutofit; fix CT_TextBodyProperties child order locally.
 d.call('raw-set',d.path,'/slide[1]','--xpath','//a:bodyPr/a:noAutofit','--action','remove')
 d.call('raw-set',d.path,'/slide[1]','--xpath','//a:bodyPr','--action','append','--xml','<a:noAutofit xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')
 d.notes('ZJB3trPn9_w 01:06-02:42: Milkshake 160pt then Stand Up warp. Milkshake not installed: Script MT Bold 150pt is an explicit substitute. OfficeCLI textSlantUp represents a slanted baseline; verify in native PowerPoint before boolean conversion. COM polish converts text to native outlines and builds subtract highlights / clipped stripe shadow.')
 d.finish()

if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',type=Path,required=True);p.add_argument('--mode',choices=['all','foundations','pairing','hierarchy','typography'],default='all');a=p.parse_args();a.outdir.mkdir(parents=True,exist_ok=True)
 jobs={'foundations':foundations,'pairing':pairing,'hierarchy':fontbasics,'typography':typography}
 for mode,f in jobs.items():
  if a.mode in ('all',mode):f(a.officecli,a.outdir)
