import json,subprocess,argparse
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args();cli=a.officecli;o=Path(a.out);o.mkdir(parents=True);f=o/'summary-base.pptx'
def call(*args):
 r=subprocess.run([cli,*map(str,args)],capture_output=True,text=True,encoding='utf-8');assert r.returncode==0,r.stdout+r.stderr;return r.stdout
call('create',f);call('set',f,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm');ops=[]
for chapter,(title,color) in enumerate([('FIELD NOTES','1D7774'),('SIGNAL LAB','D36D55'),('NEXT ROUTE','5A64B5')],1):
 for page in (1,2):
  n=(chapter-1)*2+page;ops.append(dict(command='add',parent='/',type='slide',props={'layout':'blank','background':'F4F1EA'}))
  def shp(name,text,x,y,w,h,size=24,fill='none',fg='17324E'):
   ops.append(dict(command='add',parent=f'/slide[{n}]',type='shape',props={'name':name,'text':text,'x':f'{x}cm','y':f'{y}cm','width':f'{w}cm','height':f'{h}cm','font':'Arial','size':str(size),'color':fg,'fill':fill,'line':'none','margin':'0cm'}))
  shp('chapter-block','',0,0,10.4,19.05,fill=color);shp('chapter-index',f'{chapter:02}',1.8,5,7.5,5,100,fg='FFFFFF');shp('chapter-title',title,12.2,4.5,19,2,34)
  shp('page-identity',f'CHAPTER {chapter} / PAGE {page}',12.3,7.2,18,1,17)
  shp('body','Start with a clear observation.' if page==1 else 'Show the evidence. Move forward.',12.3,9.6,18,3,27)
  shp('footer','NOVA / NATIVE SUMMARY STUDY',12.3,16.7,19,1,12)
p=o/'foundation.json';p.write_text(json.dumps(ops));d=json.loads(call('batch',f,'--input',p,'--json'));assert d['success'];call('close',f);(o/'base-validation.json').write_text(call('validate',f,'--json'));print(f)
