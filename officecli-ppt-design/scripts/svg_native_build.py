"""OfficeCLI foundation with a caller SVG; conversion happens through PowerPoint."""
import argparse,json,subprocess,zipfile,hashlib,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--svg',required=True);p.add_argument('--out',required=True);a=p.parse_args();out=Path(a.out);svg=Path(a.svg).resolve()
if any(q.exists() for q in [out,*[out.with_suffix(t) for t in ('.batch.json','.build-log.json','.validation.json','.source.json')]]):raise FileExistsError('Refusing overwrite')
root=E.fromstring(svg.read_bytes());assert root.tag=='{http://www.w3.org/2000/svg}svg'
assert all(q.tag=='{http://www.w3.org/2000/svg}path' for q in root),'Tested subset: direct solid path children only'
assert not any(k.split('}')[-1] in ('href','style') for q in root.iter() for k in q.attrib),'No external references or CSS in tested subset'
assert len(root)==2 and all(q.get('fill','').startswith('#') and len(q.get('fill',''))==7 for q in root)
assert len(set(q.get('fill').upper() for q in root))==2
for q in root:
 int(q.get('fill')[1:],16)
 assert set(q.attrib)<=set(('fill','fill-rule','d')),'Tested paths contain only geometry and solid fill'
vb=[float(v) for v in root.get('viewBox','').replace(',',' ').split()];assert len(vb)==4 and vb[2]>0 and vb[3]>0
w=min(20,12*vb[2]/vb[3]);h=w*vb[3]/vb[2];x=(33.87-w)/2;y=4.8
out.parent.mkdir(parents=True,exist_ok=True)
def run(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
run('create',out);run('set',out,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
ops=[dict(command='add',parent='/',type='slide',props={'layout':'blank','background':'F6F8FC'})]
def shape(name,text,x,y,w,h,size=16,fill='none',color='17324E'):
 ops.append(dict(command='add',parent='/slide[1]',type='shape',props=dict(name=name,text=text,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',font='Arial',size=str(size),fill=fill,color=color,line='none',margin='0cm')))
shape('title','One vector. Editable pieces.',1.7,1.3,30,1.6,30)
shape('subtitle','The opening reveals both background bands.',1.7,3.1,30,1,16,color='527086')
shape('underlay-top','',x,y,w,h/2,fill='D9E4E6');shape('underlay-bottom','',x,y+h/2,w,h/2,fill='EFDCC7')
ops.append(dict(command='add',parent='/slide[1]',type='picture',props=dict(name='svg-target',src=str(svg),x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm')))
shape('footer','NOVA / VECTOR STUDY',1.7,17.7,30,0.8,12,color='527086')
ops.append(dict(command='add',parent='/slide[1]',type='notes',props={'text':'Masterclass9Ewb0Rd3P3E35:30-36:09 SVG import and Convert to Shape. Original caller SVG, no text/filter/external resources. Keep base/source. Native conversion must produce actual editable custom geometry and preserve holes/colors.'}))
f=out.with_suffix('.batch.json');f.write_text(json.dumps(ops),encoding='utf-8');r=run('batch',out,'--input',f,'--json');out.with_suffix('.build-log.json').write_text(r,encoding='utf-8');assert json.loads(r)['success'];run('close',out)
with zipfile.ZipFile(out) as z:assert any(n.endswith('.svg') for n in z.namelist()),'Original SVG part must remain embedded'
out.with_suffix('.source.json').write_text(json.dumps({'source':str(svg),'sourceSha256':hashlib.sha256(svg.read_bytes()).hexdigest(),'viewBox':vb,'pathCount':len(root),'colors':[q.get('fill').upper() for q in root],'imageCm':[x,y,w,h]},indent=2),encoding='utf-8')
out.with_suffix('.validation.json').write_text(run('validate',out,'--json'),encoding='utf-8');print(out)
