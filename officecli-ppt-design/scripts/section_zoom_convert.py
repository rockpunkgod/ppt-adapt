"""Use the Microsoft sectionzoom schema; every mutation via OfficeCLI raw-set."""
import argparse,subprocess,json,zipfile,shutil,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--input',required=True);p.add_argument('--output',required=True);p.add_argument('--evidence',required=True);a=p.parse_args();f=Path(a.output);ev=Path(a.evidence)
if f.exists():p.error('Refusing existing output')
shutil.copyfile(a.input,f)
def call(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace')
 if r.returncode or 'UNSUPPORTED' in r.stdout+r.stderr:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','mc':'http://schemas.openxmlformats.org/markup-compatibility/2006','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','psez':'http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom','pslz':'http://schemas.microsoft.com/office/powerpoint/2016/slidezoom','p166':'http://schemas.microsoft.com/office/powerpoint/2016/6/main'}
for k,v in ns.items():E.register_namespace(k,v)
sections=json.loads((ev/'sections.json').read_text(encoding='utf-8-sig'))
for i,target in [(1,2),(2,4)]:
 name=f'section-poster-{i}';call('add',f,'/slide[1]','--type','picture','--prop',f'name={name}','--prop',f'src={ev/f"entry-{target}.png"}','--prop','x=0cm','--prop','y=0cm','--prop','width=12cm','--prop','height=6.75cm')
 d=json.loads(call('get',f,f'/slide[1]/picture[@name={name}]','--json'))['data']['results'][0];rid=d['format']['relId']
 call('raw-set',f,'/slide[1]','--xpath',f"//p:pic/p:nvPicPr/p:cNvPr[@name='{name}']",'--action','setattr','--xml','hidden=1')
 with zipfile.ZipFile(f) as z:root=E.fromstring(z.read('ppt/slides/slide1.xml'))
 block=next(x for x in root.findall('.//mc:AlternateContent',ns) if x.find(f".//p:cNvPr[@name='section-card-{i}']",ns) is not None)
 choice=block.find('mc:Choice',ns);choice.set('Requires','psez')
 gd=choice.find('.//a:graphicData',ns);gd.set('uri',ns['psez'])
 wrap=gd.find('pslz:sldZm',ns);wrap.tag='{'+ns['psez']+'}sectionZm'
 obj=wrap.find('pslz:sldZmObj',ns);obj.tag='{'+ns['psez']+'}sectionZmObj';obj.attrib.clear();obj.set('sectionId',sections[i]['id'])
 props=obj.find('pslz:zmPr',ns);props.tag='{'+ns['psez']+'}zmPr';props.set('showBg','0');props.set('returnToParent','1')
 for b in block.findall('.//a:blip',ns):b.set('{'+ns['r']+'}embed',rid)
 for sp in block.findall('.//p166:spPr',ns)+block.findall('./mc:Fallback/p:pic/p:spPr',ns):
  ln=sp.find('a:ln',ns)
  if ln is None:ln=E.SubElement(sp,'{'+ns['a']+'}ln')
  ln.clear();E.SubElement(ln,'{'+ns['a']+'}noFill')
 # Leave Fallback as an image/slide link for legacy consumers; Choice is native Section Zoom.
 call('raw-set',f,'/slide[1]','--xpath',f"//mc:AlternateContent[mc:Choice/p:graphicFrame/p:nvGraphicFramePr/p:cNvPr[@name='section-card-{i}']]",'--action','replace','--xml',E.tostring(block,encoding='unicode'))
call('close',f);print(call('validate',f));print(call('view',f,'issues','--json'))
