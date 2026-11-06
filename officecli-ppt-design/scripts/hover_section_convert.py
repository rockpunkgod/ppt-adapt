"""Adapt existing Hover Zoom frames to existing Section Zoom schema in place."""
import argparse,json,shutil,subprocess,zipfile
from pathlib import Path
import xml.etree.ElementTree as E
p=argparse.ArgumentParser(description=__doc__)
for key in ['officecli','input','output','sections','posters']:p.add_argument('--'+key,required=True)
a=p.parse_args();out=Path(a.output)
if out.exists():raise FileExistsError(out)
sections=json.loads(Path(a.sections).read_text(encoding='utf-8-sig'))
if [(s['first'],s['count']) for s in sections]!=[(1,3),(4,2),(6,2)]:raise ValueError('Expected menu + two two-page sections')
shutil.copyfile(a.input,out)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','mc':'http://schemas.openxmlformats.org/markup-compatibility/2006','a':'http://schemas.openxmlformats.org/drawingml/2006/main','pslz':'http://schemas.microsoft.com/office/powerpoint/2016/slidezoom','psez':'http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom'}
for k,v in ns.items():E.register_namespace(k,v)
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
for slide in (1,2,3):
 for branch in (1,2):
  name=f'!!zoom-{branch}'
  poster=Path(a.posters)/f'entry-{sections[branch]["first"]}.png'
  cache=f'section-poster-{branch}'
  cli('add',out,f'/slide[{slide}]','--type','picture','--prop','name='+cache,'--prop','src='+str(poster),'--prop','x=0cm','--prop','y=0cm','--prop','width=10.3cm','--prop','height=5.79375cm')
  rid=json.loads(cli('get',out,f'/slide[{slide}]/picture[@name={cache}]','--json'))['data']['results'][0]['format']['relId']
  cli('raw-set',out,f'/slide[{slide}]','--xpath',f"//p:pic/p:nvPicPr/p:cNvPr[@name='{cache}']",'--action','setattr','--xml','hidden=1')
  with zipfile.ZipFile(out) as z:r=E.fromstring(z.read(f'ppt/slides/slide{slide}.xml'))
  blocks=[x for x in r.findall('.//mc:AlternateContent',ns) if x.find(f".//p:cNvPr[@name='{name}']",ns) is not None]
  if len(blocks)!=1:raise ValueError('Expected one existing native Zoom frame')
  block=blocks[0];choice=block.find('mc:Choice',ns);choice.set('Requires','psez')
  gd=choice.find('.//a:graphicData',ns);gd.set('uri',ns['psez'])
  wrap=gd.find('pslz:sldZm',ns);obj=wrap.find('pslz:sldZmObj',ns);props=obj.find('pslz:zmPr',ns)
  wrap.tag='{'+ns['psez']+'}sectionZm';obj.tag='{'+ns['psez']+'}sectionZmObj';obj.attrib.clear();obj.set('sectionId',sections[branch]['id'])
  props.tag='{'+ns['psez']+'}zmPr';props.set('returnToParent','1');props.set('transitionDur','500')
  for blip in block.findall('.//a:blip',ns):blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed',rid)
  # Replace at the original stack position: the hole mask must stay above Zoom.
  cli('raw-set',out,f'/slide[{slide}]','--xpath',f"//mc:AlternateContent[mc:Choice/p:graphicFrame/p:nvGraphicFramePr/p:cNvPr[@name='{name}']]",'--action','replace','--xml',E.tostring(block,encoding='unicode'))
cli('close',out);print(cli('validate',out))
