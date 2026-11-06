"""Apply only our native-UI-generated Replay timing to our own CLI/InkML page."""
import argparse,zipfile
from pathlib import Path
from xml.etree import ElementTree as E
p=argparse.ArgumentParser();p.add_argument('--base',type=Path,required=True);p.add_argument('--timing',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
if a.output.exists():raise FileExistsError(a.output)
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','p14':'http://schemas.microsoft.com/office/powerpoint/2010/main','mc':'http://schemas.openxmlformats.org/markup-compatibility/2006'}
for key,value in n.items():E.register_namespace(key,value)
with zipfile.ZipFile(a.base) as z:files={x.filename:z.read(x.filename) for x in z.infolist()}
slide=E.fromstring(files['ppt/slides/slide1.xml']);ink=next(x for x in slide.findall('.//p14:cNvPr',n) if x.get('name')=='original-native-ink');audio=next(x for x in slide.findall('.//p:cNvPr',n) if x.get('name')=='audio')
timing=E.parse(a.timing).getroot();assert timing.tag=='{'+n['p']+'}timing'
for x in timing.findall('.//p:spTgt',n):
 if x.get('spid')=='200000':x.set('spid',ink.get('id'))
 elif x.get('spid')=='100006':x.set('spid',audio.get('id'))
 else:raise ValueError('Unexpected native seed target')
for old in list(slide.findall('p:timing',n)):slide.remove(old)
slide.append(timing);files['ppt/slides/slide1.xml']=E.tostring(slide,encoding='utf-8',xml_declaration=True)
with zipfile.ZipFile(a.output,'w',zipfile.ZIP_DEFLATED) as z:
 for name,data in files.items():z.writestr(name,data)
print(a.output)
