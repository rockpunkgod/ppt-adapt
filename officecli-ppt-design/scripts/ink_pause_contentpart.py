import argparse,zipfile
from pathlib import Path
from xml.etree import ElementTree as E
p=argparse.ArgumentParser();p.add_argument('--base',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
if a.output.exists():raise FileExistsError(a.output)
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','p14':'http://schemas.microsoft.com/office/powerpoint/2010/main','mc':'http://schemas.openxmlformats.org/markup-compatibility/2006','ink':'http://www.w3.org/2003/InkML','rel':'http://schemas.openxmlformats.org/package/2006/relationships','ct':'http://schemas.openxmlformats.org/package/2006/content-types'}
for prefix in ['p','a','r','p14','mc','ink']:E.register_namespace(prefix,n[prefix])
q=lambda prefix,name:'{'+n[prefix]+'}'+name
def node(parent,prefix,name,attrs=None):return E.SubElement(parent,q(prefix,name),attrs or {})
with zipfile.ZipFile(a.base) as z:files={i.filename:z.read(i.filename) for i in z.infolist()}
slide=E.fromstring(files['ppt/slides/slide1.xml']);tree=slide.find('p:cSld/p:spTree',n)
assert tree is not None and slide.find('.//p:contentPart',n) is None
alt=node(tree,'mc','AlternateContent');choice=node(alt,'mc','Choice',{'Requires':'p14'});part=node(choice,'p','contentPart',{q('r','id'):'rIdOriginalInk'})
nv=node(part,'p14','nvContentPartPr');node(nv,'p14','cNvPr',{'id':'200000','name':'original-native-ink'});node(nv,'p14','cNvContentPartPr',{'isComment':'0'});node(nv,'p14','nvPr')
xfrm=node(part,'p14','xfrm');node(xfrm,'a','off',{'x':str(int(6.75*360000)),'y':str(int(7.75*360000))});node(xfrm,'a','ext',{'cx':str(21*360000),'cy':str(5*360000)})
node(alt,'mc','Fallback')
ink=E.Element(q('ink','ink'));defs=node(ink,'ink','definitions');context=node(defs,'ink','context',{'{http://www.w3.org/XML/1998/namespace}id':'context0'});source=node(context,'ink','inkSource',{'{http://www.w3.org/XML/1998/namespace}id':'source0'});tf=node(source,'ink','traceFormat')
node(tf,'ink','channel',{'name':'X','type':'integer','units':'dev'});node(tf,'ink','channel',{'name':'Y','type':'integer','units':'dev'});node(tf,'ink','channel',{'name':'T','type':'integer'})
brush=node(defs,'ink','brush',{'{http://www.w3.org/XML/1998/namespace}id':'brush0'})
for name,value,unit in [('width','0.08','cm'),('height','0.08','cm'),('color','#FFFFFF',None),('tip','ellipse',None)]:
 attrs={'name':name,'value':value}
 if unit:attrs['units']=unit
 node(brush,'ink','brushProperty',attrs)
lines=[(0,5000,7000,0),(7000,0,14000,5000),(14000,5000,21000,0)]
for j,(x0,y0,x1,y1) in enumerate(lines):
 trace=node(ink,'ink','trace',{'{http://www.w3.org/XML/1998/namespace}id':f'trace{j}','contextRef':'#context0','brushRef':'#brush0'})
 trace.text=', '.join(f'{round(x0+(x1-x0)*i/60)} {round(y0+(y1-y0)*i/60)} {round((j+i/60)*5000/3)}' for i in range(61))
rels=E.fromstring(files['ppt/slides/_rels/slide1.xml.rels']);node(rels,'rel','Relationship',{'Id':'rIdOriginalInk','Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml','Target':'../ink/ink1.xml'})
types=E.fromstring(files['[Content_Types].xml']);node(types,'ct','Override',{'PartName':'/ppt/ink/ink1.xml','ContentType':'application/inkml+xml'})
for name,root in [('ppt/slides/slide1.xml',slide),('ppt/ink/ink1.xml',ink)]:files[name]=E.tostring(root,encoding='utf-8',xml_declaration=True)
E.register_namespace('',n['rel']);files['ppt/slides/_rels/slide1.xml.rels']=E.tostring(rels,encoding='utf-8',xml_declaration=True)
E.register_namespace('',n['ct']);files['[Content_Types].xml']=E.tostring(types,encoding='utf-8',xml_declaration=True)
with zipfile.ZipFile(a.output,'w',zipfile.ZIP_DEFLATED) as z:
 for name,data in files.items():z.writestr(name,data)
print(a.output)
