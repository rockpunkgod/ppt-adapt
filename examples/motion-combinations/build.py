"""Small editable examples composed from existing PPT Adapt primitives."""
import argparse,sys,json,struct,zlib,tempfile,xml.etree.ElementTree as E
from zipfile import ZipFile
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'officecli-ppt-design/scripts'))
from build_components import Deck

def main():
 p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);p.add_argument('--text',default='TRACE');a=p.parse_args()
 d=Deck(a.out,a.officecli);ops=[]
 def sh(s,n,x,y,w,h,t='',size=20,c='16251D',fill='none',**extra):
  props=dict(name=n,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',fill=fill,line='none',font='Noto Sans SC',margin='0cm',size=str(size),color=c,autoFit='none',valign='center')
  if t:props['text']=t
  props.update({k:str(v).lower() if isinstance(v,bool) else str(v) for k,v in extra.items()})
  ops.append(dict(command='add',parent=f'/slide[{s}]',type='shape',props=props))
 def flush():
  if ops:d.batch(ops.copy());ops.clear()
 s=d.slide('16251D')
 sh(s,'label',1.5,1.2,30,1,'TRACE / 让重要信息先被看见',17,'DAEF78')
 sh(s,'slice-source',5,6,24,4.8,a.text,100,'F2F3EC',font='Arial Black',align='center',bold=True)
 sh(s,'caption',5,13,24,2,'切开噪声，让答案浮现。',30,'F2F3EC',align='center')
 sh(s,'hint',1.5,17.6,30,.7,'单击：刀光 → 上片 → 下片；再次单击进入下一页。',12,'BCC7BB')
 flush();d.notes(s,'原创动效样页。切线、上下片及刀光由原生PowerPoint精修，原始文字隐藏保留；修改后从基础稿重建轮廓。')
 s=d.slide('F2F3EC')
 sh(s,'title',1.5,1.8,30,2,'一个对象，一条清楚的动作句。',34)
 sh(s,'sequence',1.5,4,30,1.3,'出现  →  移动与强调  →  离场',21,'637066')
 sh(s,'card-bg',4,8,6,4,fill='DAEF78');sh(s,'card-number',4.6,8.5,4.8,1.5,'01',42)
 sh(s,'card-text',4.6,10.2,4.8,1,'来源已确认',18)
 sh(s,'destination',22,13,7,2,'接回工作流',24,'637066',align='center')
 sh(s,'hint',1.5,17.6,30,.7,'第一次单击：入场＋路径＋缩放＋旋转；第二次单击：退场。',12,'637066')
 flush();d.run('add','/slide[2]','--type','group',props={'name':'working-card','shapes':','.join('/slide[2]/shape[@name='+n+']' for n in ['card-bg','card-number','card-text'])})
 # Keep fixed navigation outside the scrolling scene. The two state slides share identities.
 s=d.slide('F2F3EC')
 for month in range(6):
  x=9.2+month*7.4
  sh(s,f'!!month-{month}',x,6,7.4,1.3,['01 JAN','02 FEB','03 MAR','04 APR','05 MAY','06 JUN'][month],20)
  sh(s,f'!!grid-{month}',x,7.4,.025,7.2,fill='D6DCD0')
 for i,(label,x,y,w) in enumerate([('资料接入',9.4,8.2,10),('试用与反馈',16.8,10.7,11),('团队扩展',31.7,13.2,10)]):
  sh(s,f'!!task-{i}',x,y,w,1.5,label,20,fill='DAEF78',margin='.25cm')
 sh(s,'!!task-link-mark',25,10.95,2.1,1.18125,'↗',25,align='center')
 # The mask covers the scrolling left edge while labels remain above it.
 sh(s,'fixed-mask',0,5.7,8.95,9.6,fill='F2F3EC')
 sh(s,'fixed-title',1.5,2.7,29,1.7,'循迹 / 半年试用路线',34)
 sh(s,'fixed-label',1.5,6.4,6.7,1,'工作阶段',20)
 for i,t in enumerate(['建立资料库','验证真实任务','扩大使用范围']):sh(s,f'row-{i}',1.5,8.3+i*2.5,7,.9,t,17,'637066')
 for q,x in [(1,23.5),(2,27.5)]:sh(s,f'nav-q{q}',x,1.4,3.2,1.1,f'Q{q}',17,fill='DAEF78',align='center')
 sh(s,'hint',1.5,17.5,30,1,'点击 Q1 / Q2 滚动；点击任务中的预览进入详情，再单击返回。',13,'637066')
 flush()
 s=d.clone(3);d.run('set','/slide[4]',props={'transition':'morph-2000','advanceClick':False})
 for m in range(6):
  for kind in ['month','grid']:d.set(4,f'!!{kind}-{m}',x=f'{9.2+m*7.4-22.2}cm')
 for i,x in enumerate([9.4,16.8,31.7]):d.set(4,f'!!task-{i}',x=f'{x-22.2}cm')
 d.set(4,'!!task-link-mark',x='2.8cm')
 d.run('set','/slide[3]',props={'transition':'morph-2000','advanceClick':False})
 s=d.slide('16251D');sh(s,'label',1.5,1.2,30,1,'TRACE / 试用与反馈',17,'DAEF78')
 sh(s,'title',1.5,4,30,3,'用一项真实任务，\n判断答案是否有用。',40,'F2F3EC')
 for i,(n,t) in enumerate([('01','问题能否被找到'),('02','来源能否被核对'),('03','行动能否被接住')]):
  sh(s,f'number-{i}',2+i*10.5,10.5,8,2,n,50,'DAEF78');sh(s,f'body-{i}',2+i*10.5,13.2,9,1.5,t,23,'F2F3EC')
 sh(s,'hint',1.5,17.5,30,1,'单击返回进入时的路线页。',13,'BCC7BB');flush()
 for s,x in [(3,25),(4,2.8)]:
  d.run('add',f'/slide[{s}]','--type','zoom',props={'name':'!!task-detail','target':5,'x':f'{x}cm','y':'10.95cm','width':'2.1cm','height':'1.18125cm','returnToParent':True,'transitionDur':700})
 # Reuse the established carrier/relationship method. A custom transparent cover
 # keeps the visible regular-shape marker continuous while the native Zoom owns clicks.
 ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','mc':'http://schemas.openxmlformats.org/markup-compatibility/2006','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','pslz':'http://schemas.microsoft.com/office/powerpoint/2016/slidezoom','p166':'http://schemas.microsoft.com/office/powerpoint/2016/6/main'}
 for k,v in ns.items():E.register_namespace(k,v)
 def chunk(tag,data):return struct.pack('>I',len(data))+tag+data+struct.pack('>I',zlib.crc32(tag+data)&0xffffffff)
 with tempfile.TemporaryDirectory(prefix='ppt-transparent-cover-') as td:
  transparent=Path(td)/'transparent.png'
  transparent.write_bytes(b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',1,1,8,6,0,0,0))+chunk(b'IDAT',zlib.compress(b'\x00\x00\x00\x00\x00'))+chunk(b'IEND',b''))
  for s in [3,4]:
   parent=f'/slide[{s}]';carrier='transparent-zoom-cover'
   d.run('add',parent,'--type','picture',props={'name':carrier,'src':str(transparent),'x':'0cm','y':'0cm','width':'1cm','height':'1cm'})
   node=d.run('get',parent+'/picture[@name='+carrier+']')['data']['results'][0];rid=node['format']['relId']
   d.run('raw-set',parent,'--xpath',f"//p:pic/p:nvPicPr/p:cNvPr[@name='{carrier}']",'--action','setattr','--xml','hidden=1')
   d.run('close')
   with ZipFile(d.file) as z:root=E.fromstring(z.read(f'ppt/slides/slide{s}.xml'))
   block=next(n for n in root.findall('.//mc:AlternateContent',ns) if n.find(".//p:cNvPr[@name='!!task-detail']",ns) is not None)
   block.find('.//pslz:zmPr',ns).set('imageType','cover')
   for blip in block.findall('.//a:blip',ns):blip.set('{'+ns['r']+'}embed',rid)
   for line in block.findall('.//a:ln',ns):line.clear();E.SubElement(line,'{'+ns['a']+'}noFill')
   d.run('raw-set',parent,'--xpath',"//mc:AlternateContent[mc:Choice/p:graphicFrame/p:nvGraphicFramePr/p:cNvPr[@name='!!task-detail']]",'--action','replace','--xml',E.tostring(block,encoding='unicode'))
 d.run('close');d.run('validate');print(d.file)
if __name__=='__main__':main()
