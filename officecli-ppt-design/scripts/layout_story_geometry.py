"""OfficeCLI foundation for source merge-shapes and 2-hour design lesson."""
import argparse,math,struct,zlib
from pathlib import Path
from layout_story_build import Deck,TEAL,INK,GOLD,PAPER
def pattern(p):
 w,h=1200,800;raw=bytearray()
 for y in range(h):
  raw.append(0)
  for x in range(w):
   color=(20+int(50*y/h),100+int(90*x/w),120+int(60*y/h))
   if (x-860)**2+(y-210)**2<160**2:color=(230,180,85)
   if y>500+70*math.sin(x/160):color=(25,65,62)
   raw.extend(color)
 def c(t,d):return struct.pack('>I',len(d))+t+d+struct.pack('>I',zlib.crc32(t+d)&0xffffffff)
 p.write_bytes(b'\x89PNG\r\n\x1a\n'+c(b'IHDR',struct.pack('>IIBBBBB',w,h,8,2,0,0,0))+c(b'IDAT',zlib.compress(raw))+c(b'IEND',b''))
def main(exe,out):
 if out.exists():raise FileExistsError('Output directory already exists; choose a new --outdir')
 out.mkdir(parents=True);path=out/'geometry-base.pptx'
 if path.exists():raise FileExistsError('Refusing to overwrite '+str(path))
 art=out/'original-pattern.png';pattern(art);d=Deck(exe,path)
 s=d.slide('把多个形状合成一个轮廓','01 / Union → 图片与轮廓Intersect')
 d.batch([d.add('picture',s,name='source-image',src=str(art),x='11cm',y='5cm',width='12cm',height='11cm'),d.shape(s,'ball-left',12,6.5,4.5,4.5,fill=GOLD,preset='ellipse'),d.shape(s,'ball-right',16,6.5,4.5,4.5,fill=GOLD,preset='ellipse'),d.shape(s,'ball-top',14,5.3,4.5,4.5,fill=GOLD,preset='ellipse'),d.shape(s,'cone',13,9.4,7,6,fill=GOLD,preset='triangle',rotation=180)])
 d.note(s,'Leila Merge02:54–05:19: three circles+edited triangle→Union; picture selected first, custom silhouette second→Intersect; Crop changes image viewport. Original pattern substitutes photo; standard triangle instead of edited points.')
 for mode,title in [('Intersect','图片进入字形'),('Subtract','字形从图片中挖空'),('Fragment','把图片与字形分解成独立片段')]:
  s=d.slide(title,'02 / '+mode+' 保留源素材副本');d.batch([d.add('picture',s,name='source-image',src=str(art),x='3cm',y='6cm',width='28cm',height='8cm'),d.shape(s,'mask-text',3.5,7,27,5,'ICE CREAM',90,fill='none',font='Arial Black',bold=True,align='center',valign='center')]);d.note(s,'Leila Merge05:20–09:07: photo first, bold large text second; Intersect/Subtract/Fragment. Source90pt Rockwell; adapted Arial Black90pt. Merge destroys editable text: retain OfficeCLI base deck, final geometric result is not editable words.')
 s=d.slide('在母版布局中保留真正的图片占位框','03 / 占位框与直接图片是两种对象');d.batch([d.shape(s,'placeholder-note',2,5,11,8,'布局定义裁切框，\n内容页替换图片。\n\n保留原生placeholder\n才有点击换图入口。',24,color=TEAL)]);d.note(s,'Long12:15–15:07: SlideMaster duplicate layout, add picture placeholder, name layout, apply to slide, replace image. This slide receives a COM-authored CustomLayout picture placeholder; do not call ordinary filled rectangle a placeholder.')
 s=d.slide('照片按一个方向依次旋入','04 / Pinwheel 一秒 + AfterPrevious')
 for i,(x,y,rot)in enumerate([(3.5,5.8,-12),(10.8,5.5,10),(18,7.5,-6),(24,6,15)]):d.batch([d.shape(s,f'photo-{i}',x,y,6,8,image=str(art),line='FFFFFF:10',rotation=rot)])
 d.note(s,'Long1:10:49–1:18:28: picture placeholders have white10pt border, centered blur20pt/80% transparency shadow; fill automatically rotated. Pinwheel1sec, firstclick then remainingAfterPrevious. Adapted image-filled native shapes; COM uses native Pinwheel; this autoplay verification uses firstWithPrevious rather than click.')
 s=d.slide('用两个端点与控制柄建立平滑连接','06 / 原生Bézier曲线');ops=[]
 for i,(x,y,t)in enumerate([(3,6,'需求'),(27,12,'服务')]):ops.append(d.shape(s,'node-'+str(i),x,y,4,4,t,22,fill=TEAL,color='FFFFFF',preset='ellipse',align='center',valign='center'))
 ops.append(d.shape(s,'native-curve',7,8,20,6,fill='none',line=TEAL+':9',lineCap='round'));d.batch(ops)
 geom='<a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="100000" h="100000" fill="none"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:cubicBezTo><a:pt x="50000" y="0"/><a:pt x="50000" y="100000"/><a:pt x="100000" y="100000"/></a:cubicBezTo></a:path></a:pathLst></a:custGeom>'
 d.call('raw-set',d.path,s,'--xpath','//p:sp[p:nvSpPr/p:cNvPr[@name="native-curve"]]/p:spPr/a:prstGeom','--action','replace','--xml',geom);d.note(s,'Long1:33:33–1:40:27: freeform2points/editpoint handles smoother than manyclick Curve, sendbehind nodes; source four-stop line gradient and overlap shadows. Here tests simple editable cubic, not four-stop gradient/shadow or chart quantities.')
 # Preserve source aspect ratio before native boolean image clipping.
 d.call('set',d.path,'/slide[1]/picture[@name=source-image]','--prop','crop=13.63636,0,13.63636,0')
 for si in [2,3,4]:d.call('set',d.path,f'/slide[{si}]/picture[@name=source-image]','--prop','crop=0,28.57143,0,28.57143')
 d.finish();print(path)
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--officecli',default='officecli');p.add_argument('--outdir',type=Path,required=True);a=p.parse_args();main(a.officecli,a.outdir)
