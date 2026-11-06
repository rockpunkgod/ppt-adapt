"""Original fireworks GIF and black-backed confetti video, as editable-deck inputs."""
from pathlib import Path
import argparse,math,subprocess,random
from PIL import Image,ImageDraw
p=argparse.ArgumentParser();p.add_argument('--ffmpeg',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);frames=[]
for k in range(24):
 im=Image.new('RGBA',(400,300));d=ImageDraw.Draw(im)
 for cx,cy,offset,color in [(190,150,0,'#DAEF78'),(290,90,8,'#F4CE6A'),(90,105,16,'#F2F3EC')]:
  phase=((k+offset)%24)/24;rad=8+phase*65
  for j in range(10):
   t=j*math.tau/10;inner=max(0,rad-9);d.line((cx+inner*math.cos(t),cy+inner*math.sin(t),cx+rad*math.cos(t),cy+rad*math.sin(t)),fill=color,width=3)
 frames.append(im)
frames[0].save(a.outdir/'fireworks.gif',save_all=True,append_images=frames[1:],duration=100,loop=0,disposal=2,transparency=0,optimize=False)
rng=random.Random(419);particles=[]
for i in range(65):particles.append((rng.randrange(960),rng.randrange(540),rng.uniform(35,95),rng.uniform(0,math.tau),rng.choice(['#DAEF78','#F4CE6A','#EF3C85','#58D9CE','#F2F3EC'])))
encoder=subprocess.Popen([a.ffmpeg,'-v','error','-f','rawvideo','-pix_fmt','rgb24','-video_size','960x540','-framerate','30','-i','pipe:0','-c:v','libx264','-preset','fast','-crf','18','-pix_fmt','yuv420p',str(a.outdir/'confetti.mp4')],stdin=subprocess.PIPE)
for frame in range(360):
 t=frame/30;im=Image.new('RGB',(960,540),'black');d=ImageDraw.Draw(im)
 for x,y,speed,angle,color in particles:
  px=(x+30*math.sin(t+angle))%960;py=(y+t*speed)%540;theta=angle+t*1.7;c=math.cos(theta);s=math.sin(theta)
  points=[(px+xx*c-yy*s,py+xx*s+yy*c) for xx,yy in [(-5,-3),(5,-3),(5,3),(-5,3)]];d.polygon(points,fill=color)
 if frame==0:im.save(a.outdir/'poster.png')
 encoder.stdin.write(im.tobytes())
encoder.stdin.close()
if encoder.wait()!=0:raise RuntimeError('Original video encoding failed')
print(a.outdir)
