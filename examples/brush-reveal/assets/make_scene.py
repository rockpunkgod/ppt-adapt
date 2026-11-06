"""Original bright RGB diagram, chosen to retain tones after brightness -54%."""
import argparse
from pathlib import Path
from PIL import Image,ImageDraw
p=argparse.ArgumentParser();p.add_argument('--out',type=Path,required=True);a=p.parse_args()
if a.out.exists():raise FileExistsError(a.out)
im=Image.new('RGB',(1600,900),(218,228,216));d=ImageDraw.Draw(im)
for y in range(900):d.line((0,y,1600,y),fill=(218-int(25*y/900),228-int(18*y/900),216-int(12*y/900)))
for i,x in enumerate([360,800,1240]):
 d.line((x,240,x,680),fill=(154,182,167),width=2)
 d.ellipse((x-115,285,x+115,515),fill=[(53,135,129),(84,157,145),(119,178,151)][i],outline=(239,246,223),width=8)
 d.ellipse((x-75,325,x+75,475),outline=(225,240,208),width=5)
 d.rectangle((x-15,360,x+15,440),fill=(237,204,100));d.rectangle((x-40,385,x+40,415),fill=(237,204,100))
d.line((475,400,685,400),fill=(82,137,129),width=5);d.line((915,400,1125,400),fill=(82,137,129),width=5)
a.out.parent.mkdir(parents=True,exist_ok=True);im.save(a.out)
