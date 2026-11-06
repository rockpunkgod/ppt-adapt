"""Original closed stroke contours and actual convex polygon intersections."""
STROKES={
 'left':[(270,250),(310,295),(110,500),(55,455)],
 'right':[(270,300),(325,260),(550,450),(500,510)],
 'vertical':[(270,60),(330,60),(330,550),(270,550)],
 'horizontal':[(80,170),(520,170),(520,230),(80,230)]}
CURVED_LEFT=[('M',(270,250)),('L',(310,295)),('C',(250,380,170,475,110,500)),('L',(55,455)),('C',(145,410,220,325,270,250)),('Z',())]
def flatten(commands,steps=80):
 out=[];current=None
 for kind,coords in commands:
  if kind in ['M','L']:current=coords;out.append(current)
  elif kind=='C':
   p0=current;p1=coords[:2];p2=coords[2:4];p3=coords[4:]
   for i in range(1,steps+1):
    t=i/steps;u=1-t;out.append(tuple(u**3*p0[j]+3*u*u*t*p1[j]+3*u*t*t*p2[j]+t**3*p3[j] for j in [0,1]))
   current=p3
 return out
STROKES['left']=flatten(CURVED_LEFT)
def area(poly):return sum(x1*y2-x2*y1 for (x1,y1),(x2,y2) in zip(poly,poly[1:]+poly[:1]))/2
def intersect(subject,clip):
 """Sutherland–Hodgman: convex closed clipping polygon, retain fractional vertices."""
 out=[tuple(map(float,p)) for p in subject];direction=1 if area(clip)>0 else -1
 for a,b in zip(clip,clip[1:]+clip[:1]):
  src=out;out=[]
  if not src:break
  cross=lambda p:direction*((b[0]-a[0])*(p[1]-a[1])-(b[1]-a[1])*(p[0]-a[0]))
  prev=src[-1];dp=cross(prev)
  for cur in src:
   dc=cross(cur)
   if (dc>=0)!=(dp>=0):
    t=dp/(dp-dc);out.append((prev[0]+t*(cur[0]-prev[0]),prev[1]+t*(cur[1]-prev[1])))
   if dc>=0:out.append(cur)
   prev,dp=cur,dc
 return out
LOGICAL_BRIDGE=intersect(STROKES['right'],STROKES['vertical'])
BRIDGE=intersect(STROKES['right'],[(268,60),(332,60),(332,550),(268,550)])
PATCHES={
 'patch-vh':intersect(STROKES['vertical'],[(0,230),(600,230),(600,300),(0,300)]),
 'patch-vr':intersect(STROKES['vertical'],[(270,210),(330,166.3636363636),(330,256.3636363636),(270,300)]),
 'patch-lv':intersect(STROKES['left'],[(170,0),(270,0),(270,600),(170,600)])}
PATCH_RECEIVERS={'patch-vh':'vertical','patch-vr':'vertical','patch-lv':'left'}
def inside(point,poly):
 x,y=point;c=False
 for (ax,ay),(bx,by) in zip(poly,poly[1:]+poly[:1]):
  if (ay>y)!=(by>y) and x<(bx-ax)*(y-ay)/(by-ay)+ax:c=not c
 return c
