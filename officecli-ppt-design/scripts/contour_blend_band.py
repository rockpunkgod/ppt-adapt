"""Create filled outline bands with fixed Shapely2.1.2; input caller glyph contours."""
import sys,json,argparse,hashlib
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--input',required=True);p.add_argument('--output',required=True);p.add_argument('--geometry-deps',required=True);a=p.parse_args();out=Path(a.output)
if out.exists():raise FileExistsError('Refusing overwrite')
sys.path.insert(0,a.geometry_deps)
import shapely
from shapely.geometry import MultiLineString
from shapely.geometry.polygon import orient
assert shapely.__version__=='2.1.2'
g=json.loads(Path(a.input).read_text(encoding='utf-8-sig'));u=g['unitsToCm'];b=g['bounds'];rings=[[(float(x-b['x'])*u,float(y-b['y'])*u) for x,y in r] for r in g['original']]
closed=[r+[r[0]] for r in rings];band=MultiLineString(closed).buffer(.03,quad_segs=64,join_style='round');assert band.is_valid and band.geom_type=='MultiPolygon'
contours=[]
for poly in band.geoms:
 poly=orient(poly,sign=1)
 for r in [poly.exterior,*poly.interiors]:contours.append([[x/u+b['x'],y/u+b['y']] for x,y in list(r.coords)[:-1]])
assert len(contours)==4
g['strokes']=[{'offsetMm':.3,'contours':contours,'pointCount':sum(map(len,contours))}];g['strokeMethod']='Shapely2.1.2 MultiLineString closed boundary buffer .03cm, quad_segs64 round joins';g['sourceInputSha256']=hashlib.sha256(Path(a.input).read_bytes()).hexdigest();out.write_text(json.dumps(g,indent=2),encoding='utf-8');print(json.dumps({'path':str(out),'rings':len(contours),'points':g['strokes'][0]['pointCount'],'valid':band.is_valid}))
