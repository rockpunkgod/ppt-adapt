"""Resolve true rounded polygon offsets instead of self-intersecting widened paths."""
import argparse,json,sys
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--source',type=Path,required=True);p.add_argument('--out',type=Path,required=True);p.add_argument('--engine-path',type=Path);a=p.parse_args()
if a.out.exists():raise FileExistsError(a.out)
if a.engine_path:sys.path.insert(0,str(a.engine_path.resolve()))
import shapely
from shapely.geometry import Polygon
from shapely.geometry.polygon import orient
data=json.loads(a.source.read_text(encoding='utf-8-sig'))
if data.get('text')!='6' or data.get('font')!='Arial Black' or abs(data['unitsToCm']*data['bounds']['height']-11)>1e-6:
 raise ValueError('Validated fixture only: Arial Black 6 at 11 cm')
original=Polygon(data['original'][0],data['original'][1:])
assert original.is_valid,shapely.is_valid_reason(original)
def contours(g):
 assert g.is_valid and not g.is_empty
 polygons=[g] if g.geom_type=='Polygon' else list(g.geoms)
 result=[]
 for poly in polygons:
  poly=orient(poly,sign=1)
  result.append([list(p) for p in list(poly.exterior.coords)[:-1]])
  result.extend([[list(p) for p in list(r.coords)[:-1]] for r in poly.interiors])
 return result
data['offsets']=[]
for mm in [2,4,6]:
 distance=mm/10/data['unitsToCm'];offset=original.buffer(distance,quad_segs=64,join_style='round')
 assert offset.covers(original)
 data['offsets'].append({'offsetMm':mm,'contours':contours(offset),'area':offset.area,'valid':offset.is_valid})
band=original.boundary.buffer(.2/data['unitsToCm'],quad_segs=64,join_style='round')
data['outlinedStroke']=contours(band)
data['engine']={'shapely':shapely.__version__,'geos':shapely.geos_version_string,'join':'round','quadSegs':64,'originalValid':original.is_valid}
a.out.write_text(json.dumps(data,ensure_ascii=False),encoding='utf-8')
print({'out':str(a.out),'offsetPoints':[[len(c) for c in row['contours']] for row in data['offsets']],'strokePoints':[len(c) for c in data['outlinedStroke']]})
