import argparse,io,json,zipfile,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('deck',type=Path);p.add_argument('--report',type=Path,required=True);p.add_argument('--edited',action='store_true');a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
n={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart','a':'http://schemas.openxmlformats.org/drawingml/2006/main','s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(a.deck) as z:
 r=E.fromstring(z.read('ppt/charts/chart1.xml'));ar=r.find('.//c:areaChart',n);li=r.find('.//c:lineChart',n)
 assert ar is not None and li is not None
 def vals(e):return [float(q.text) for q in e.findall('.//c:val//c:pt/c:v',n)]
 av,lv=vals(ar),vals(li);assert av==lv==[24,31 if a.edited else 29,27,35,42,46],(av,lv)
 ax=lambda e:[q.get('val') for q in e.findall('c:axId',n)]
 assert ax(ar)==ax(li)
 va=r.findall('.//c:valAx',n);assert len(va)==1 and va[0].find('c:scaling/c:min',n).get('val')=='0'
 alpha=[int(q.get('val')) for q in ar.findall('.//a:gradFill//a:alpha',n)];assert 0 in alpha and max(alpha)>0
 assert li.find('.//c:smooth',n).get('val')=='0'
 book=next(x for x in z.namelist() if x.startswith('ppt/embeddings/') and x.endswith('.xlsx'))
 assert r.find('c:externalData',n) is not None
 with zipfile.ZipFile(io.BytesIO(z.read(book))) as w:
  sh=E.fromstring(w.read('xl/worksheets/sheet1.xml'));cells={q.get('r'):q.find('s:v',n).text for q in sh.findall('.//s:c',n) if q.find('s:v',n) is not None}
  bv=[float(cells['B'+str(i)]) for i in range(2,8)];dv=[float(cells['D'+str(i)]) for i in range(2,8)];assert bv==dv==av
 assert len(ar.findall('.//c:numRef',n))==1 and len(li.findall('.//c:numRef',n))==1
 result={'passed':True,'values':av,'types':['area','line'],'same_axes':ax(ar),'value_axes':1,'minimum':0,'embedded_workbook':book,'workbook_series_match':True,'alpha_stops':alpha,'smooth':False,'native_edit_probe':a.edited}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
