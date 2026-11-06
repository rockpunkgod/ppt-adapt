"""Independently check native PPT table dimensions, real merges, borders and arithmetic."""
import argparse,json,zipfile,xml.etree.ElementTree as E
from pathlib import Path
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def txt(cell):return ''.join(t.text or '' for t in cell.findall('.//a:t',N))
def verify(file):
    with zipfile.ZipFile(file) as z:
        root=E.fromstring(z.read('ppt/slides/slide1.xml'));tables=root.findall('.//a:tbl',N);assert len(tables)==1
        table=tables[0];rows=table.findall('a:tr',N);grid=table.findall('a:tblGrid/a:gridCol',N)
        assert len(rows)==6 and len(grid)==5
        cells=[r.findall('a:tc',N) for r in rows];assert all(len(row)==5 for row in cells)
        assert [int(c.get('w')) for c in grid]==[2520000,2088000,2088000,2088000,2088000]
        assert all(int(r.get('h'))==576000 for r in rows)
        assert cells[0][0].get('rowSpan')=='2' and cells[1][0].get('vMerge') in ('1','true')
        assert cells[0][1].get('gridSpan')=='4' and all(cells[0][c].get('hMerge') in ('1','true') for c in (2,3,4))
        hidden=sum(c.get('hMerge') in ('1','true') or c.get('vMerge') in ('1','true') for row in cells for c in row);assert hidden==4
        assert txt(cells[0][0])=='服务渠道' and txt(cells[0][1])=='服务笔数（笔）'
        assert [txt(c) for c in cells[1][1:]]==['1月','2月','3月','季度合计']
        assert [txt(cells[r][0]) for r in (2,3,4,5)]==['线上','现场','伙伴','合计']
        monthly=[];quarter=[]
        for r in (2,3,4):
            values=[int(txt(c)) for c in cells[r][1:]];assert all(v>=0 for v in values)
            assert sum(values[:3])==values[3];monthly.append(values[:3]);quarter.append(values[3])
        total=[int(txt(c)) for c in cells[5][1:]]
        assert total[:3]==[sum(row[c] for row in monthly) for c in range(3)]
        assert total[3]==sum(quarter)==sum(total[:3])==700
        border_count=0
        for r in range(6):
            for c in range(5):
                cell=cells[r][c];pr=cell.find('a:tcPr',N)
                for edge in ('L','R','T','B'):
                    line=pr.find('a:ln'+edge,N);assert line is not None
                    expected=None
                    if edge=='B' and (r==1 or (r==0 and c==0)):expected=(19050,'176F73')
                    if edge=='B' and r in (2,3):expected=(6350,'D8E1DC')
                    if edge=='B' and r==4:expected=(19050,'176F73')
                    if r==5 and edge in ('T','B'):expected=(19050,'176F73')
                    if expected:
                        assert int(line.get('w'))==expected[0]
                        assert line.find('a:solidFill/a:srgbClr',N).get('val')==expected[1];border_count+=1
                    else:assert line.find('a:noFill',N) is not None,(r,c,edge,'unexpected border')
                if txt(cell):
                    for run in cell.findall('.//a:rPr',N):
                        assert run.get('sz')=='2200'
                        assert run.find('a:ea',N).get('typeface')=='Microsoft YaHei'
                if r>=2 and c>=1:assert cell.find('a:txBody/a:p/a:pPr',N).get('algn')=='r'
        # Three explanatory text objects only; all data cells live in the real table.
        shapes=root.findall('p:cSld/p:spTree/p:sp',N);assert len(shapes)==3
        texts={s.find('p:nvSpPr/p:cNvPr',N).get('name'):txt(s) for s in shapes}
        assert '700' in texts['table-title'] and '单位：笔' in texts['table-source'] and '原创演示' in texts['table-source']
        return {'passed':True,'file':str(file),'rows':6,'columns':5,'rawCells':30,'visibleCells':26,'horizontalMergeSpan':4,'verticalMergeSpan':2,'monthlyTotals':total[:3],'quarterTotal':700,'explicitCellBorderEdges':border_count,'verticalBorders':0,'dataIsNativeTable':True,'limits':'Structure and arithmetic checked; inspect native PowerPoint merged cell geometry and PNG separately.'}
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('file');p.add_argument('--report',required=True);a=p.parse_args();out=Path(a.report)
    if out.exists():raise FileExistsError('Refusing overwrite')
    result=verify(a.file);out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(result,ensure_ascii=True,indent=2))
