"""Independent OOXML checks for the packaged native layout mechanisms."""
import argparse,json,zipfile,xml.etree.ElementTree as E,io,struct,posixpath
from pathlib import Path
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def jpeg_size(data):
    f=io.BytesIO(data);assert f.read(2)==b'\xff\xd8'
    while True:
        b=f.read(1)
        while b!=b'\xff':b=f.read(1)
        while b==b'\xff':b=f.read(1)
        marker=b[0];length=struct.unpack('>H',f.read(2))[0]
        if marker in (0xC0,0xC1,0xC2):_,h,w=struct.unpack('>BHH',f.read(5));return w,h
        f.seek(length-2,1)
def verify(file):
    with zipfile.ZipFile(file) as z:
        roots=[E.fromstring(z.read(f'ppt/slides/slide{i}.xml')) for i in range(1,14)]
        def shapes(root):return {s.find('p:nvSpPr/p:cNvPr',N).get('name'):s for s in root.findall('.//p:sp',N)}
        outlines=[s for name,s in shapes(roots[2]).items() if name.startswith('outline-')]
        assert len(outlines)==24
        for s in outlines:
            rp=s.find('.//a:rPr',N);assert rp.find('a:noFill',N) is not None and rp.find('a:ln',N) is not None
        table=roots[6].find('.//a:tbl',N);assert table is not None
        assert len(table.findall('a:tr',N))==1 and len(table.findall('.//a:tc',N))==5
        alphas=[int(c.find('a:tcPr/a:solidFill/a:srgbClr/a:alpha',N).get('val')) for c in table.findall('.//a:tc',N)]
        assert alphas==[55000,45000,55000,45000,55000],alphas
        seventh=shapes(roots[6]);ids={s.find('p:nvSpPr/p:cNvPr',N).get('id'):name for name,s in seventh.items()}
        effect_targets=[]
        for effect in roots[6].findall('.//p:cTn',N):
            if effect.get('presetClass')!='entr':continue
            target=effect.find('.//p:spTgt',N);name=ids[target.get('spid')]
            assert name.startswith(('ed-num-','ed-text-')),name
            assert effect.get('presetID')==('30' if name.startswith('ed-num-') else '2'),(name,effect.attrib)
            effect_targets.append(name)
        assert sorted(effect_targets)==sorted([f'ed-num-{i}' for i in range(5)]+[f'ed-text-{i}' for i in range(5)])
        for i in range(5):
            text=''.join(t.text or '' for t in seventh[f'ed-text-{i}'].findall('.//a:t',N))
            assert len(text)>10 and f'ed-head-{i}' not in seventh and f'ed-copy-{i}' not in seventh
        clipped=shapes(roots[10])
        for name,prst in [('clip-rect','rect'),('clip-hexagon','hexagon'),('clip-triangle','triangle')]:
            assert clipped[name].find('p:spPr/a:prstGeom',N).get('prst')==prst
            assert clipped[name].find('p:spPr/a:blipFill',N) is not None
        glyphs=shapes(roots[12])
        for i,ch in enumerate('VIEW'):
            s=glyphs[f'glyph-{i}'];assert ''.join(t.text or '' for t in s.findall('.//a:t',N))==ch
            assert s.find('.//a:rPr/a:blipFill/a:blip',N) is not None
        image_count=0
        for index,root in enumerate(roots,1):
            rel=E.fromstring(z.read(f'ppt/slides/_rels/slide{index}.xml.rels'))
            rels={x.get('Id'):posixpath.normpath('ppt/slides/'+x.get('Target')) for x in rel}
            for s in root.findall('.//p:sp',N):
                fill=s.find('p:spPr/a:blipFill',N)
                if fill is None:continue
                image_count+=1;blip=fill.find('a:blip',N);iw,ih=jpeg_size(z.read(rels[blip.get('{'+N['r']+'}embed')]))
                crop=fill.find('a:srcRect',N);cuts={k:float(crop.get(k,'0'))/100000 for k in ['l','t','r','b']} if crop is not None else dict.fromkeys(['l','t','r','b'],0)
                ext=s.find('p:spPr/a:xfrm/a:ext',N)
                source_ratio=iw*(1-cuts['l']-cuts['r'])/(ih*(1-cuts['t']-cuts['b']))
                shape_ratio=int(ext.get('cx'))/int(ext.get('cy'))
                assert abs(source_ratio/shape_ratio-1)<.002,(index,source_ratio,shape_ratio)
    result={'passed':True,'file':str(Path(file).resolve()),'slides':13,'nativeGlyphOutlines':24,'nativeTableAlpha':alphas,'editorialEffects':10,'glyphPictureFills':4,'aspectPreservedShapeImages':image_count,'limits':'Structure does not prove native motion or visual hierarchy; separately inspect exported PNG/movie. Text glyph fill uses native stretch and is excluded from shape-image crop ratio checks.'}
    print(json.dumps(result,ensure_ascii=False,indent=2));return result
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('file');p.add_argument('--report');args=p.parse_args();r=verify(args.file)
    if args.report:
        out=Path(args.report)
        if out.exists():raise FileExistsError('Refusing verification report overwrite')
        out.write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8')
