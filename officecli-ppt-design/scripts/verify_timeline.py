"""Validate invariants that materially affect native timeline and delay behavior."""
import argparse,json,zipfile,xml.etree.ElementTree as E
from pathlib import Path
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/package/2006/relationships','p159':'http://schemas.microsoft.com/office/powerpoint/2015/09/main'}
def verify(path):
    with zipfile.ZipFile(path) as z:
        roots=[E.fromstring(z.read(f'ppt/slides/slide{i}.xml')) for i in range(1,6)]
        assert len([n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')])==5
        def shapes(root):
            out={}
            for s in root.findall('.//p:sp',NS):
                n=s.find('./p:nvSpPr/p:cNvPr',NS)
                if n is not None:out[n.attrib['name']]=s
            return out
        states=[shapes(r) for r in roots[1:4]]
        # Same phone components survive across all three timeline states.
        assert set(states[0]).intersection(states[1],states[2]) >= {'!!phone-0-case','!!phone-1-case','!!phone-2-case','!!axis'}
        geom=lambda s:s.find('./p:spPr/a:xfrm',NS)
        # All station centers translate by exactly 7.9cm; not just a cross-fade.
        for a,b in zip(states,states[1:]):
            for j in range(3):
                ga,gb=geom(a[f'!!year-circle-{j}']),geom(b[f'!!year-circle-{j}'])
                delta=int(gb.find('a:off',NS).attrib['x'])-int(ga.find('a:off',NS).attrib['x'])
                assert delta==-2844000,(j,delta)
        # Active phone body is full size and differs from inactive body.
        for active,ss in enumerate(states):
            sizes=[int(geom(ss[f'!!phone-{j}-case']).find('a:ext',NS).attrib['cx']) for j in range(3)]
            assert sizes[active]==1152000 and all(s<sizes[active] for j,s in enumerate(sizes) if j!=active)
        # Morph is a real p159 extension on each target timeline state.
        for root in roots[1:4]:assert any(n.tag.endswith('}morph') for n in root.iter())
        timing=roots[4].find('p:timing',NS);assert timing is not None
        paths=timing.findall('.//p:animMotion',NS);assert len(paths)==7
        assert len({n.attrib.get('path') for n in paths})==1
        looping=[n for n in timing.findall('.//p:cTn',NS) if n.attrib.get('repeatCount')=='3000' and n.attrib.get('autoRev')=='1']
        assert len(looping)==7,('Missing repeat/autoReverse on timing nodes',len(looping))
        # Delays 0..420ms belong to seven separate parallel timing branches.
        delay_values={int(n.attrib['delay']) for n in timing.findall('.//p:cond',NS) if n.attrib.get('delay','').isdigit()}
        assert {0,70,140,210,280,350,420}.issubset(delay_values),delay_values
        assert all(n.attrib.get('origin') in ('layout','parent') for n in paths)
        # Linking shapes carry no text, so native hyperlink styling cannot recolor labels.
        links=0
        for root in roots:
            for s in root.findall('.//p:sp',NS):
                if s.find('.//a:hlinkClick',NS) is not None:
                    links+=1;assert not ''.join(n.text or '' for n in s.findall('.//a:t',NS))
        assert links>=15
    result={'file':str(Path(path).resolve()),'checks':'passed','timeline_states':3,'native_motion_paths':7,'hyperlink_hotspots':links,'native_movie':'separately visually checked; this script does not prove rendering'}
    print(json.dumps(result,ensure_ascii=False,indent=2))
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('file');verify(p.parse_args().file)
