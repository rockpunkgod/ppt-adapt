"""OfficeCLI-native parent image fill inherited by grouped rectangle windows."""
import argparse,json,subprocess,zipfile,xml.etree.ElementTree as E,struct,hashlib
from pathlib import Path
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def jpeg_size(path):
    with open(path,'rb') as f:
        if f.read(2)!=b'\xff\xd8':raise ValueError('JPEG image required')
        while True:
            b=f.read(1)
            while b and b!=b'\xff':b=f.read(1)
            if not b:raise ValueError('JPEG dimensions missing')
            while b==b'\xff':b=f.read(1)
            marker=b[0];length=struct.unpack('>H',f.read(2))[0]
            if marker in (0xC0,0xC1,0xC2):_,h,w=struct.unpack('>BHH',f.read(5));return w,h
            f.seek(length-2,1)
def build(cli,out,image,probes=False):
    out=Path(out).resolve();image=Path(image).resolve()
    if any(q.exists() for q in [out,*[out.with_suffix(v) for v in ('.batch.json','.build-log.json','.model.json','.validation.json')]]):raise FileExistsError('Refusing overwrite of deck or companion evidence')
    out.parent.mkdir(parents=True,exist_ok=True)
    iw,ih=jpeg_size(image);w=min(27.5,10.7*iw/ih);h=w*ih/iw;x=(33.87-w)/2;y=5.6
    adjacent=[(0,0,0.5,0.5),(0.5,0,0.5,0.5),(0,0.5,0.5,0.5),(0.5,0.5,0.5,0.5)]
    separated=[(0,0,0.4,0.4),(0.6,0,0.4,0.4),(0,0.6,0.4,0.4),(0.6,0.6,0.4,0.4)]
    def run(*args):
        r=subprocess.run([cli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
        if r.returncode:raise RuntimeError(r.stdout+r.stderr)
        return r.stdout
    run('create',out);run('set',out,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
    ops=[]
    def add(parent,kind,**props):ops.append(dict(command='add',parent=parent,type=kind,props={k:str(v) for k,v in props.items()}))
    def shape(s,name,x,y,w,h,**props):
        base=dict(name=name,x=f'{x:.6f}cm',y=f'{y:.6f}cm',width=f'{w:.6f}cm',height=f'{h:.6f}cm',line='none',margin='0cm',fill='none');base.update(props)
        if 'image' in base:base.pop('fill',None)
        add(f'/slide[{s}]','shape',**base)
    titles=['同一张图片，使用统一坐标。','四块拼接，仍是一幅完整图。','留出间隙，保留原图的位置关系。']
    if probes:titles+=['实验：Group.Fill.UserPicture','实验：真实 group fill 与子形状继承']
    for i,title in enumerate(titles,1):
        add('/','slide',layout='blank',background='F5F3EE')
        shape(i,'title',1.6,1.3,30.2,1.8,text=title,size=33,font='Microsoft YaHei',color='173743')
        shape(i,'subtitle',1.6,3.5,30.2,1,text='所有窗口引用同一原图；核对位置与裁切，不在每块里重复整图。',size=18,font='Microsoft YaHei',color='557078')
        shape(i,'footer',1.6,17.7,30.2,0.8,text='原创窗口结构 · 照片由调用方提供',size=13,font='Microsoft YaHei',color='557078')
        if i==1:shape(i,'reference',x,y,w,h,image=str(image),srcRect='0,0,0,0')
        else:
            pieces=separated if i==3 else adjacent
            for j,(u,v,a,b) in enumerate(pieces):
                crop=','.join(str(round(q*100000)) for q in (u,v,1-u-a,1-v-b))
                props=({'image':str(image),'srcRect':'0,0,0,0'} if j==0 else {'fill':'92B7AF'}) if i in (2,3) else ({'image':str(image),'srcRect':'0,0,0,0'} if i==5 and j==0 else {'fill':'92B7AF'})
                shape(i,'piece-'+str(j),x+u*w,y+v*h,a*w,b*h,**props)
            add(f'/slide[{i}]','group',name={2:'mapped-adjacent',3:'mapped-separated',4:'api-group',5:'inherited-group'}[i],shapes=','.join(f'/slide[{i}]/shape[@name=piece-{j}]' for j in range(4)))
        add(f'/slide[{i}]','notes',text='BV1VTDGYsEmd P06 945-1074s / 1914-1947s. Original geometry, caller JPEG. Real parent image fill and inherited child fills; axis-aligned rectangles tested. No tutorial photo is bundled.')
    batch=out.with_suffix('.batch.json');batch.write_text(json.dumps(ops,ensure_ascii=False,indent=2),encoding='utf-8')
    result=run('batch',out,'--input',batch,'--json');out.with_suffix('.build-log.json').write_text(result,encoding='utf-8')
    if json.loads(result).get('success') is False:raise RuntimeError(result)
    run('close',out)
    for slide in ([2,3,5] if probes else [2,3]):
        with zipfile.ZipFile(out) as z:
            root=E.fromstring(z.read(f'ppt/slides/slide{slide}.xml'));rid=root.find('.//p:grpSp/p:sp/p:spPr/a:blipFill/a:blip',N).get('{'+N['r']+'}embed')
        fill=f'<a:blipFill xmlns:a="{N["a"]}" xmlns:r="{N["r"]}"><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
        run('raw-set',out,f'/slide[{slide}]','--xpath','//p:grpSp/p:grpSpPr','--action','append','--xml',fill)
        run('raw-set',out,f'/slide[{slide}]','--xpath','//p:grpSp/p:sp/p:spPr/*[self::a:solidFill or self::a:blipFill or self::a:noFill]','--action','replace','--xml',f'<a:grpFill xmlns:a="{N["a"]}"/>')
        run('close',out)
    out.with_suffix('.model.json').write_text(json.dumps({'image':str(image),'imageSha256':hashlib.sha256(image.read_bytes()).hexdigest(),'imageSize':[iw,ih],'stageCm':[x,y,w,h],'adjacent':adjacent,'separated':separated,'probes':probes},indent=2),encoding='utf-8')
    out.with_suffix('.validation.json').write_text(run('validate',out,'--json'),encoding='utf-8');print(out)
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);p.add_argument('--image',required=True);a=p.parse_args();build(a.officecli,a.out,a.image)
