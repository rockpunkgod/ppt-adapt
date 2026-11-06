"""Independent artifact + native screen audit. Missing evidence always stays unverified."""
import argparse, hashlib, json, math, pathlib, re, subprocess, unicodedata, zipfile
import xml.etree.ElementTree as ET
from PIL import Image, ImageChops, ImageDraw

NS={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main','c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
EMU=360000

def load(path):
    return json.loads(pathlib.Path(path).read_text(encoding='utf-8-sig'))

def sha(path):
    return hashlib.sha256(pathlib.Path(path).read_bytes()).hexdigest().upper()

def canonical_sha(value):
    return hashlib.sha256(json.dumps(value,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf-8')).hexdigest()

def norm(s):
    return re.sub(r'\s+','',unicodedata.normalize('NFKC',str(s)))

def rgb(value):
    v=int(value);return (v&255,(v>>8)&255,(v>>16)&255)

def contrast(a,b):
    def lum(c):
        s=[v/255 for v in c];s=[v/12.92 if v<=0.04045 else ((v+0.055)/1.055)**2.4 for v in s]
        return sum(x*y for x,y in zip(s,[0.2126,0.7152,0.0722]))
    x,y=sorted([lum(a),lum(b)]);return (y+0.05)/(x+0.05)

def flatten(shapes):
    for x in shapes:
        yield x
        yield from flatten(x.get('children',[]))

def resolve_name(deck_id,o):
    return '!!'+o['morph_key'] if o.get('morph_key') else deck_id+'.'+o['id']

def rect_px(bounds,image,native,spec,desktop=False):
    x,y,w,h=bounds
    if desktop:
        ox,oy,cw,ch=native['contentRectDwmPixels']
        sx=cw/spec['canvas_cm']['width'];sy=ch/spec['canvas_cm']['height']
    else:
        ox=oy=0;sx=image.width/spec['canvas_cm']['width'];sy=image.height/spec['canvas_cm']['height']
    return tuple(round(v) for v in (ox+x*sx,oy+y*sy,ox+(x+w)*sx,oy+(y+h)*sy))

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--deck',required=True);ap.add_argument('--plan',required=True);ap.add_argument('--spec',required=True);ap.add_argument('--data',required=True);ap.add_argument('--evidence',required=True);ap.add_argument('--out',required=True);ap.add_argument('--regression-deck');ap.add_argument('--regression-evidence');ap.add_argument('--composition');ap.add_argument('--ffmpeg');ap.add_argument('--visual-reviewer',help='Only supply after personally inspecting contact and real movie');args=ap.parse_args()
    out=pathlib.Path(args.out)
    if out.exists():raise SystemExit('Refusing overwrite report')
    ev=pathlib.Path(args.evidence);plan=load(args.plan);spec=load(args.spec);data=load(args.data);checks=[];composition=load(args.composition) if args.composition else None
    def check(name,ok,detail=None):checks.append({'name':name,'passed':bool(ok),'detail':detail})
    try:
        native=load(ev/'native.json')
        check('native_capture_succeeded',native.get('passed') and native.get('status')=='recorded-pending-independent-audit')
        for key,path in [('deck',args.deck),('plan',args.plan),('spec',args.spec)]:check('input_sha_'+key,native.get('inputSha256',{}).get(key)==sha(path))
        cleanup=native.get('cleanup',{});check('own_cleanup',cleanup.get('ownDeckClosed') and cleanup.get('ownShowCount')==0 and cleanup.get('ownerEnabledOrDestroyed') and all(cleanup.get('hashUnchanged',{}).get(k) for k in ['deck','plan','spec']))
        check('guarded_own_surface',native.get('ownPid',0)>0 and native.get('surfaceChecks',0)>100 and len(native.get('dwm',[]))==4)
        check('eight_pages_plan',len(plan['pages'])==8);check('layout_diversity',len({p['layout_family'] for p in plan['pages']})>=5)
        check('plan_information_roles',len({p['role'] for p in plan['pages']})>=6)
        check('simulation_notice',bool(data.get('disclaimer')) and '模拟' in data['disclaimer'])
        demand=data['demand'];v=demand['series'][0]['values'];w=demand['series'][1]['values']
        check('data_250_200_150_60pp',v==[50,70,80,50] and w==[50,70,80,0] and sum(v)==250 and sum(w)==200 and demand['current']['covered']==50 and demand['proposed']['covered']==200 and demand['increment']['covered']==150 and demand['increment']['percentage_points']==60 and abs(demand['proposed']['coverage_rate']-demand['current']['coverage_rate']-0.6)<1e-9)
        check('budget_8_6_4_equals_18',[x['amount'] for x in data['budget']['items']]==[8,6,4] and sum(x['amount'] for x in data['budget']['items'])==data['budget']['total']==18)
        check('roadmap_12_weeks',[(p['start_week'],p['end_week']) for p in data['roadmap']['phases']]==[(1,2),(3,10),(11,12)])
        z=zipfile.ZipFile(args.deck);slide_parts=sorted([n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)],key=lambda n:int(re.search(r'slide(\d+)',n).group(1)))
        check('pptx_eight_slides',len(slide_parts)==8)
        roots=[ET.fromstring(z.read(n)) for n in slide_parts]
        native_slides={int(s['index']):s for s in native.get('slides',[])}
        for i,(page,root) in enumerate(zip(plan['pages'],roots),1):
            objects={n.get('name') for n in root.findall('.//p:cNvPr',NS)};shapes={s['name']:s for s in flatten(native_slides.get(i,{}).get('shapes',[]))}
            alltext=''.join(t.text or '' for t in root.findall('.//a:t',NS))
            check(f'{page["id"]}_simulation_visible_text','模拟' in alltext)
            for o in page['objects']:
                name=resolve_name(plan['deck_id'],o);check('object_exists_'+o['id'],name in objects and name in shapes)
                if name not in shapes:continue
                ns=shapes[name]
                if o.get('text'):
                    check('editable_text_'+o['id'],norm(ns.get('text',''))==norm(o['text']) and norm(o['text']) in norm(alltext))
                if o.get('font_role') and ns.get('text'):
                    role=o['font_role'];expected=spec['tokens']['font_sizes_pt'].get(role,spec['tokens']['font_sizes_pt']['body']);font=ns.get('font',{});size=float(font.get('size',0));check('font_size_'+o['id'],size>=spec['quality']['min_body_pt'] if role=='body' else size>=10)
                    ea=font.get('ea','');latin=font.get('latin','');check('resolved_font_'+o['id'],ea in ['Microsoft YaHei','微软雅黑'] and latin=='Arial',font)
                    tb=ns.get('textBoundsPt');x=ns['leftPt'];y=ns['topPt'];ww=ns['widthPt'];hh=ns['heightPt']
                    check('text_inside_frame_'+o['id'],tb is not None and tb[0]>=x-1.5 and tb[1]>=y-1.5 and tb[0]+tb[2]<=x+ww+1.5 and tb[1]+tb[3]<=y+hh+1.5,{'text':tb,'shape':[x,y,ww,hh]})
                    bg=spec['tokens']['colors']['night' if page.get('variant')=='night' else 'paper'];bg=tuple(bytes.fromhex(bg));fg=rgb(font.get('rgb',0));ratio=contrast(fg,bg)
                    if o['kind']=='shape' and ns.get('fillRgb') is not None:ratio=contrast(fg,rgb(ns['fillRgb']))
                    check('text_contrast_'+o['id'],ratio>=(3 if size>=24 else 4.5),round(ratio,3))
            # Actual text ink bounding boxes, excluding deliberate group/container nesting.
            texts=[s for s in shapes.values() if s.get('text') and s.get('textBoundsPt') and not s.get('children')]
            overlap=[]
            for ai,a in enumerate(texts):
                ax,ay,aw,ah=a['textBoundsPt']
                for b in texts[ai+1:]:
                    bx,by,bw,bh=b['textBoundsPt'];iw=min(ax+aw,bx+bw)-max(ax,bx);ih=min(ay+ah,by+bh)-max(ay,by)
                    if iw>2 and ih>2:overlap.append([a['name'],b['name'],round(iw*ih,2)])
            check('actual_text_nonoverlap_'+page['id'],not overlap,overlap)
        chartparts=[n for n in z.namelist() if re.fullmatch(r'ppt/charts/chart\d+\.xml',n)];check('one_true_chart',len(chartparts)==1)
        if chartparts:
            ch=ET.fromstring(z.read(chartparts[0]));series=ch.findall('.//c:barChart/c:ser',NS);actual=[]
            for s in series:actual.append([float(t.text) for t in s.findall('./c:val/c:numRef/c:numCache/c:pt/c:v',NS)] or [float(t.text) for t in s.findall('./c:val/c:numLit/c:pt/c:v',NS)])
            check('true_chart_2x4_values',actual==[[50,70,80,50],[50,70,80,0]],actual)
            cats=series[0].findall('./c:cat//c:pt/c:v',NS) if series else [];check('four_chart_categories',len(cats)==4)
            native_charts=[s['chart'] for s in flatten(native_slides.get(3,{}).get('shapes',[])) if s.get('chart')]
            check('native_chart_values_match',len(native_charts)==1 and [[float(v) for v in s['values']] for s in native_charts[0]['series']]==[[50,70,80,50],[50,70,80,0]])
            check('native_value_axis_0_to_100',len(native_charts)==1 and any(a['type']==2 and a['minimum']==0 and a['maximum']==100 for a in native_charts[0]['axes']))
        table=roots[3].find('.//a:tbl',NS);check('true_table_6x5',table is not None and len(table.findall('a:tr',NS))==6 and all(len(r.findall('a:tc',NS))==5 for r in table.findall('a:tr',NS)))
        if table is not None:
            rows=[[''.join(t.text or '' for t in c.findall('.//a:t',NS)) for c in r.findall('a:tc',NS)] for r in table.findall('a:tr',NS)];check('table_data_matches',[[norm(c) for c in r] for r in rows]==[[norm(c) for c in r] for r in [data['comparison']['columns']]+data['comparison']['rows']],rows)
            native_tables=[s['table'] for s in flatten(native_slides.get(4,{}).get('shapes',[])) if s.get('table')]
            cells=[c for r in native_tables[0]['cells'] for c in r] if native_tables else []
            check('native_table_body_font_floor',len(cells)==30 and all(c['size']>=spec['quality']['min_body_pt'] for c in cells))
            check('native_table_ea_latin',len(cells)==30 and all(c.get('ea') in ['Microsoft YaHei','微软雅黑'] and c.get('latin')=='Arial' for c in cells))
            clipped=[]
            for i,c in enumerate(cells):
                b=c.get('textBoundsPt');x,y,w,h=c['boundsPt']
                # PowerPoint cell TextFrame2 returns local BoundLeft and baseline-relative
                # BoundTop (observed -12 for every 20pt cell), unlike ordinary shapes.
                # Compare measured width/height to available cell space, not false global XY.
                if not b or b[0]<-1.5 or b[2]+2*max(0,b[0])>w+1.5 or b[3]>h+1.5:clipped.append(i)
            check('native_table_text_dimensions_fit',len(cells)==30 and not clipped,{'bad_cells':clipped,'coordinate_limit':'Cell TextFrame2 XY is local/baseline-relative; dimensions plus actual native ink margins are checked separately'})
            table_image=Image.open(ev/'native-4.png').convert('RGB');bad_ink=[]
            for i,c in enumerate(cells):
                x,y,w,h=c['boundsPt'];cm=[v/(72/2.54) for v in [x,y,w,h]];box=rect_px(cm,table_image,native,spec);crop=table_image.crop(box);fg=rgb(c.get('rgb',0));points=[]
                for yy in range(crop.height):
                    for xx in range(crop.width):
                        if max(abs(crop.getpixel((xx,yy))[k]-fg[k]) for k in range(3))<55:points.append((xx,yy))
                if not points:bad_ink.append({'cell':i,'reason':'no actual foreground ink'})
                else:
                    xs,ys=zip(*points);margins=[min(xs),min(ys),crop.width-1-max(xs),crop.height-1-max(ys)]
                    if min(margins)<2:bad_ink.append({'cell':i,'ink_margins_px':margins})
            check('native_table_ink_has_cell_edge_clearance',len(cells)==30 and not bad_ink,bad_ink)
            check('native_table_text_contrast',len(cells)==30 and all(contrast(rgb(c.get('rgb',0)),rgb(c.get('fillRgb',0)))>=4.5 for c in cells))
        theme=ET.fromstring(z.read('ppt/theme/theme1.xml'));cs=theme.find('.//a:clrScheme',NS);check('native_theme_accent1',cs.find('a:accent1/a:srgbClr',NS).get('val','').upper()==spec['tokens']['colors']['accent1'])
        fontscheme=theme.find('.//a:fontScheme',NS);check('theme_ea_latin',all(fontscheme.find('a:'+role+'/a:ea',NS).get('typeface')=='Microsoft YaHei' and fontscheme.find('a:'+role+'/a:latin',NS).get('typeface')=='Arial' for role in ['majorFont','minorFont']))
        for index in [4,5]:
            names=[n.get('name') for n in roots[index].findall('.//p:cNvPr',NS)];check('morph_keys_unique_'+str(index+1),all(names.count('!!'+key)==1 for key in ['coverage-window','focusTag']))
        morphs=roots[5].findall('.//p:transition',NS);expected_morph_ms=str(plan['pages'][5]['motion']['duration_ms']);check('real_morph_matches_plan_duration',any(any(x.tag.split('}')[-1]=='morph' for x in t.iter()) and any(k.split('}')[-1]=='dur' and v==expected_morph_ms for k,v in t.attrib.items()) for t in morphs))
        cnames=[x for x in roots[2].iter() if x.tag.split('}')[-1] in ['chart','bldChart','bldDgm']]
        build=roots[2].findall('.//a:bldChart',NS)+roots[2].findall('.//p:bldChart',NS);check('native_chart_category_build',any(x.get('bld')=='category' for x in build),[x.attrib for x in build])
        events=native.get('events',[]);check('eleven_physical_right_events',len(events)==11 and all(e.get('input')=='physical-Right' and e.get('actualSlide')==e.get('expectedSlide') for e in events))
        builds=[e for e in events if e.get('expectedSlide')==3];check('four_real_chart_click_indices',[e.get('actualClick') for e in builds]==[0,1,2,3,4])
        effects=native_slides.get(3,{}).get('effects',[]);planned_groups=plan['pages'][2]['motion']['groups'];check('four_chart_native_effect_durations',len(effects)==len(planned_groups)==4 and all(abs(float(e['duration'])-g['duration_ms']/1000)<0.002 and e['trigger']==1 for e,g in zip(effects,planned_groups)),effects)
        # Dynamic chart evidence: successive category clusters must change in increasing x order.
        chart_o=next(o for o in plan['pages'][2]['objects'] if o['kind']=='chart');centers=[];deltas=[]
        for a,b in zip(builds,builds[1:]):
            ia=Image.open(ev/a['capture']).convert('RGB');ib=Image.open(ev/b['capture']).convert('RGB');box=rect_px(chart_o['bounds_cm'],ia,native,spec,True);diff=ImageChops.difference(ia.crop(box),ib.crop(box));mask=diff.convert('L').point(lambda v:255 if v>35 else 0);bbox=mask.getbbox();hist=mask.histogram();n=hist[255];deltas.append(n);centers.append(None if bbox is None else (bbox[0]+bbox[2])/2)
        check('chart_four_visible_build_changes',len(deltas)==4 and min(deltas)>50,deltas)
        check('chart_categories_progress_left_to_right',len(centers)==4 and all(x is not None for x in centers) and all(b-a>10 for a,b in zip(centers,centers[1:])),centers)
        # Morph window: inspect actual accent run at the planned middle y; no fixed screen pixels.
        s5event=next(e for e in events if e.get('actualSlide')==5);s6event=next(e for e in events if e.get('actualSlide')==6);color=tuple(bytes.fromhex(spec['tokens']['colors']['accent1']))
        window_a=next(o for o in plan['pages'][4]['objects'] if o.get('morph_key')=='coverage-window');window_b=next(o for o in plan['pages'][5]['objects'] if o.get('morph_key')=='coverage-window');track=next(o for o in plan['pages'][4]['objects'] if o['id'].endswith('.track'));track_bounds=track['bounds_cm']
        def window_run(path):
            im=Image.open(path).convert('RGB');box=rect_px(track_bounds,im,native,spec,True);y=(box[1]+box[3])//2;flags=[max(abs(im.getpixel((x,y))[k]-color[k]) for k in range(3))<30 for x in range(box[0],box[2])];run=best=0
            for f in flags:run=run+1 if f else 0;best=max(best,run)
            return best/(box[2]-box[0])*track_bounds[2]
        start=window_run(ev/s5event['capture']);end=window_run(ev/s6event['capture']);midfiles=sorted(ev.glob('morph-mid-*.png'),key=lambda p:float(p.stem.split('-')[-1]));mids=[window_run(f) for f in midfiles]
        check('morph_endpoint_geometry',abs(start-window_a['bounds_cm'][2])<0.15 and abs(end-window_b['bounds_cm'][2])<0.15,{'start_cm':start,'end_cm':end})
        check('morph_real_intermediate_geometry',len(mids)>=3 and any(start+0.2<m<end-0.2 for m in mids) and all(b+0.15>=a for a,b in zip(mids,mids[1:])),mids)
        for index in [5,6]:
            im=Image.open(ev/f'native-{index}.png').convert('RGB');top=im.getpixel((round(im.width*0.97),round(im.height*0.08)));bottom=im.getpixel((round(im.width*0.97),round(im.height*0.98)));check('night_and_light_footer_'+str(index),max(abs(top[k]-bytes.fromhex(spec['tokens']['colors']['night'])[k]) for k in range(3))<3 and max(abs(bottom[k]-bytes.fromhex(spec['tokens']['colors']['paper'])[k]) for k in range(3))<3)
        movie=ev/native.get('movie','integration-live.mp4');check('movie_exists',movie.exists() and movie.stat().st_size>1000)
        if args.ffmpeg and movie.exists():
            decode=subprocess.run([args.ffmpeg,'-v','error','-i',str(movie),'-map','0:v:0','-f','null','-'],capture_output=True,text=True,timeout=35)
            check('full_movie_decodes',decode.returncode==0,decode.stderr[-1000:])
        else:check('full_movie_decodes',False,'Provide --ffmpeg for independent full video decode')
        contact=Image.new('RGB',(1600,490),'#ddd');draw=ImageDraw.Draw(contact)
        for i in range(8):im=Image.open(ev/f'native-{i+1}.png').convert('RGB');im.thumbnail((400,225));contact.paste(im,((i%4)*400,(i//4)*245));draw.text(((i%4)*400+4,(i//4)*245+226),plan['pages'][i]['id']+' '+plan['pages'][i]['role'],fill='black')
        check('native_static_pages_visually_distinct',len({sha(ev/f'native-{i}.png') for i in range(1,9)})==8)
        contactpath=out.with_name(out.stem+'-contact.jpg')
        if contactpath.exists():raise FileExistsError('Refusing existing contact')
        contact.save(contactpath)
        # Optional regression input: same content with accent1 changed, direct night preserved.
        if args.regression_deck and args.regression_evidence:
            rz=zipfile.ZipFile(args.regression_deck);rt=ET.fromstring(rz.read('ppt/theme/theme1.xml'));newaccent=rt.find('.//a:clrScheme/a:accent1/a:srgbClr',NS).get('val');check('regression_global_accent_changed',newaccent.upper()!=spec['tokens']['colors']['accent1'])
            regression_native=load(pathlib.Path(args.regression_evidence)/'native.json')
            check('regression_real_native_capture_and_deck_bound',regression_native.get('passed') and regression_native.get('inputSha256',{}).get('deck')==sha(args.regression_deck) and regression_native.get('inputSha256',{}).get('spec')==sha(args.spec))
            regression_context={'prior_plan_sha256':regression_native.get('inputSha256',{}).get('plan'),'final_plan_sha256':sha(args.plan),'note':'Regression recording precedes the final COM DAG declaration. Only operations metadata was added afterward; compare current main/variant slide XML text and native states independently. No page change is inferred from a plan hash alone.'}
            same=[]
            for n in slide_parts:
                a=ET.fromstring(z.read(n));b=ET.fromstring(rz.read(n));same.append([x.text for x in a.findall('.//a:t',NS)]==[x.text for x in b.findall('.//a:t',NS)])
            check('regression_content_preserved',all(same));reg_ev=pathlib.Path(args.regression_evidence);changes=[]
            for sn in [1,3,5,6]:
                a=Image.open(ev/f'native-{sn}.png').convert('RGB');b=Image.open(reg_ev/f'native-{sn}.png').convert('RGB');check('regression_same_dimensions_'+str(sn),a.size==b.size);changes.append(ImageChops.difference(a,b).getbbox() is not None)
                if sn in [5,6]:check('regression_local_night_preserved_'+str(sn),a.getpixel((a.width-20,20))==b.getpixel((b.width-20,20)))
            check('regression_accent_visibly_changes',any(changes))
        else:check('global_token_and_local_exception_regression',False,'Supply --regression-deck and --regression-evidence after separate native control; no assumption from XML alone')
        if composition:
            bound={k:v for k,v in composition.items() if k not in {'plan_sha256','verification','status'}}
            check('composition_digest_valid',canonical_sha(bound)==composition.get('plan_sha256'))
            check('composition_spec_and_plan_bound',composition.get('input_hashes',{}).get('spec')==canonical_sha(spec) and composition.get('input_hashes',{}).get('deck_plan')==canonical_sha(plan))
            idx=composition.get('index',{});check('composition_index_bytes_bound',pathlib.Path(idx.get('path','')).is_file() and sha(idx['path']).lower()==idx.get('sha256'))
            selected=composition.get('selected_recipes',[])
            check('composition_reuse_scope_declared',bool(selected) and all(bool(x.get('reuse_scope')) for x in selected))
        else:check('composition_hash_binding',False,'Supply --composition for final eight-page acceptance binding')
    except Exception as exc:
        check('audit_completed',False,type(exc).__name__+': '+str(exc))
    check('explicit_visual_contact_and_movie_review',bool(args.visual_reviewer),args.visual_reviewer or 'Pending actual visual review; do not auto-fill reviewer')
    passed=bool(checks) and all(c['passed'] for c in checks)
    report={'version':1,'deck_id':plan.get('deck_id'),'scope':'whole-deck','kind':'native-powerpoint','passed':passed,'acceptance_id':ev.resolve().name,'combination_verified':passed,'status':'verified' if passed else 'unverified','checks':checks,'passed_count':sum(c['passed'] for c in checks),'total':len(checks),'visual_review_scope':args.visual_reviewer,'limits':['Visual review is native static pages plus sampled whole-duration movie frames and key category/transition states; not continuous normal-speed viewing. No frame-exact key/video synchronization claim.','Regression and composition binding are mandatory for final true; optional CLI arguments only support staged evidence.'],'artifacts':{'contact':out.with_name(out.stem+'-contact.jpg').name},'pages':[p['id'] for p in plan['pages']],'evidence':[]}
    for i,page in enumerate(plan['pages'],1):
        path=ev/f'native-{i}.png'
        if path.is_file():report['evidence'].append({'kind':'png','page_id':page['id'],'path':str(path.resolve()),'sha256':sha(path).lower()})
    movie=ev/'integration-live.mp4'
    if movie.is_file():report['evidence'].append({'kind':'native-recording','path':str(movie.resolve()),'sha256':sha(movie).lower()})
    report['deck_sha256']=sha(args.deck).lower()
    if args.regression_deck and args.regression_evidence:report['regression_context']=locals().get('regression_context',{'note':'Regression audit did not reach context binding'})
    if composition:report.update({'plan_sha256':composition.get('plan_sha256'),'index_sha256':composition.get('index',{}).get('sha256'),'spec_sha256':composition.get('input_hashes',{}).get('spec'),'deck_plan_sha256':composition.get('input_hashes',{}).get('deck_plan')})
    out.parent.mkdir(parents=True,exist_ok=True);out.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');print(json.dumps({'combination_verified':passed,'passed':report['passed_count'],'total':len(checks),'failed':[c['name'] for c in checks if not c['passed']]},ensure_ascii=False));raise SystemExit(0 if passed else 1)

if __name__=='__main__':main()
