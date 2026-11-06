"""Read-only true Kiosk/dependent trigger/link/native pixel verification."""
import argparse,json,zipfile,re
from pathlib import Path
import xml.etree.ElementTree as E
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('--deck',type=Path,required=True);p.add_argument('--evidence-directory',type=Path,required=True);p.add_argument('--correct-answer',type=int,choices=[4,5,6],required=True);p.add_argument('--migration',type=Path);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
if a.output.exists():p.error('Fresh report required')
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'};checks={};metrics={}
def need(name,ok):checks[name]=bool(ok)
with zipfile.ZipFile(a.deck) as z:
 root=E.fromstring(z.read('ppt/slides/slide1.xml'));result=E.fromstring(z.read('ppt/slides/slide2.xml'));props=E.fromstring(z.read('ppt/presProps.xml'));rels=E.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels'))
named={q.find('p:nvSpPr/p:cNvPr',n).get('name'):q for q in root.findall('.//p:sp',n)};ids={name:q.find('p:nvSpPr/p:cNvPr',n).get('id') for name,q in named.items()}
text=lambda q:''.join(x.text or '' for x in q.findall('.//a:t',n))
question=text(named['question']);nums=re.search(r'(\d+)\s*\+\s*(\d+)',question);need('actual_arithmetic_matches_correct',nums is not None and int(nums[1])+int(nums[2])==a.correct_answer)
need('saved_true_kiosk',props.find('p:showPr/p:kiosk',n) is not None)
need('no_speaker_or_browse_saved',props.find('p:showPr/p:present',n) is None and props.find('p:showPr/p:browse',n) is None)
need('no_main_animation_sequence',not root.findall('.//p:cTn[@nodeType="mainSeq"]',n))
seqs=root.findall('.//p:cTn[@nodeType="interactiveSeq"]',n);need('three_true_interactive_sequences',len(seqs)==3)
actual_triggers=[];next_membership=[]
for seq in seqs:
 trigger=seq.find('p:stCondLst/p:cond[@evt="onClick"]/p:tgtEl/p:spTgt',n).get('spid');answer=next(v for v in [4,5,6] if ids['answer-'+str(v)]==trigger);actual_triggers.append(answer)
 colors=seq.findall('.//p:animClr',n);need(str(answer)+'-two_color_behaviors',len(colors)==2)
 byattr={x.find('p:cBhvr/p:attrNameLst/p:attrName',n).text:x for x in colors}
 for attr,want in [('fillcolor','21645E' if answer==a.correct_answer else 'AF4E45'),('style.color','FFFFFF')]:
  c=byattr[attr];need(str(answer)+'-'+attr+'-target',c.find('p:cBhvr/p:tgtEl/p:spTgt',n).get('spid')==trigger)
  need(str(answer)+'-'+attr+'-color',c.find('p:to/a:srgbClr',n).get('val')==want)
  timing=c.find('p:cBhvr/p:cTn',n);need(str(answer)+'-'+attr+'-250ms_hold',timing.get('dur')=='250' and timing.get('fill')=='hold')
 entries=seq.findall('.//p:cTn[@presetClass="entr"]',n)
 if entries:next_membership.append(answer)
 need(str(answer)+'-next_dependency',len(entries)==(1 if answer==a.correct_answer else 0))
 if entries:
  entry=entries[0];need('next_fly_after_previous',entry.get('presetID')=='2' and entry.get('presetSubtype')=='4' and entry.get('nodeType')=='afterEffect')
  need('next_target_is_existing_next',all(x.get('spid')==ids['next'] for x in entry.findall('.//p:spTgt',n)))
  need('next_xy_duration_500ms',all(x.find('p:cBhvr/p:cTn',n).get('dur')=='500' for x in entry.findall('.//p:anim',n)))
need('triggers_are_all_three_answers',sorted(actual_triggers)==[4,5,6]);need('next_only_in_correct_group',next_membership==[a.correct_answer])
for value in [4,5,6]:need(str(value)+'-answer_is_not_page_link',not named['answer-'+str(value)].findall('.//a:hlinkClick',n))
link=named['next'].find('p:nvSpPr/p:cNvPr/a:hlinkClick',n);rid=link.get('{'+n['r']+'}id');target=next(x.get('Target') for x in rels if x.get('Id')==rid)
need('next_shape_has_real_slide2_link',link.get('action')=='ppaction://hlinksldjump' and target.endswith('slide2.xml'))
need('next_text_does_not_get_hyperlink_blue_style',not named['next'].findall('p:txBody//a:hlinkClick',n))
need('result_math_matches_question',any(f'2 + {a.correct_answer-2} = {a.correct_answer}' in (x.text or '') for x in result.findall('.//a:t',n)))
report=json.loads((a.evidence_directory/'report.json').read_text(encoding='utf-8-sig'));need('native_physical_run_passed',report.get('passed') is True);need('native_correct_answer',report.get('correct')==a.correct_answer)
states={x['name']:x for x in report.get('states',[])}
for name in ['initial','hidden-next-blocked','blank-blocked','keys-blocked','wrong-does-not-unlock','correct-unlocks','correct-still-needs-next','result-page']:
 need(name,states.get(name,{}).get('passed') is True)
for answer in [4,5,6]:
 if answer!=a.correct_answer:need('wrong-'+str(answer),states.get('wrong-'+str(answer),{}).get('passed') is True)
logs=report.get('checks',[])
for name in ['right','down','left','up','pagedown','space','right-after-correct']:need('key-'+name,any(x.get('name')==name and x.get('passed') is True for x in logs))
need('speaker_blank_negative_control',any(x.get('name')=='speaker-control' and x.get('passed') is True and x.get('actual')==2 for x in logs))
need('click_cursor_positions_exact',all(x.get('actualPoint')==x.get('expectedPoint') for x in logs if x.get('kind')=='physical-click'))
need('next_physical_click_navigated',any(x.get('name')=='visible-next' and x.get('before')==1 and x.get('after')==2 for x in logs))
need('desktop_movie_nonempty',(a.evidence_directory/'quiz-live.mp4').stat().st_size>10000)
scale=report['scale'];ox,oy=report['offset'];initial=Image.open(a.evidence_directory/'initial.png').convert('RGB')
def white_count(im,box):return sum(min(px)>245 for px in im.crop(box).getdata())
for answer in [4,5,6]:
 name='correct-unlocks' if answer==a.correct_answer else 'wrong-'+str(answer);im=Image.open(a.evidence_directory/(name+'.png')).convert('RGB');i=answer-4
 box=(int(ox+(2.2+i*10.1)*scale),int(oy+8.4*scale),int(ox+(11.4+i*10.1)*scale),int(oy+11.7*scale))
 baseline=white_count(initial,box);after=white_count(im,box);metrics[str(answer)+'WhitePixels']={'baseline':baseline,'feedback':after};need(str(answer)+'-feedback_keeps_actual_white_text',baseline>200 and 0.85<after/baseline<1.15)
im=Image.open(a.evidence_directory/'correct-unlocks.png').convert('RGB');box=(int(ox+25.6*scale),int(oy+15.2*scale),int(ox+31.6*scale),int(oy+17.4*scale));metrics['nextWhitePixels']=white_count(im,box);need('next_actual_caption_is_white',metrics['nextWhitePixels']>200)
if a.migration:
 migration=json.loads(a.migration.read_text(encoding='utf-8-sig'));need('native_existing_next_moved_from_old_group',migration['beforeTrigger']!=migration['afterTrigger'] and migration['afterTrigger']=='answer-'+str(a.correct_answer))
result={'passed':all(checks.values()),'checks':checks,'metrics':metrics,'boundary':'Two-page native Kiosk quiz navigation, not security/access control. Esc remains normal; no global hooks.'}
a.output.parent.mkdir(parents=True,exist_ok=True);a.output.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({'passed':result['passed'],'checks':len(checks),'failed':[k for k,v in checks.items() if not v],'metrics':metrics},indent=2));raise SystemExit(0 if result['passed'] else 1)
