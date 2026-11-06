"""Author the TRACE example with existing OfficeCLI primitives; no new deck engine."""
import argparse
import json
from pathlib import Path
import sys

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'officecli-ppt-design/scripts'))
from build_components import Deck

PAPER='F2F3EC'; INK='16251D'; ACCENT='DAEF78'; MUTED='637066'
DISPLAY='Noto Serif SC'; BODY='Noto Sans SC'

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--officecli',required=True);ap.add_argument('--out',required=True)
    args=ap.parse_args();d=Deck(args.out,args.officecli);pending=[]
    d.run('set','/',props={'theme.font.major.eastAsia':DISPLAY,'theme.font.major.latin':DISPLAY,
        'theme.font.minor.eastAsia':BODY,'theme.font.minor.latin':BODY,
        'theme.color.accent1':ACCENT,'theme.color.accent2':INK})
    def shape(s,name,x,y,w,h,text='',size=20,color=INK,fill='none',display=False,**extra):
        p={'name':name,'x':f'{x}cm','y':f'{y}cm','width':f'{w}cm','height':f'{h}cm','fill':fill,'line':'none'}
        if text:
            role='mj' if display else 'mn'
            p.update(text=text,font=DISPLAY if display else BODY,size=str(size),color=color,
                margin='0cm',valign='center',autoFit='none',**{'font.ea':f'+{role}-ea','font.latin':f'+{role}-lt','lineSpacing':'1.18x'})
        p.update({k:(str(v).removesuffix('pt') if k=='spacing' else str(v).lower() if isinstance(v,bool) else str(v)) for k,v in extra.items()})
        pending.append({'command':'add','parent':f'/slide[{s}]','type':'shape','props':p})
        return name
    def flush():
        if pending:d.batch(pending.copy());pending.clear()
    def footer(s,dark=False):
        c='BCC7BB' if dark else MUTED
        shape(s,'deck-label',1.5,17.85,28,.6,'循迹 TRACE  /  原创概念产品演示',11,c)
        shape(s,'page-number',31.1,17.85,1.25,.6,f'{s:02}',11,c,align='right')
    def label(s,value,dark=False):shape(s,'section-label',1.5,1.1,30.87,.7,value,12,ACCENT if dark else MUTED)
    def heading(s,value,dark=False):shape(s,'page-title',1.5,2.15,30.87,2.05,value,38,PAPER if dark else INK,display=True,spacing='-0.4pt')
    def notes(s,text):flush();d.notes(s,text+'\n循迹为原创概念产品；界面与数字用于设计演示，非在售功能或实测性能。')

    # 1 — Product, not a generic technology wallpaper.
    s=d.slide(INK);label(s,'团队知识工作台  /  概念发布',True)
    shape(s,'wordmark',1.5,2.2,14.5,3.3,'TRACE',68,PAPER,display=True,spacing='-1.6pt')
    shape(s,'hero-statement',1.5,6.1,14.3,4.8,'让团队的知识，\n成为此刻的答案。',42,PAPER,display=True,spacing='-0.3pt')
    shape(s,'hero-lead',1.5,12.5,13.8,2.0,'找回证据，保留上下文。\n把答案接回正在做的工作。',20,'D7DED4')
    x,y,w,h=16.8,4.0,15.57,10.8
    shape(s,'hero-ui-border',x,y,w,h,fill='7A8778');shape(s,'hero-ui',x+.035,y+.035,w-.07,h-.07,fill='FCFDF8')
    shape(s,'hero-ui-brand',x+.65,y+.45,10,.8,'TRACE / 团队空间',13)
    shape(s,'hero-ui-search',x+.65,y+1.65,w-1.3,1.3,'找上周确认的设计决策',16,fill='E9EDE2',margin='0.3cm')
    shape(s,'hero-ui-question',x+.65,y+3.4,w-1.3,2.6,'首版决策：\n统一检索入口。',27,display=True)
    shape(s,'hero-ui-answer',x+.65,y+6.4,w-1.3,1.5,'负责人：林禾\n先接产品资料，再接会议记录。',18)
    shape(s,'hero-ui-cite',x+.65,y+8.65,8.5,.8,'2条来源 / 1项待办',13,MUTED)
    shape(s,'hero-ui-action',x+9.5,y+8.5,5.2,1.1,'查看待办',14,INK,ACCENT,align='center')
    footer(s,True);notes(s,'从具体工作开始介绍循迹：它是团队知识工作台，将资料检索、来源和任务交接放在同一个上下文里。右侧为可编辑概念界面。')

    # 2 — A type-led problem page, without a repeated feature-card row.
    s=d.slide(PAPER);label(s,'01 / 为什么需要循迹');heading(s,'找到文件，还没接上工作。')
    shape(s,'scattered',1.5,6.1,15.0,5.4,'散落。',88,display=True,spacing='-1.2pt')
    for i,(title,body) in enumerate([('对话在聊天里','讨论过的背景，很快被新消息推远。'),('文档在文件夹里','找到了文件，还要重新拼出前因后果。'),('决策在人的记忆里','知道结论，却不知道为什么这样决定。')]):
        yy=5.8+i*3.25
        shape(s,f'problem-number-{i}',18.0,yy,.9,.7,f'{i+1:02}',12,MUTED)
        shape(s,f'problem-title-{i}',19.35,yy,13.0,1.0,title,23)
        shape(s,f'problem-body-{i}',19.35,yy+1.25,13.0,1.55,body,18,MUTED,valign='top')
    shape(s,'problem-conclusion',1.5,13.0,14.2,2.0,'知识的价值，\n发生在继续工作的那一刻。',22)
    footer(s);notes(s,'问题不只在于资料能不能被找到，更在于能否理解背景并继续推进。分别指出对话、文档和决策的断点，不虚构客户调查或工具数量。')

    # 3 — A wide, editable product scene; group identity is reused on page 4.
    s=d.slide(PAPER);label(s,'01 / 产品工作台');heading(s,'从证据，到下一步。')
    ui=[]
    def u(key,x,y,w,h,text='',size=18,color=INK,fill='none',**extra):
        name='!!ui-'+key;ui.append(shape(s,name,x,y,w,h,text,size,color,fill,**extra));return name
    u('border',1.5,5,30.87,10.4,fill='C6CFC2');u('ground',1.535,5.035,30.8,10.33,fill='FCFDF8')
    u('header',1.535,5.035,30.8,1.25,fill='E6EBDF');u('brand',2.1,5.24,4.4,.7,'TRACE',14)
    u('workspace',24.6,5.24,6.8,.7,'产品组 / 团队空间',13,MUTED)
    u('sidebar',1.535,6.285,5.55,9.08,fill='EDF0E6');u('sidebar-head',2.05,6.7,4.5,.7,'工作空间',13,MUTED)
    for i,txt in enumerate(['全部资料','产品设计','项目会议','已确认决策']):
        u('sidebar-'+str(i),2.0,8.05+i*1.45,4.6,1.0,txt,15,INK,ACCENT if i==3 else 'none',margin='0.15cm')
    u('search',8.0,6.85,23.0,1.25,'找上周确认的设计决策',18,INK,'EDF0E6',margin='0.3cm')
    u('answer-label',8.0,8.6,10,.7,'与当前问题相关的结论',13,MUTED)
    u('answer',8.0,9.55,22.8,2.4,'决定：首版保留一个统一检索入口。\n负责人：林禾｜先接产品资料，再接会议记录。',22)
    u('citations',8.0,12.1,22.0,.85,'来源 1 / 产品讨论     来源 2 / 首版决策记录',14,MUTED)
    u('status',8.0,13.65,14,.75,'待办：由林禾整理首批试用资料',13,MUTED)
    u('action',25.0,13.55,5.6,1.1,'查看待办',15,INK,ACCENT,align='center')
    footer(s);flush()
    d.run('add',f'/slide[{s}]','--type','group',props={'name':'!!product-ui','shapes':','.join(f'/slide[{s}]/shape[@name={name}]' for name in ui)})
    d.notes(s,'先看到问题与答案，再看到来源。界面为原生可编辑形状和文字。下一页保留同一个产品场景，缩移后展示工作交接。')

    # 4 — Existing continuous-scene Morph: same group, uniform scale, new work context.
    s=d.clone(3)
    d.set(s,'page-title',text='从证据，到下一步。')
    d.set(s,'page-number',text='04')
    d.run('set',f'/slide[{s}]/group[@name=!!product-ui]',props={'x':'1.5cm','y':'6.2cm','width':f'{30.87*.62}cm','height':f'{10.4*.62}cm'})
    shape(s,'handoff-arrow',20.9,9.25,.75,.48,fill=INK,geometry='rightArrow')
    shape(s,'handoff-title',22.2,5.65,10.17,3.4,'结论有来源，\n行动有归属。',31,display=True)
    shape(s,'handoff-body',22.2,10.1,10.17,3.6,'把引用放进决策。\n把行动交给负责人。',20)
    shape(s,'handoff-meta',1.5,14.9,19.2,.85,'同一上下文，连接检索、判断和交接。',15,MUTED)
    flush();d.run('set',f'/slide[{s}]',props={'transition':'morph-850','advanceClick':True})
    d.notes(s,'产品窗口连续缩移，为右侧的行动留下空间。重点是从获得答案到继续工作；不是把检索结果当作自动授权，关键动作仍由人确认。')

    # 5 — A deliberate pause and chapter change.
    s=d.slide(INK);label(s,'02 / 工作流，不只检索',True)
    shape(s,'chapter-statement',1.5,4.1,30.87,6.2,'从“查到”，\n到“接着做”。',56,PAPER,display=True,spacing='-0.5pt')
    for i,txt in enumerate(['检索','引用','交接']):shape(s,f'chapter-step-{i}',1.5+i*10.5,13.55,8.8,1.6,txt,25,ACCENT)
    shape(s,'chapter-link-a',6.5,14.25,4.0,.035,fill='899781');shape(s,'chapter-link-b',17.0,14.25,4.0,.035,fill='899781')
    footer(s,True);notes(s,'章节停顿。将产品能力收束为检索、引用、交接三步，后续分别解释效率、控制和试用。')

    # 6 — Existing native-chart series build, with explicitly illustrative values.
    s=d.slide(PAPER);label(s,'02 / 一次工作流的演示');heading(s,'把时间留给判断。')
    shape(s,'chart-unit',1.5,4.55,21.4,.75,'单次资料定位与整理 / 分钟',15,MUTED)
    shape(s,'delta',24.0,5.3,8.37,4.6,'11',88,display=True,spacing='-1.2pt')
    shape(s,'delta-caption',24.0,10.2,8.37,2.0,'分钟\n演示差值',20)
    shape(s,'delta-math',24.0,13.0,8.37,1.0,'18 − 7 = 11',18,MUTED)
    shape(s,'chart-disclaimer',1.5,16.3,30.87,.8,'演示情景，非产品实测。数值只用于说明比较图与逐步讲解。',14,MUTED)
    flush()
    d.run('add',f'/slide[{s}]','--type','chart',props={'name':'evidence-chart','chartType':'bar','categories':'资料定位与整理',
        'data':'常规整理:18;循迹辅助:7','x':'1.5cm','y':'5.55cm','width':'21.4cm','height':'9.8cm','title':'none',
        'legend':'bottom','legendFont':f'16:{INK}:{BODY}','axisfont':f'15:{MUTED}:{BODY}',
        'axismin':'0','axismax':'20','majorunit':'5','chartareafill':'none','chartborder':'none','plotareafill':'none','plotborder':'none',
        'gridlines':'D9DED3:0.5','colors':'727B70,203D29','overlap':'0'})
    d.run('add',f'/slide[{s}]/chart[1]','--type','animation',props={'effect':'fade','class':'entrance','chartBuild':'series','trigger':'onClick','duration':'450'})
    footer(s);notes(s,'点击两次，先展示常规整理18分钟，再展示循迹辅助7分钟。数值为演示情景，不是测试结果或性能承诺。差值11分钟仅为同一情景内的算术比较。')

    # 7 — Relationship diagram: boundaries have meaning, rather than decorative cards.
    s=d.slide(INK);label(s,'02 / 设计原则',True);heading(s,'控制权，留在团队手里。',True)
    shape(s,'boundary-top',1.5,5.65,30.87,.035,fill='697D69');shape(s,'boundary-bottom',1.5,15.3,30.87,.035,fill='697D69')
    shape(s,'boundary-label',1.5,4.55,30.87,.7,'团队工作域 / 概念设计原则',13,'BCC7BB')
    for i,(title,body) in enumerate([('资料层','先确认权限，\n再调用资料。'),('推理层','给出引用，\n保留上下文。'),('行动层','关键动作，\n由人确认。')]):
        xx=1.5+i*10.5
        shape(s,f'trust-number-{i}',xx,6.65,9.7,1.0,f'0{i+1}',18,ACCENT)
        shape(s,f'trust-title-{i}',xx,8.2,9.7,1.65,title,31,PAPER,display=True)
        shape(s,f'trust-body-{i}',xx,10.65,9.7,2.7,body,20,'D7DED4')
        if i<2:shape(s,f'trust-arrow-{i}',xx+8.9,8.65,.7,.5,fill='BCC7BB',geometry='rightArrow')
    footer(s,True);notes(s,'这是产品设计原则，不声称已交付某种安全认证。按资料、推理、行动三层解释权限、引用和人工确认的边界。')

    # 8 — A single accent ground, distinct composition and a concrete next step.
    s=d.slide(ACCENT);label(s,'03 / 从小范围开始')
    shape(s,'pilot-statement',1.5,3.2,18.4,6.7,'接一组资料，\n验证一个工作流。',44,display=True,spacing='-0.4pt')
    for i,txt in enumerate(['一个团队','一组内部资料','一个真实问题']):
        shape(s,f'pilot-item-{i}',23.0,4.0+i*2.2,9.3,1.4,txt,22)
    shape(s,'pilot-rule',1.5,12.5,30.87,.035,fill='7C8C44')
    for i,(title,week) in enumerate([('界定','第1周'),('连接','第2周'),('试用','第3周'),('复盘','第4周')]):
        xx=1.5+i*7.9;shape(s,f'pilot-stage-{i}',xx,13.35,7.0,1.1,title,26,display=True)
        shape(s,f'pilot-week-{i}',xx,14.9,7.0,.85,week,15)
    footer(s);notes(s,'建议性四周试用安排，不是已签约项目。先选择一个团队、一组有权使用的资料和一个真实问题，再检验答案正确性、引用可追溯性与交接体验。')

    # 9 — Close on the intended action, holding the final state.
    s=d.slide(INK);label(s,'循迹 TRACE',True)
    shape(s,'closing-statement',1.5,3.1,30.87,5.8,'把下一次搜索，\n变成工作的起点。',49,PAPER,display=True,spacing='-0.4pt')
    shape(s,'closing-action',1.5,11.1,30.87,1.7,'选一个问题，开始小范围试用。',23,ACCENT)
    shape(s,'closing-summary',1.5,14.0,30.87,1.0,'找到证据   /   保留上下文   /   推进行动',18,'D7DED4')
    footer(s,True);notes(s,'停留在结尾，邀请讨论适合验证的具体问题。循迹是原创概念设计，界面与数值用于演示，不提供真实购买或注册入口。')
    for i in range(1,10):
        if i!=4:d.run('set',f'/slide[{i}]',props={'transition':'fade-350','advanceClick':True})
    d.finish()
    timing={'slide':6,'shape':'evidence-chart','effect':1,'duration':.45,'delay':0,'accelerate':0,'decelerate':0}
    Path(args.out).with_suffix('.timing.json').write_text(json.dumps(timing),encoding='utf-8')

if __name__=='__main__':main()
