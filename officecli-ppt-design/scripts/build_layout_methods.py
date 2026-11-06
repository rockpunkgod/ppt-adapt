"""OfficeCLI layout lab. Uses only stdlib and sibling build_components.py.

Reconstructs transferable operations from BV1oPHgeiEgb without source artwork.
All content is editable; sample values are illustrative, not business data.
"""
import argparse
from pathlib import Path
from build_components import Deck

INK='172A35'; MUTED='52636D'; ACCENT='E95537'; PAPER='F4F0E7'

def label(d,s,name,x,y,w,h,text,size=20,**p):
    d.shape(s,name,x,y,w,h,text,size=size,**p)

def rect(d,s,name,x,y,w,h,fill,**p):
    d.shape(s,name,x,y,w,h,fill=fill,**p)

def head(d,s,title,source):
    label(d,s,'title',1.4,.9,30,1.5,title,32,bold=True)
    label(d,s,'source',1.45,17.7,30,.7,source+' · 机制重建 / 示例内容',12,color=MUTED)

def build(out,exe):
    d=Deck(out,exe,static=True)
    # 1. Frame and grid before decorating; shared edge and baseline.
    s=d.slide(PAPER);head(d,s,'网格先行，信息各有位置','50P · P03 417–758s')
    x0,y0,cw,ch=2,4.1,4.98,2.06
    for i in range(7):
        rect(d,s,f'grid-v-{i}',x0+i*cw,y0,.016,ch*5,'D4D3CD')
    for i in range(6):
        rect(d,s,f'grid-h-{i}',x0,y0+i*ch,cw*6,.016,'D4D3CD')
    label(d,s,'grid-cn',x0,y0,9,2.6,'神奇的\n格子',26,bold=True)
    label(d,s,'grid-en',x0,y0+ch*2,12,4,'MAGICAL\nGRID',42,bold=True)
    label(d,s,'grid-G',x0+cw*3,y0+ch,14.8,9,'G',210,bold=True,font='Arial')
    label(d,s,'grid-anchor',x0,y0+ch*4.55,6,.8,'01  /  STRUCTURE',14)
    d.notes(s,'源P03网格字形对齐：先版心后网格，再放标题和大G。本例保留网格方便验证，交付稿可隐藏。所有文字是原生文本，未转轮廓，不宣称字形路径编辑。')
    # 2. Same copy, three spacing levels, readable slides rather than poster tiny type.
    s=d.slide(PAPER);head(d,s,'关联成组，留白分级','50P · P05 596–716s')
    rect(d,s,'paper',11.2,3.1,12,13.5,'FFFFFF')
    label(d,s,'kicker',12.1,3.8,10.2,.8,'SUMMER / NEW COLLECTION',15,align='center',color=MUTED)
    label(d,s,'summer',12,5.2,10.4,2.1,'夏季 × 新品',36,bold=True,align='center')
    label(d,s,'offer',13.8,8.1,6.8,1,'探索新的生活方式',20,align='center')
    label(d,s,'sale',13.8,11.4,6.8,1.4,'S A L E',26,bold=True,align='center')
    label(d,s,'copy',12,13.5,10.4,2,'材质、色彩与日常\n每一处细节，都值得留意',18,align='center',color=MUTED)
    for i,(y,r,t) in enumerate([(7.4,.2,'组内'),(9.9,.4,'组间'),(12.7,.6,'段落')]):
        d.shape(s,f'space-{i}',26,y,r*2,r*2,preset='ellipse',fill=ACCENT)
        label(d,s,f'space-label-{i}',27.6,y-.2,4,1,t,18)
    d.notes(s,'P05三档圆形间距尺：本例只验证三档相对关系，非固定行距比例。右侧圆为演示标尺，正式页删除。')
    # 3. Balance is not compulsory centering.
    s=d.slide(PAPER);head(d,s,'以视觉重量平衡两侧','50P · P06 537–1173s')
    label(d,s,'primary',2.1,5.4,17,5.8,'50%',138,bold=True,font='Arial')
    label(d,s,'secondary',21,6.0,10.7,1.7,'让重点先被看见',30,bold=True)
    label(d,s,'detail',21,8.6,10.7,3.3,'数字承担第一视觉\n说明沿共享左边线排列\n用距离调整画面的重心',21,color=MUTED)
    rect(d,s,'rule',21,13.4,10.7,.04,ACCENT)
    label(d,s,'context',2.3,13,15,2,'示意比例，不代表真实统计',18,color=MUTED)
    d.notes(s,'源P06视觉重量与基线调整。画面左右并不等宽、等字号；原生预览应先看到50%，其次标题，最后说明。')
    # 4. Static visual routing; circles and connector outlines remain editable.
    s=d.slide(PAPER);head(d,s,'点定义焦点，线串联视线','50P · P07–08 315–606s')
    for i,(x,y,w,h) in enumerate([(8.4,4.1,16,12),(10.2,5.3,12.4,9.6),(12.1,6.7,8.6,6.7)]):
        d.shape(s,f'orbit-{i}',x,y,w,h,preset='ellipse',fill='none',line='A5AAA7',lineWidth='0.7pt')
    for i,(x,y,r) in enumerate([(14.1,8.6,2.4),(9.3,6.4,.8),(21.6,7.2,.6),(12.3,13.3,.5),(23.2,12.5,.35)]):
        d.shape(s,f'node-{i}',x,y,r*2,r*2,preset='ellipse',fill=INK if i!=1 else ACCENT)
    label(d,s,'route-primary',2.1,5.9,6,2,'主次与路线',26,bold=True)
    label(d,s,'route-caption',2.1,8.6,5.7,4,'尺度形成主次\n轮廓连接分散元素\n控制节点数量',19,color=MUTED)
    d.notes(s,'P08静态视觉路线重建，不是轨道动画或数据关系图。原生椭圆与节点可编辑；这些轮廓没有语义连接关系。')
    # 5. Figure-ground and canvas crop use native elements fully inside slide.
    s=d.slide(INK);head(d,s,'先搭大面，再安放次级信息','50P · P09 191–486s')
    d.set(s,'title',color='FFFFFF');d.set(s,'source',color='D3DCD8')
    rect(d,s,'main-surface',2,4,18.1,11.7,'E95537')
    d.shape(s,'counter-circle',14,5.7,9,9,preset='ellipse',fill=PAPER)
    label(d,s,'main-word',3,5.3,14,5.5,'可见的\n秩序',62,bold=True,color='FFFFFF')
    label(d,s,'support',24.5,6.1,7,1.2,'FORM / 05',19,color='FFFFFF')
    label(d,s,'body',24.5,8.3,7,5,'大面积建立框架\n小面积承载解释\n先结构，后装饰',20,color='D3DCD8')
    d.notes(s,'P09点线面构图，主面与次面叠放。未执行布尔相交；此页不冒充人物抠图或自由曲面剪切。')
    # 6. Same brief can prioritize date, message, or identity.
    s=d.slide(PAPER);head(d,s,'同一份文案，先决定强调什么','50P · P45 53–636s')
    for i in range(3):
        x=1.8+10.4*i;rect(d,s,f'variant-{i}',x,4,9.7,11.8,'FFFFFF')
    label(d,s,'v1-year',2.4,5.2,8.5,2.4,'2026',51,bold=True,font='Arial')
    label(d,s,'v1-body',2.4,9,8.4,4,'设计基地\n学门好手艺\n练习 · 观察 · 修正',21)
    label(d,s,'v2-title',12.9,5.2,8.4,4.3,'设计\n基地',44,bold=True)
    label(d,s,'v2-body',12.9,10.1,8.4,3,'2026 / 开放练习\n学门好手艺',20)
    label(d,s,'v3-logo',23.2,5.1,8.5,3.6,'SD',84,bold=True,font='Arial',color=ACCENT)
    label(d,s,'v3-body',23.2,10,8.4,3.5,'设计基地 / 2026\n学门好手艺',20)
    d.notes(s,'源P45：改变主次目标产生变体，而非随机换位置。三栏同一文案有轻微排布差异，均保持原生可编辑与阅读顺序。')
    # 7. Semantic direction, without fabricating numerical growth.
    s=d.slide(PAPER);head(d,s,'用方向表达主题，不制造数据','50P · P46 438–526s')
    label(d,s,'knowledge',2.2,4.0,17,3.4,'知识',67,bold=True)
    label(d,s,'change',7.6,7.6,20,3,'改变命运',65,bold=True)
    rect(d,s,'growth-line',6.3,8.0,.06,7.6,ACCENT)
    d.shape(s,'growth-arrow',5.9,7.3,.86,.9,preset='triangle',fill=ACCENT)
    for i in range(6):rect(d,s,f'tick-{i}',6.3,8.5+i*1.35,.5 if i%2 else .9,.035,ACCENT)
    label(d,s,'explain',9,12.5,20,2.6,'将成长的语义转成向上的方向\n刻度仅为装饰，不表示测量值',22,color=MUTED)
    d.notes(s,'源P46语义箭头和刻度。此页刻度不对应数值，文本明确注明，禁止用于真实增长图表替代。')
    # 8. Repetition creates modules; faint background never carries deadline.
    s=d.slide(PAPER);head(d,s,'重复模块，统一节奏','50P · P47 438–957s')
    label(d,s,'background-year',13,3.6,19,3.8,'2026',92,bold=True,font='Arial',color='DDDCD4')
    label(d,s,'collection-title',2.1,4.1,18,3.1,'年度作品征集',48,bold=True)
    label(d,s,'deadline',2.1,8,15,1.1,'截止日期：2026.12.01（示例）',22)
    for i,(cn,en) in enumerate([('视觉设计','VISUAL DESIGN'),('产品设计','PRODUCT DESIGN'),('空间设计','SPATIAL DESIGN')]):
        x=2.1+10.2*i;rect(d,s,f'module-rule-{i}',x,11,9.1,.035,INK)
        d.shape(s,f'module-dot-{i}',x,11.7,.48,.48,preset='ellipse',fill=ACCENT)
        label(d,s,f'module-title-{i}',x,12.6,9.1,1.2,cn,25,bold=True)
        label(d,s,f'module-en-{i}',x,14.2,9.1,1.2,en,13,color=MUTED)
    d.notes(s,'源P47重复中英文类别模块、弱化年份装饰。截止日期单独用高对比文本呈现，不隐藏在背景年号中。')
    d.run('validate');d.run('save');return d.file

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--out',required=True);p.add_argument('--officecli',default='officecli');a=p.parse_args()
    print(build(Path(a.out),a.officecli))
