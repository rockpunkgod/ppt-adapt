"""Editable OfficeCLI mechanism examples; Python stdlib only.

Use --recipe for one example. Existing output is never overwritten.
The sample is a reconstruction of a mechanism, not a copy of source artwork.
"""
import argparse
import json
import math
from pathlib import Path
import shutil
import subprocess
import tempfile

W, H = 33.87, 19.05
FONT = 'Microsoft YaHei'


class Deck:
    def __init__(self, output, executable, static=False):
        self.file = Path(output).resolve()
        if self.file.exists():
            raise FileExistsError(f'Refusing to overwrite: {self.file}')
        self.file.parent.mkdir(parents=True, exist_ok=True)
        self.exe, self.static, self.count = executable, static, 0
        self.log = []
        self.run('create')
        self.run('set', '/', props={'slideWidth': f'{W}cm', 'slideHeight': f'{H}cm',
                 'theme.font.major.eastAsia': FONT, 'theme.font.minor.eastAsia': FONT})

    def run(self, verb, *args, props=None):
        cmd = [self.exe, verb, str(self.file), *map(str, args)]
        for key, value in (props or {}).items():
            cmd += ['--prop', f'{key}={str(value).lower() if isinstance(value, bool) else value}']
        cmd += ['--json']
        r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
        self.log.append({'verb': verb, 'args': list(map(str, args)), 'props': props,
                         'returncode': r.returncode, 'stdout': r.stdout, 'stderr': r.stderr})
        if r.returncode:
            raise RuntimeError(f'{verb} failed: {r.stdout}\n{r.stderr}')
        try:
            data = json.loads(r.stdout)
        except json.JSONDecodeError:
            return r.stdout
        if data.get('success') is False:
            raise RuntimeError(data)
        return data

    def batch(self, commands):
        # An actual UTF-8 file avoids shell quoting of Chinese, !!, $, and newlines.
        with tempfile.TemporaryDirectory(prefix='ppt-design-') as folder:
            path = Path(folder) / 'commands.json'
            path.write_text(json.dumps(commands, ensure_ascii=False), encoding='utf-8')
            result = self.run('batch', '--input', path)
            if '"atomicRolledBack": true' in json.dumps(result):
                raise RuntimeError('OfficeCLI batch rolled back')
            return result

    def slide(self, background='F3F6F8'):
        self.count += 1
        self.run('add', '/', '--type', 'slide', props={'layout': 'blank', 'background': background})
        return self.count

    def clone(self, source):
        self.run('add', '/', '--from', f'/slide[{source}]')
        self.count += 1
        return self.count

    def shape(self, s, name, x, y, width, height, text='', **props):
        p = {'name': name, 'x': f'{x:.4f}cm', 'y': f'{y:.4f}cm',
             'width': f'{width:.4f}cm', 'height': f'{height:.4f}cm',
             'fill': 'none', 'line': 'none'}
        if text:
            p.update(text=text, font=FONT, size=20, color='172A35', margin='0cm',
                     valign='center', autoFit='none')
        p.update(props)
        self.run('add', f'/slide[{s}]', '--type', 'shape', props=p)

    def set(self, s, name, **props):
        self.run('set', f'/slide[{s}]/shape[@name={name}]', props=props)

    def morph(self, s):
        if not self.static:
            self.run('set', f'/slide[{s}]', props={'transition': 'morph-700', 'advanceClick': True})

    def heading(self, s, title, subtitle):
        self.shape(s, f'heading-{s}', 1.5, 1.25, 30.8, 1.8, title, size=36, bold=True)
        self.shape(s, f'caption-{s}', 1.55, 3.2, 30, 1.0, subtitle, size=18, color='52636D')

    def notes(self, s, text):
        self.run('add', f'/slide[{s}]', '--type', 'notes', props={'text': text})

    def finish(self):
        self.run('close')
        validation = self.run('validate')
        issues = self.run('view', 'issues')
        self.file.with_suffix('.build-log.json').write_text(
            json.dumps(self.log, ensure_ascii=False, indent=2), encoding='utf-8')
        self.file.with_suffix('.checks.json').write_text(
            json.dumps({'schema': validation, 'issues': issues,
                        'runtime_motion': 'not_verified_by_this_script'}, ensure_ascii=False, indent=2),
            encoding='utf-8')
        print(f'Created {self.file} ({self.count} slides); motion needs target-player verification.')


def cards(d):
    s = d.slide()
    d.heading(s, '把并列内容，展开成清晰目录', '卡片尺寸保持不变；背景、标题和正文作为一组平移。')
    colors = ['224638', '355F4A', '517C62', '709A79']
    titles = ['识别问题', '建立方案', '验证结果', '推进落地']
    body = [('明确目标', '界定需要解决的矛盾'), ('比较路径', '说明选择方案的依据'),
            ('核对证据', '用事实确认方案有效'), ('安排下一步', '写清责任与行动顺序')]
    # Build a fully readable end-state first, then clone and collapse it.
    for i in range(4):
        x = 1.6 + i * 7.2
        d.shape(s, f'!!card-{i}-bg', x, 5.25, 8.6, 11.8, fill=colors[i],
                geometry='roundRect', shadow='000000-5-45-2-18')
        d.shape(s, f'!!card-{i}-number', x+.65, 6.0, 5.8, 2.6, f'{i+1:02}',
                size=64, bold=True, color='FFFFFF')
        d.shape(s, f'!!card-{i}-title', x+.7, 9.5, 5.7, 1.15, titles[i],
                size=26, bold=True, color='FFFFFF')
        for j, line in enumerate(body[i]):
            d.shape(s, f'!!card-{i}-line-{j}', x+.7, 11.6+j*1.3, 5.8, 1.0,
                    line, size=18, color='FFFFFF')
    # The returned path after grouping is queried by stable group name, not shape indices.
    for i in range(4):
        names = [f'!!card-{i}-bg', f'!!card-{i}-number', f'!!card-{i}-title',
                 f'!!card-{i}-line-0', f'!!card-{i}-line-1']
        paths = ','.join(f'/slide[{s}]/shape[@name={n}]' for n in names)
        d.run('add', f'/slide[{s}]', '--type', 'group', props={'name': f'!!card-{i}', 'shapes': paths})
    collapsed = d.clone(s)
    for i in range(4):
        d.run('set', f'/slide[{collapsed}]/group[@name=!!card-{i}]', props={'x': f'{23+i*.35}cm'})
    # Order: collapsed -> expanded. move is a real slide-tree operation.
    d.run('move', f'/slide[{collapsed}]', '--before', f'/slide[{s}]')
    d.morph(s+1)
    for page in (s, s+1):
        d.notes(page, 'R01 卡片抽拉。原作包含数字与卡片布尔结合；本例保留可编辑文本，未复刻数字镂空轮廓。')


def focus(d):
    s = d.slide('172A35')
    d.shape(s, f'focus-heading-{s}', 1.5, 1.35, 30, 1.8, '一次，只把一个重点讲清楚',
            size=36, bold=True, color='FFFFFF')
    for i, title in enumerate(['事实', '判断', '行动']):
        x = 2.1+i*10.1
        d.shape(s, f'!!focus-card-{i}', x, 6.4, 9.1, 8.9, fill=['355F4A','174A64','6A5140'][i], geometry='roundRect')
        d.shape(s, f'!!focus-title-{i}', x+.7, 8.0, 7.7, 1.7, title, size=34, bold=True, color='FFFFFF')
        d.shape(s, f'!!focus-body-{i}', x+.7, 11.2, 7.7, 1.4,
                ['先呈现证据','再解释含义','最后安排下一步'][i], size=20, color='FFFFFF')
    # Four rectangles make a real rectangular window. No opaque circle is passed off as a hole.
    def mask(page, x):
        y, w, h = 6.1, 9.7, 9.5
        panels = {'top': (0,4.8,W,y-4.8), 'bottom': (0,y+h,W,17.5-y-h),
                  'left': (0,y,x,h), 'right': (x+w,y,W-x-w,h)}
        for name, rect in panels.items():
            if page==s:
                d.shape(page, f'!!mask-{name}', *rect, fill='10191F', opacity=.83)
            else:
                d.set(page, f'!!mask-{name}', x=f'{rect[0]}cm', y=f'{rect[1]}cm', width=f'{rect[2]}cm',height=f'{rect[3]}cm')
        if page==s:
            d.shape(page, '!!focus-outline',x,y,w,h,fill='none',line='D7AE66:1.5',geometry='roundRect')
        else:
            d.set(page,'!!focus-outline',x=f'{x}cm')
    mask(s,1.8)
    for i in (1,2):
        t=d.clone(d.count);mask(t,1.8+i*10.1);d.morph(t)
    for page in range(s,s+3):
        d.notes(page,'R10 聚光灯工程实现：四块半透明矩形形成窗口，底层内容固定。原作圆孔以矩形窗口替代，形状差异明确。')


def petals(d):
    s=d.slide('101E24')
    d.shape(s,f'petals-heading-{s}',1.5,1.3,30,1.8,'把共同起点，展开为五条路径',size=36,bold=True,color='FFFFFF')
    # Editable lens outline. High-level shape presets do not express the circle-intersection lens.
    geom='<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="w" b="h"/><a:pathLst><a:path w="100000" h="100000"><a:moveTo><a:pt x="50000" y="0"/></a:moveTo><a:cubicBezTo><a:pt x="116667" y="33333"/><a:pt x="116667" y="66667"/><a:pt x="50000" y="100000"/></a:cubicBezTo><a:cubicBezTo><a:pt x="-16667" y="66667"/><a:pt x="-16667" y="33333"/><a:pt x="50000" y="0"/></a:cubicBezTo><a:close/></a:path></a:pathLst></a:custGeom>'
    for i in range(5):
        name=f'!!petal-{i}'
        d.shape(s,name,14.7,7.3,4.4,8.4,fill=['23664B','2B8760','36A879','439B91','367987'][i],opacity=.78)
        xpath=f'//p:sp[p:nvSpPr/p:cNvPr[@name="{name}"]]/p:spPr/a:prstGeom'
        d.run('raw-set',f'/slide[{s}]','--xpath',xpath,'--action','replace','--xml',geom)
    t=d.clone(s)
    pivot_x,pivot_y=16.9,15.7
    for i,angle in enumerate([-48,-24,0,24,48]):
        # PowerPoint rotates about bounding-box centre. Move the centre so all tips share a pivot.
        theta=math.radians(angle);r=4.2
        cx=pivot_x+r*math.sin(theta);cy=pivot_y-r*math.cos(theta)
        d.set(t,f'!!petal-{i}',x=f'{cx-2.2}cm',y=f'{cy-4.2}cm',rotation=angle)
        label_x=pivot_x+10*math.sin(theta)-2.2;label_y=pivot_y-10*math.cos(theta)-1.3
        d.shape(t,f'petal-label-{i}',label_x,label_y,4.4,1.2,
                ['目标','证据','方案','验证','行动'][i],size=22,bold=True,color='FFFFFF',align='center')
    d.morph(t)
    for page in (s,t):d.notes(page,'R05 花瓣展开。自定义原生透镜形替代GUI圆形相交，终态共用底部支点；实际Morph弧线需实播检查。')


def palette(d):
    s=d.slide('F3F6F8');d.heading(s,'配色的关键，是颜色之间的关系','选色系 → 定色值 → 调比例；先保证持续阅读，再营造氛围。')
    swatches=[(1.6,18.3,'E5EEF2','底色 / 辅色','留出大面积阅读空间','172A35'),
              (20.3,7.8,'174A64','主色','建立页面结构','FFFFFF'),
              (28.5,3.8,'C28A45','点缀','强调重点','172A35')]
    for i,(x,w,color,title,body,fg) in enumerate(swatches):
        d.shape(s,f'palette-{i}',x,6,w,10.4,fill=color,geometry='roundRect')
        d.shape(s,f'palette-title-{i}',x+.45,7.4,w-.9,1.4,title,size=24 if i<2 else 20,bold=True,color=fg)
        if i<2:d.shape(s,f'palette-body-{i}',x+.45,10,w-.9,1.5,body,size=20,color=fg)
        d.shape(s,f'palette-hex-{i}',x+.45,14.2,w-.9,1,color,size=16,color=fg)
    d.notes(s,'参考《审美无平替，但有方法。》：主色不一定占最大面积，点缀色少量使用。所示色板为本例自拟。')


def timeline(d):
    s=d.slide();d.heading(s,'让年份变化，与事件同步','固定文字槽位；移动数字条，用背景色遮罩控制窗口。')
    d.shape(s,'!!timeline-panel',2,6,11.4,9.8,fill='174A64',geometry='rect')
    for col,digits in enumerate(['2','0','2','345']):
        # Separate digits make fixed-width columns without depending on multiline line-height.
        for j,num in enumerate(digits):
            d.shape(s,f'!!digit-{col}-{j}',3+col*2.2,8+j*3.1,2.1,2.6,num,
                    font='Arial',size=70,bold=True,color='FFFFFF',align='center')
    # These opaque masks clip the rolling units column visually, not through crop-only metadata.
    d.shape(s,'!!digit-mask-top',9.4,6.0,3.5,2.0,fill='174A64')
    d.shape(s,'!!digit-mask-bottom',9.4,10.6,3.5,5.2,fill='174A64')
    d.shape(s,'!!digit-outer-top',9.4,0,3.5,6.0,fill='F3F6F8')
    d.shape(s,'!!digit-outer-bottom',9.4,15.8,3.5,H-15.8,fill='F3F6F8')
    d.shape(s,'!!year-event-title',16,7.7,15.5,1.7,'2023：建立方案',size=30,bold=True)
    d.shape(s,'!!year-event-body',16,10.2,15.5,2.5,'把问题转化为可执行的方案。',size=22)
    for step,(title,body) in enumerate([('2024：验证结果','用真实结果检查方法有效。'),('2025：推进落地','明确下一阶段的行动。')],1):
        t=d.clone(d.count)
        for j in range(3):d.set(t,f'!!digit-3-{j}',y=f'{8+(j-step)*3.1}cm')
        d.set(t,'!!year-event-title',text=title);d.set(t,'!!year-event-body',text=body);d.morph(t)
    for page in range(s,s+3):
        # Outside-window masks must not erase title glyphs. Put editorial text above them.
        d.run('remove',f'/slide[{page}]/shape[@name=heading-{s}]')
        d.run('remove',f'/slide[{page}]/shape[@name=caption-{s}]')
        d.heading(page,'让年份变化，与事件同步','固定文字槽位；移动数字条，用背景色遮罩控制窗口。')
        d.notes(page,'R11 数字时间轴。用可编辑数字条和背景色矩形模拟窗口，保持各年事件在同一状态页配对；这些年份为示例内容。')


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--recipe',choices=['cards','focus','petals','palette','timeline','all'],default='cards')
    parser.add_argument('--out',required=True)
    parser.add_argument('--officecli',default=shutil.which('officecli'))
    parser.add_argument('--static',action='store_true')
    args=parser.parse_args()
    if not args.officecli:parser.error('officecli not found; pass --officecli /absolute/path/to/officecli')
    d=Deck(args.out,args.officecli,args.static)
    recipes={'cards':cards,'focus':focus,'petals':petals,'palette':palette,'timeline':timeline}
    try:
        for name in recipes if args.recipe=='all' else [args.recipe]:recipes[name](d)
        d.finish()
    except Exception:
        try:d.run('close')
        except Exception:pass
        raise


if __name__=='__main__':main()
