"""Render this example's approved plan using existing OfficeCLI recipe primitives."""
import argparse
import json
import sys
import hashlib
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'officecli-ppt-design' / 'scripts'))
from build_components import Deck
sys.path.insert(0, str(ROOT / 'ppt-composition-planner' / 'scripts'))
from plan_composition import plan_digest, verify_source_files


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--officecli', required=True)
    ap.add_argument('--out', type=Path, required=True)
    ap.add_argument('--composition', type=Path, required=True)
    args = ap.parse_args()
    folder = Path(__file__).resolve().parent
    load = lambda name: json.loads((folder / name).read_text(encoding='utf-8-sig'))
    spec, plan, data = map(load, ['design-spec.json', 'deck-plan.json', 'data.json'])
    comp = json.loads(args.composition.read_text(encoding='utf-8-sig'))
    if comp['status'] != 'planned':
        raise ValueError('Composition must pass planning before authoring')
    if plan_digest(comp) != comp.get('plan_sha256'):
        raise ValueError('Composition payload changed; replan before authoring')
    verify_source_files(comp)
    digest = lambda value: hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(',', ':')).encode('utf-8')).hexdigest()
    if comp['input_hashes'] != {'spec': digest(spec), 'deck_plan': digest(plan)}:
        raise ValueError('Approved composition inputs changed; replan before authoring')
    if comp['index']['sha256'] != hashlib.sha256((ROOT / 'officecli-ppt-design/references/recipe-index.md').read_bytes()).hexdigest():
        raise ValueError('Recipe index changed; replan before authoring')
    for evidence in comp.get('source_files', {}).get('evidence', []):
        if evidence.get('local') and hashlib.sha256(Path(evidence['path']).read_bytes()).hexdigest() != evidence['sha256']:
            raise ValueError('Evidence changed; replan before authoring: ' + evidence['id'])
    spec, plan = comp['effective_spec'], comp['effective_deck_plan']
    colors = spec['tokens']['colors']
    fonts = spec['tokens']['fonts']
    d = Deck(args.out, args.officecli)
    d.run('set', '/', props={
        'theme.font.major.eastAsia': fonts['heading']['ea'],
        'theme.font.minor.eastAsia': fonts['body']['ea'],
        'theme.font.major.latin': fonts['heading']['latin'], 'theme.font.minor.latin': fonts['body']['latin'],
        'theme.color.accent1': colors['accent1'], 'theme.color.accent2': colors['accent2']})
    color = lambda role: role if role in ('accent1', 'accent2') else colors[role]
    for page in plan['pages']:
        s = d.slide(colors['night'] if page['variant'] == 'night' else colors['paper'])
        for obj in page['objects']:
            name = obj['native_name']
            x, y, w, h = obj['bounds_cm']
            if obj['kind'] in ('text', 'shape'):
                role = obj.get('font_role', 'body')
                props = {'geometry': obj.get('geometry', 'rect')}
                if obj.get('align'):
                    props['align'] = obj['align']
                if obj.get('fill_role'):
                    props['fill'] = color(obj['fill_role'])
                if obj.get('text'):
                    props.update(size=obj.get('font_size_pt', obj.get('size_pt', spec['tokens']['font_sizes_pt'][role])),
                                 color=color(obj['color_role']), bold=role == 'title',
                                 **{'font.ea': '+mj-ea' if role == 'title' else '+mn-ea',
                                    'font.latin': '+mj-lt' if role == 'title' else '+mn-lt'})
                d.shape(s, name, x, y, w, h, obj.get('text', ''), **props)
            elif obj['kind'] == 'chart':
                chart = data['demand']
                d.run('add', f'/slide[{s}]', '--type', 'chart', props={
                    'name': name, 'chartType': 'column',
                    'categories': ','.join(chart['categories']),
                    'data': ';'.join(series['label'] + ':' + ','.join(map(str, series['values'])) for series in chart['series']),
                    'x': f'{x}cm', 'y': f'{y}cm', 'width': f'{w}cm', 'height': f'{h}cm',
                    'title': 'none', 'legend': 'bottom', 'legendFont': f"16:{colors['ink']}:{fonts['body']['ea']}",
                    'axisfont': f"16:{colors['muted']}:{fonts['body']['ea']}", 'axismin': '0', 'axismax': str(obj['axis']['maximum']),
                    'majorunit': '20', 'chartareafill': 'none', 'chartborder': 'none',
                    'plotareafill': 'none', 'plotborder': 'none', 'gridlines': 'D5DBD9:0.5',
                    'colors': colors['accent2'] + ',' + colors['accent1'], 'overlap': '0'})
                d.run('add', f'/slide[{s}]/chart[1]', '--type', 'animation', props={
                    'effect': 'fade', 'class': 'entrance', 'chartBuild': 'category',
                    'trigger': 'onClick', 'duration': '500'})
            elif obj['kind'] == 'table':
                rows = [data['comparison']['columns'], *data['comparison']['rows']]
                d.run('add', f'/slide[{s}]', '--type', 'table', props={
                    'name': name, 'data': ';'.join(','.join(row) for row in rows),
                    'x': f'{x}cm', 'y': f'{y}cm', 'width': f'{w}cm', 'height': f'{h}cm',
                    'colWidths': '5cm,4cm,7.5cm,6.3cm,8.07cm', 'style': 'none',
                    'firstRow': False, 'lastRow': False, 'bandedRows': False, 'border.all': 'none'})
                ops = []
                for r, row in enumerate(rows, 1):
                    for c in range(1, 6):
                        props = {'font': fonts['body']['ea'], 'font.ea': '+mn-ea', 'font.latin': '+mn-lt',
                                 'size': str(spec['tokens']['font_sizes_pt']['body']), 'valign': 'center', 'align': 'left', 'padding': '0.18cm',
                                 'fill': colors['paper'], 'color': colors['ink'], 'border.all': 'none',
                                 'border.bottom': '0.5pt solid D5DBD9', 'bold': str(r in (1, 4, 6)).lower()}
                        if r == 1:
                            props.update(fill='accent1', color=colors['on_dark'])
                        ops.append({'command': 'set', 'path': f'/slide[{s}]/table[1]/tr[{r}]/tc[{c}]', 'props': props})
                d.batch(ops)
        prefix = plan['deck_id'] + '.' + page['id']
        d.shape(s, prefix + '.footer-band', 0, 18.15, 33.87, .9, fill=colors['paper'])
        d.shape(s, prefix + '.footer', 1.5, 18.24, 29, .5,
                spec['shared_components']['footer']['label'] + '  |  ' + spec['layout_rules']['footer_text'],
                size=spec['tokens']['font_sizes_pt']['footer'], color=colors['ink'], **{'font.ea': '+mn-ea', 'font.latin': '+mn-lt'})
        d.shape(s, prefix + '.page-number', 31.2, 18.24, 1.1, .5, f'{s:02}',
                size=spec['tokens']['font_sizes_pt']['footer'], color=colors['ink'],
                **{'font.ea': '+mn-ea', 'font.latin': '+mn-lt'})
        if page['motion']['transition'] == 'morph':
            d.run('set', f'/slide[{s}]', props={'transition': f"morph-{page['motion']['duration_ms']}", 'advanceClick': True})
        d.notes(s, page['message'] + '\n' + data['disclaimer'])
    # CLI adds !! to all shapes while enabling Morph. Preserve only planned pairs.
    for page in plan['pages']:
        if page['id'] not in ('s05', 's06'):
            continue
        ordinary = [o['native_name'] for o in page['objects'] if not o.get('morph_key')]
        prefix = plan['deck_id'] + '.' + page['id']
        ordinary += [prefix + '.' + suffix for suffix in ('footer-band', 'footer', 'page-number')]
        for name in ordinary:
            d.run('raw-set', f"/slide[{page['order']}]", '--xpath',
                  f"//p:cNvPr[@name='!!{name}']", '--action', 'setattr', '--xml', 'name=' + name)
    d.run('close')
    # Bind chart series to the same semantic theme slots as the page shapes.
    with zipfile.ZipFile(d.file) as package:
        chart_parts = [n for n in package.namelist() if '/charts/' in n and n.endswith('.xml')]
    if len(chart_parts) != 1:
        raise ValueError('This example requires exactly one native chart')
    for series, slot in ((1, 'accent2'), (2, 'accent1')):
        d.run('raw-set', '/' + chart_parts[0], '--xpath', f'//c:ser[{series}]/c:spPr/a:solidFill',
              '--action', 'replace', '--xml', '<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
              f'<a:schemeClr val="{slot}"/></a:solidFill>')
    d.finish()
    timing = [{'slide': 3, 'shape': plan['deck_id'] + '.s03.chart', 'effect': n,
               'duration': .5, 'delay': 0, 'accelerate': 0, 'decelerate': 0} for n in range(1, 5)]
    # Existing PS5 helper accepts a JSON object reliably; CLI authored all four at 500ms.
    args.out.with_suffix('.timing-plan.json').write_text(json.dumps(timing[0], indent=2), encoding='utf-8')


if __name__ == '__main__':
    main()
