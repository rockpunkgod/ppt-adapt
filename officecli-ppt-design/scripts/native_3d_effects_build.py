"""Build an editable native 3D interaction page using an existing local GLB."""
import argparse
import json
from pathlib import Path
import subprocess

p = argparse.ArgumentParser(description=__doc__)
p.add_argument('--officecli', required=True)
p.add_argument('--model', required=True)
p.add_argument('--out', required=True)
a = p.parse_args()
out = Path(a.out).resolve()
model = Path(a.model).resolve()
if out.exists() or not model.is_file():
    raise ValueError('Use an existing local GLB and a fresh output path')


def run(*args):
    result = subprocess.run([a.officecli, *map(str, args)], capture_output=True, text=True, encoding='utf-8')
    if result.returncode:
        raise RuntimeError(result.stdout + result.stderr)
    return result.stdout


out.parent.mkdir(parents=True, exist_ok=True)
run('create', out)
run('set', out, '/', '--prop', 'slideWidth=33.87cm', '--prop', 'slideHeight=19.05cm')
ops = [{'command': 'add', 'parent': '/', 'type': 'slide', 'props': {'layout': 'blank', 'background': '16251D'}}]
for name, text, x, y, w, h, font, size, color in [
    ('kicker', 'TRACE / 原生3D动作', 1.5, 1.3, 28, .7, 'Noto Sans SC', 16, 'D7EE6D'),
    ('title', '让物体，\n回应每一次点击。', 1.5, 5.5, 16, 4.5, 'Noto Serif SC', 34, 'F4F4ED'),
    ('instruction', '单击空白开始。\n持续旋转时，点击模型。', 1.5, 11.8, 16, 2, 'Noto Sans SC', 19, 'BEC8BC'),
    ('footer', '3D入场 / 持续转台 / 点击跳跃转身', 1.5, 17.4, 28, .7, 'Noto Sans SC', 14, 'BEC8BC'),
]:
    ops.append({'command': 'add', 'parent': '/slide[1]', 'type': 'shape', 'props': {
        'name': name, 'text': text, 'x': f'{x}cm', 'y': f'{y}cm', 'width': f'{w}cm', 'height': f'{h}cm',
        'font': font, 'size': str(size), 'color': color, 'fill': 'none', 'line': 'none',
    }})
ops.append({'command': 'add', 'parent': '/slide[1]', 'type': 'model3d', 'props': {
    'name': 'interactive-model', 'src': str(model), 'x': '20cm', 'y': '5.5cm',
    'width': '6cm', 'height': '10cm', 'rotation': '0,20,0',
}})
batch = out.with_suffix('.batch.json')
batch.write_text(json.dumps(ops, ensure_ascii=False), encoding='utf-8')
result = json.loads(run('batch', out, '--input', batch, '--json'))
if not result['success']:
    raise RuntimeError(result)
run('close', out)
print(out)
