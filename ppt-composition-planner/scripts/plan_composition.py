"""Plan existing PPT recipes and detect composition conflicts. Standard library only."""
import argparse
import copy
import hashlib
import json
import math
import re
import fnmatch
from pathlib import Path

PRIORITY = ['content correctness and completeness', 'readability', 'global consistency', 'animation', 'decoration']
BACKENDS = {'officecli-native', 'powerpoint-com', 'hybrid', 'external-render'}


def sha(data):
    return hashlib.sha256(data).hexdigest()


def canonical(value):
    return json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(',', ':')).encode('utf-8')


def plan_digest(plan):
    """Verification/status are mutable attestations, not part of the immutable plan."""
    return sha(canonical({k: v for k, v in plan.items() if k not in {'plan_sha256', 'verification', 'status'}}))


def source_files(spec_path, deck_path, spec, deck):
    """CLI boundary only: bind real local evidence relative to deck-plan directory."""
    result = {'spec': {'path': str(spec_path.resolve()), 'sha256': sha(spec_path.read_bytes())},
              'deck_plan': {'path': str(deck_path.resolve()), 'sha256': sha(deck_path.read_bytes())}, 'evidence': []}
    registry = spec.get('evidence', deck.get('evidence', []))
    items = registry if isinstance(registry, list) else [{'id': key, **value} for key, value in registry.items()]
    for item in items:
        if not isinstance(item, dict) or not item.get('path'):
            continue
        declared = item['path']
        if not isinstance(declared, str):
            raise ValueError('Evidence path must be a local path string')
        if re.match(r'^https?://', declared):
            result['evidence'].append({'id': item.get('id'), 'uri': declared, 'local': False, 'note': 'Remote content is not fetched or SHA-bound by this offline planner'})
            continue
        file = Path(declared)
        if not file.is_absolute():
            file = deck_path.resolve().parent / file
        file = file.resolve()
        if not file.is_file():
            raise ValueError('Local evidence file missing: ' + str(file))
        result['evidence'].append({'id': item.get('id'), 'declared_path': declared, 'path': str(file), 'sha256': sha(file.read_bytes()), 'bytes': file.stat().st_size, 'local': True})
    return result


def verify_source_files(plan):
    """Builder calls this before execution; no network, no mutation."""
    registry = plan.get('source_files', {})
    for item in [registry.get('spec'), registry.get('deck_plan'), *registry.get('evidence', [])]:
        if item and item.get('path'):
            if sha(Path(item['path']).read_bytes()) != item['sha256']:
                raise ValueError('Source file changed since planning: ' + item['path'])
    index = plan.get('index')
    if index and sha(Path(index['path']).read_bytes()) != index['sha256']:
        raise ValueError('Recipe index changed since planning')
    return True


def parse_index(path):
    """Read actual published table rows; retain boundaries rather than infer new APIs."""
    raw = Path(path).read_bytes()
    backend = None
    recipes = []
    for line_no, line in enumerate(raw.decode('utf-8-sig').splitlines(), 1):
        if line.startswith('## '):
            candidate = line[3:].strip()
            backend = candidate if candidate in BACKENDS else None
        if backend and line.startswith('|'):
            cells = [c.strip() for c in line.strip().strip('|').split('|')]
            if len(cells) != 3 or cells[0] in {'设计用途', '---'} or set(cells[0]) <= {'-', ':'}:
                continue
            entrypoints = re.findall(r'`([^`]+)`', cells[2])
            refs = re.findall(r'\]\(([^)]+)\)', cells[2])
            if not entrypoints and not refs:
                continue
            recipes.append({'id': f'index-line-{line_no}', 'line': line_no,
                            'purpose': cells[0], 'mechanism': cells[1], 'backend': backend,
                            'entrypoints': entrypoints, 'references': refs, 'scope': cells[2],
                            'base_evidence_status': 'index-listed-with-stated-boundaries'})
    if not recipes:
        raise ValueError('No recipe table rows found in actual index')
    return recipes, sha(raw)


def select_recipe(requirement, recipes):
    query = requirement if isinstance(requirement, str) else requirement.get('query', '')
    pin = None if isinstance(requirement, str) else requirement.get('recipe')
    if not isinstance(query, str) or not query.strip():
        return [], 'Missing mechanism query'
    exact = [r for r in recipes if query == r['mechanism'] or query in r['entrypoints'] or query == r['id']]
    matches = exact or [r for r in recipes if query.casefold() in (r['mechanism'] + ' ' + r['purpose'] + ' ' + ' '.join(r['entrypoints'])).casefold()]
    if pin:
        matches = [r for r in matches if pin == r['id'] or any(fnmatch.fnmatch(pin, e) for e in r['entrypoints'])]
    return matches, None if len(matches) == 1 else ('No existing recipe matches' if not matches else 'Ambiguous recipe; use full mechanism text or index-line id')


def topological(nodes):
    by_id = {n['id']: n for n in nodes}
    if len(by_id) != len(nodes):
        raise ValueError('Duplicate operation id')
    indegree = {key: 0 for key in by_id}
    children = {key: [] for key in by_id}
    for n in nodes:
        for dep in n.get('depends_on', []):
            if dep not in by_id:
                raise ValueError('Unknown operation dependency: ' + str(dep))
            indegree[n['id']] += 1
            children[dep].append(n['id'])
    ready = sorted(k for k, v in indegree.items() if v == 0)
    ordered = []
    while ready:
        key = ready.pop(0)
        ordered.append(key)
        for c in children[key]:
            indegree[c] -= 1
            if not indegree[c]:
                ready.append(c)
                ready.sort()
    if len(ordered) != len(nodes):
        raise ValueError('Operation DAG contains a cycle')
    return ordered


def compile_plan(spec, deck, index):
    spec, deck = copy.deepcopy(spec), copy.deepcopy(deck)
    original = {'spec': sha(canonical(spec)), 'deck_plan': sha(canonical(deck))}
    recipes, index_hash = parse_index(index)
    conflicts, fixes, defaults, selected, timelines, nodes = [], [], [], [], [], []

    def issue(category, location, reason):
        conflicts.append({'category': category, 'location': location, 'reason': reason, 'resolved': False, 'blocking': True})

    def fix(category, location, before, after, reason, priority=3):
        if before is None:
            defaults.append({'category': category, 'location': location, 'value': copy.deepcopy(after), 'reason': reason})
            return
        fixes.append({'category': category, 'location': location, 'before': copy.deepcopy(before),
                      'after': copy.deepcopy(after), 'reason': reason, 'priority': priority,
                      'priority_rule': PRIORITY[priority - 1]})
        conflicts.append({'category': category, 'location': location, 'reason': reason, 'resolved': True, 'blocking': False})

    for key in ['deck_id', 'canvas_cm', 'tokens']:
        if key not in spec:
            issue('schema', 'spec.' + key, 'Required v1 field missing')
    if spec.get('version', 1) != 1 or deck.get('version', 1) != 1:
        issue('schema', 'version', 'Only contract version1 supported')
    tokens = spec.get('tokens', {})
    for key in ['fonts', 'font_sizes_pt', 'colors', 'spacing_cm']:
        if key not in tokens:
            issue('schema', 'tokens.' + key, 'Required v1 token field missing')
    for role in ['heading', 'body']:
        for script in ['ea', 'latin']:
            if not tokens.get('fonts', {}).get(role, {}).get(script):
                issue('font', f'tokens.fonts.{role}.{script}', 'Explicit EA and Latin font references required')
    for role in ['title', 'body', 'caption', 'footer']:
        value = tokens.get('font_sizes_pt', {}).get(role)
        if not isinstance(value, (float, int)) or isinstance(value, bool) or not math.isfinite(value) or value <= 0:
            issue('font', 'tokens.font_sizes_pt.' + role, 'Font size must be positive finite points')
    min_body = spec.get('quality', {}).get('min_body_pt')
    body_size = tokens.get('font_sizes_pt', {}).get('body')
    if isinstance(min_body, (int, float)) and isinstance(body_size, (int, float)) and body_size < min_body:
        issue('readability', 'tokens.font_sizes_pt.body', 'Body token violates declared minimum; revise layout/content with author, never silently shrink')
    canvas = spec.get('canvas_cm', [])
    if isinstance(canvas, dict):
        canvas = [canvas.get('width'), canvas.get('height')]
    if not isinstance(canvas, list) or len(canvas) != 2 or not all(isinstance(x, (int, float)) and not isinstance(x, bool) and math.isfinite(x) and x > 0 for x in canvas):
        issue('bounds', 'canvas_cm', 'Expected [width,height] finite positive cm')
        canvas = None
    pages = deck.get('pages', [])
    if not isinstance(pages, list) or not pages:
        issue('schema', 'pages', 'Nonempty pages list required')
        pages = []
    page_ids = [p.get('id') for p in pages]
    if None in page_ids or len(set(page_ids)) != len(page_ids):
        issue('naming', 'pages', 'Page ids must be nonempty and unique')
    orders = [p.get('order') for p in pages]
    if any(not isinstance(v, int) or isinstance(v, bool) or v < 1 for v in orders) or len(set(orders)) != len(orders):
        issue('layout', 'pages.order', 'Positive unique integer order required')
    elif pages != sorted(pages, key=lambda p: p['order']):
        ordered_pages = sorted(pages, key=lambda p: p['order'])
        fix('layout', 'pages.order', [p['id'] for p in pages], [p['id'] for p in ordered_pages], 'Honor declared narrative order without deleting content', priority=1)
        pages[:] = ordered_pages
    if deck.get('deck_id', spec.get('deck_id')) != spec.get('deck_id'):
        issue('interface', 'deck_id', 'Spec and deck plan refer to different decks')
    evidence_registry = spec.get('evidence', deck.get('evidence'))
    if isinstance(evidence_registry, list):
        evidence_registry = {x['id']: x for x in evidence_registry if isinstance(x, dict) and 'id' in x}
    for page in pages:
        pid = page.get('id', '<missing>')
        for key in ['role', 'chapter', 'message', 'evidence_refs', 'layout_family', 'required_mechanisms', 'objects', 'relations', 'motion']:
            if key not in page:
                issue('schema', f'{pid}.{key}', 'Required page field missing')
        if not page.get('message'):
            issue('content', pid, 'Page message must remain complete; never manufacture or remove it')
        if evidence_registry is not None:
            for ref in page.get('evidence_refs', []):
                if ref not in evidence_registry:
                    issue('content', pid, 'Unknown evidence reference: ' + str(ref))
        variant = page.get('variant', 'default')
        variant_spec = {} if variant in {None, 'default'} else spec.get('variants', {}).get(variant)
        if variant_spec is None or (variant not in {None, 'default'} and pid not in variant_spec.get('pages', [])):
            issue('theme', f'{pid}.variant', 'Undeclared variant or page not explicitly allowed in variant.pages')
            variant_spec = {}
        effective = copy.deepcopy(tokens)
        for key, values in variant_spec.get('tokens', {}).items():
            if isinstance(values, dict):
                effective.setdefault(key, {}).update(values)
        page['style_context'] = {'variant': variant, 'background_role': variant_spec.get('background_role', 'paper'), 'text_role': variant_spec.get('text_role', 'ink'), 'tokens': effective}
        for field, category in [('master_id', 'master'), ('layout_id', 'layout')]:
            expected = variant_spec.get(field, spec.get(field))
            if expected is not None and page.get(field) != expected:
                fix(category, f'{pid}.{field}', page.get(field), expected, 'Use declared scope; explicit chapter variant remains permitted')
                page[field] = expected
        objects = page.get('objects', [])
        ids = [o.get('id') for o in objects]
        if any(not isinstance(x, str) or not x for x in ids) or len(set(ids)) != len(ids):
            issue('naming', pid, 'Duplicate/missing object ids cannot be safely renamed when references are ambiguous')
        morph_keys = []
        for obj in objects:
            oid = obj.get('id', '<missing>');loc = f'{pid}.{oid}'
            expected_name = '!!' + str(obj['morph_key']) if obj.get('morph_key') else f"{spec.get('deck_id', 'deck')}.{oid}"
            if obj.get('native_name') and obj['native_name'] != expected_name:
                fix('morph' if obj.get('morph_key') else 'naming', loc + '.native_name', obj['native_name'], expected_name, 'Native name and stable object identity must agree')
            obj['native_name'] = expected_name
            if obj.get('name') != expected_name:
                fix('naming', loc + '.name', obj.get('name'), expected_name, 'Stable native object identity is scoped by deck/page/object')
                obj['name'] = expected_name
            b = obj.get('bounds_cm')
            valid = isinstance(b, list) and len(b) == 4 and all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) for v in b)
            if not valid or b[2] <= 0 or b[3] <= 0 or b[0] < 0 or b[1] < 0 or (canvas and (b[0] + b[2] > canvas[0] + 1e-8 or b[1] + b[3] > canvas[1] + 1e-8)):
                issue('bounds', loc, 'Invalid/out-of-canvas bounds; reflow required, no clipping/deleting/shrinking text auto-fix')
            if obj.get('morph_key'):
                morph_keys.append(obj['morph_key'])
            if obj.get('font_role'):
                role = obj['font_role'];family = 'heading' if role == 'title' else 'body' if role in {'caption', 'footer'} else role
                fonts = effective.get('fonts', {}).get(family)
                size = effective.get('font_sizes_pt', {}).get('title' if role == 'heading' else role)
                if fonts is None or size is None:
                    issue('font', loc, 'Unknown font_role')
                else:
                    if obj.get('fonts') != fonts:
                        fix('font', loc + '.fonts', obj.get('fonts'), fonts, 'Apply explicit EA/Latin role tokens; preserve text')
                        obj['fonts'] = copy.deepcopy(fonts)
                    explicit = obj.get('font_size_pt', obj.get('size_pt', size))
                    valid_size = isinstance(explicit, (int, float)) and not isinstance(explicit, bool) and math.isfinite(explicit) and explicit > 0
                    if not valid_size:
                        issue('font', loc + '.font_size_pt', 'Explicit object size must be positive finite points')
                    else:
                        lower = spec.get('quality', {}).get('min_' + role + '_pt', spec.get('quality', {}).get('min_body_pt') if role in {'body', 'heading', 'title'} else None)
                        if isinstance(lower, (int, float)) and explicit < lower:
                            issue('readability', loc + '.font_size_pt', 'Object font size violates declared readability minimum; no auto-shrink or deletion')
                        if obj.get('font_size_pt') is not None and obj.get('size_pt') is not None and obj['size_pt'] != explicit:
                            fix('font', loc + '.size_pt', obj['size_pt'], explicit, 'Canonical font_size_pt wins over legacy size_pt; blocked readability remains blocked', priority=2)
                        obj['font_size_pt'] = explicit
                        obj['size_pt'] = explicit
                if not obj.get('color_role'):
                    color_role = spec.get('layout_rules', {}).get('footer_color_role', 'ink') if role == 'footer' else page['style_context']['text_role']
                    if color_role in effective.get('colors', {}):
                        obj['color_role'] = color_role
            if obj.get('color_role'):
                target = effective.get('colors', {}).get(obj['color_role'])
                if target is None:
                    issue('theme', loc, 'Unknown color_role')
                elif obj.get('color') != target:
                    fix('theme', loc + '.color', obj.get('color'), target, 'Use declared semantic color or allowed chapter variant')
                    obj['color'] = target
        if len(morph_keys) != len(set(morph_keys)):
            issue('morph', pid, 'Same MorphKey assigned to multiple objects on one slide')
        for rel in page.get('relations', []):
            if rel.get('type') == 'morph-pair':
                target = next((p for p in pages if p.get('id') == rel.get('target_page')), None)
                if target is None:
                    issue('morph', pid, 'Morph target page does not exist')
                else:
                    for key in rel.get('keys', []):
                        if sum(o.get('morph_key') == key for o in objects) != 1 or sum(o.get('morph_key') == key for o in target.get('objects', [])) != 1:
                            issue('morph', pid, 'Morph pair key must resolve exactly once on both pages: ' + key)
                        else:
                            a = next(o for o in objects if o.get('morph_key') == key)
                            b = next(o for o in target.get('objects', []) if o.get('morph_key') == key)
                            if a.get('kind') != b.get('kind'):
                                issue('morph', pid, 'Morph pair changes object kind; not covered by generic identity matching')
                continue
            endpoints = rel.get('objects', [rel.get('from'), rel.get('to')])
            if any(x not in ids for x in endpoints):
                issue('layout', pid, 'Relation contains unknown object endpoint')
                continue
            if rel.get('type') == 'left_of' and len(endpoints) == 2:
                a, b = [next(o for o in objects if o.get('id') == key).get('bounds_cm', []) for key in endpoints]
                if len(a) == len(b) == 4 and a[0] + a[2] + rel.get('gap_cm', 0) > b[0] + 1e-8:
                    issue('layout', pid, 'left_of/gap constraint violated; reflow without reducing content')
            elif rel.get('type') in {'equal-width', 'align-left', 'align-top'}:
                axis = {'equal-width': 2, 'align-left': 0, 'align-top': 1}[rel['type']]
                values = [next(o for o in objects if o.get('id') == key).get('bounds_cm', []) for key in endpoints]
                if values and all(len(b) == 4 for b in values) and max(b[axis] for b in values) - min(b[axis] for b in values) > rel.get('tolerance_cm', 0.01):
                    issue('layout', pid, rel['type'] + ' spatial constraint violated')
            elif rel.get('type') != 'left_of':
                issue('layout', pid, 'Unsupported spatial relation requires explicit adapter: ' + str(rel.get('type')))
        chosen = []
        seen_recipes = set()
        for requirement in page.get('required_mechanisms', []):
            candidates, error = select_recipe(requirement, recipes)
            if error:
                issue('interface', pid, error + ': ' + str(requirement))
            else:
                if candidates[0]['id'] not in seen_recipes:
                    chosen.append(candidates[0]);seen_recipes.add(candidates[0]['id'])
                    reuse = requirement.get('reuse_scope') if isinstance(requirement, dict) else None
                    excluded = requirement.get('excluded_capabilities', []) if isinstance(requirement, dict) else []
                    if reuse is not None and (not isinstance(reuse, str) or not reuse.strip()):
                        issue('interface', pid, 'reuse_scope must be nonempty explicit application text')
                    if not isinstance(excluded, list) or any(not isinstance(x, str) for x in excluded):
                        issue('interface', pid, 'excluded_capabilities must be a list of strings')
                    selected.append({'page': pid, 'requirement': requirement, **candidates[0], 'reuse_scope': reuse, 'excluded_capabilities': excluded, 'application_review': 'explicit-subset-requires-native-acceptance' if reuse else 'application-scope-unspecified', 'full_recipe_verification_inherited': False})
        motion = page.get('motion', {})
        duration = motion.get('duration_ms', 0)
        if not isinstance(duration, (float, int)) or isinstance(duration, bool) or not math.isfinite(duration) or duration < 0:
            issue('timeline', pid, 'Transition duration must be finite nonnegative ms')
        elif isinstance(motion.get('duration_seconds'), (int, float)) and abs(motion['duration_seconds'] - duration / 1000) > 1e-8:
            fix('timeline', pid + '.motion.duration_seconds', motion['duration_seconds'], duration / 1000, 'Canonical time is ms; COM seconds = ms/1000', priority=4)
            motion['duration_seconds'] = duration / 1000
        elif isinstance(duration, (int, float)):
            motion['duration_seconds'] = duration / 1000
        groups = motion.get('groups', [])
        if isinstance(spec.get('quality', {}).get('max_motion_groups_per_page'), int) and len(groups) > spec['quality']['max_motion_groups_per_page']:
            issue('timeline', pid, 'Motion density exceeds declared page budget; no silent effect deletion')
        writes = []
        for g in groups:
            start, dur = g.get('start_ms', 0), g.get('duration_ms')
            if not all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) and v >= 0 for v in [start, dur]):
                issue('timeline', pid, 'Group start/duration must be finite nonnegative ms')
                continue
            end = start + dur
            for field, target in [('start_seconds', start / 1000), ('duration_seconds', dur / 1000)]:
                if field in g and g[field] != target:
                    fix('timeline', f"{pid}.{g.get('id')}.{field}", g[field], target, 'Canonical time is milliseconds; convert to COM seconds exactly once', priority=4)
                g[field] = target
            for write in g.get('writes', []):
                if write.get('object') not in ids:
                    issue('timeline', pid, 'Writer targets unknown object')
                if not write.get('property'):
                    issue('timeline', pid, 'Writer property missing')
                item = {**write, 'group': g.get('id'), 'start_ms': start, 'end_ms': end, 'trigger': g.get('trigger', 'onClick'), 'sequence': g.get('sequence', 'main'), 'click_index': g.get('click_index')}
                for old in writes:
                    separate_click = old['sequence'] == item['sequence'] and old['trigger'] in {'click', 'onClick'} and item['trigger'] in {'click', 'onClick'} and old.get('click_index') is not None and item.get('click_index') is not None and old['click_index'] != item['click_index']
                    if separate_click:
                        continue
                    same_clock = old['sequence'] == item['sequence'] and old['trigger'] == item['trigger']
                    if old.get('object') == item.get('object') and old.get('property') == item.get('property') and (not same_clock or max(start, old['start_ms']) < min(end, old['end_ms']) or start == end == old['start_ms']):
                        issue('timeline', pid, 'Concurrent or incomparable-clock double writer: ' + str(write.get('property')))
                writes.append(item)
                if str(motion.get('transition', '')).lower() == 'morph' and g.get('trigger') in {'auto', 'withPrevious', 'afterPrevious'} and start < duration and write.get('property') in {'bounds_cm', 'position', 'rotation', 'scale'} and any(o.get('id') == write.get('object') and o.get('morph_key') for o in objects):
                    issue('morph', pid, 'Automatic local motion overlaps declared Morph transition property owner')
            timelines.append({'page': pid, 'group': g.get('id'), 'start_ms': start, 'duration_ms': dur, 'start_seconds': start / 1000, 'duration_seconds': dur / 1000, 'trigger': g.get('trigger', 'onClick'), 'sequence': g.get('sequence', 'main'), 'click_index': g.get('click_index'), 'writes': copy.deepcopy(g.get('writes', []))})
        page['_property_write_sets'] = writes
        build = pid + ':officecli-base';flush = pid + ':officecli-close'
        previous_tail = nodes[-1]['id'] if nodes else None
        if any(r['backend'] == 'external-render' for r in chosen):
            external = pid + ':external-assets'
            nodes.append({'id': external, 'backend': 'external-render', 'depends_on': [previous_tail] if previous_tail else [], 'action': 'Generate declared deterministic external assets before CLI embeds them; preserve editable source separately', 'executes': False})
            previous_tail = external
        static_writes = [{'page': pid, 'object': obj.get('id'), 'property': key, 'value': copy.deepcopy(obj[key])} for obj in objects for key in ['bounds_cm', 'text', 'fonts', 'font_size_pt', 'color', 'native_name', 'data_ref'] if key in obj]
        nodes.extend([{'id': build, 'backend': 'officecli-native', 'depends_on': [previous_tail] if previous_tail else [], 'writer': 'one-process-at-a-time', 'writes': static_writes, 'recipes': [r['id'] for r in chosen], 'executes': False}, {'id': flush, 'backend': 'officecli-native', 'depends_on': [build], 'action': 'CLI close/flush before any external reader', 'executes': False}])
        if any(r['backend'] in {'hybrid', 'powerpoint-com'} for r in chosen):
            nodes.append({'id': pid + ':com-polish', 'backend': 'powerpoint-com', 'depends_on': [flush], 'writer': 'exclusive shared PowerPoint slot; own presentation only', 'executes': False})
    close_ids = {n['id'] for n in nodes if n['id'].endswith(':officecli-close')}
    nodes.extend(copy.deepcopy(deck.get('operations', [])))
    try:
        order = topological(nodes)
    except (ValueError, TypeError) as error:
        issue('dag', 'operations', str(error));order = []
    for node in nodes:
        if node.get('backend') not in BACKENDS:
            issue('interface', node.get('id'), 'Unknown backend')
        if node.get('backend') == 'powerpoint-com' and not any(dep in close_ids for dep in node.get('depends_on', [])):
            issue('interface', node.get('id'), 'COM author requires explicit CLI close dependency')
    payload = {'version': 1, 'deck_id': spec.get('deck_id'), 'input_hashes': original, 'index': {'path': str(Path(index).resolve()), 'sha256': index_hash}, 'priority': PRIORITY, 'effective_spec': spec, 'effective_deck_plan': deck, 'selected_recipes': selected, 'operation_dag': {'nodes': nodes, 'topological_order': order}, 'timeline': timelines, 'single_writer': {'enforced_at_execution': False, 'required': True, 'rule': 'CLI resident must close before COM; exclusive COM/UI slot; never simultaneous writers on one PPTX'}, 'verification': {'status': 'unverified', 'base_recipe_status_does_not_imply_combination_verified': True}, 'status': 'blocked' if any(c['blocking'] for c in conflicts) else 'planned'}
    payload['plan_sha256'] = plan_digest(payload)
    return payload, {'version': 1, 'blocked': payload['status'] == 'blocked', 'conflicts': conflicts, 'automatic_repairs': fixes, 'defaults_applied': defaults, 'remaining': sum(c['blocking'] for c in conflicts)}


def bind_native_report(plan, report_path, deck_path):
    """Bind whole-deck trusted native evidence; never accept base recipe reports."""
    report = json.loads(Path(report_path).read_text(encoding='utf-8-sig'))
    verify_source_files(plan)
    required = {'scope': 'whole-deck', 'kind': 'native-powerpoint', 'passed': True, 'plan_sha256': plan['plan_sha256'], 'deck_sha256': sha(Path(deck_path).read_bytes()), 'index_sha256': plan['index']['sha256'], 'spec_sha256': plan['input_hashes']['spec'], 'deck_plan_sha256': plan['input_hashes']['deck_plan']}
    if plan['status'] != 'planned' or any(report.get(k) != v for k, v in required.items()):
        raise ValueError('Native report scope/hash/status mismatch')
    if any(not r.get('reuse_scope') for r in plan['selected_recipes']):
        raise ValueError('Explicit per-recipe reuse_scope required before combination verification')
    if sorted(report.get('pages', [])) != sorted(p['id'] for p in plan['effective_deck_plan']['pages']):
        raise ValueError('Native report must cover exact full page set')
    evidence = report.get('evidence', [])
    png_pages = {e.get('page_id') for e in evidence if e.get('kind') == 'png'}
    if png_pages != set(report['pages']):
        raise ValueError('Actual native PNG evidence required for every page_id')
    if not report.get('acceptance_id'):
        raise ValueError('Whole-deck acceptance run id required')
    dynamic = any(p.get('motion', {}).get('groups') or p.get('motion', {}).get('transition', 'none').lower() not in {'none', 'cut'} for p in plan['effective_deck_plan']['pages'])
    if dynamic and not any(e.get('kind') == 'native-recording' for e in evidence):
        raise ValueError('Dynamic combinations require actual native recording evidence')
    for item in evidence:
        file = Path(item['path'])
        if not file.is_absolute():
            raise ValueError('Evidence paths must be absolute')
        if sha(file.read_bytes()) != item.get('sha256'):
            raise ValueError('Native evidence file hash mismatch')
    plan['verification'] = {'status': 'verified', 'native_report_path': str(Path(report_path).resolve()), 'native_report_sha256': sha(Path(report_path).read_bytes()), 'deck_sha256': required['deck_sha256'], 'trust_boundary': 'Caller supplies genuine native playback evidence; hash checks cannot prove recording provenance.'}
    plan['status'] = 'verified'


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--spec', required=True, type=Path)
    parser.add_argument('--deck-plan', required=True, type=Path)
    parser.add_argument('--index', required=True, type=Path)
    parser.add_argument('--out', required=True, type=Path)
    parser.add_argument('--native-report', type=Path)
    parser.add_argument('--deck', type=Path)
    args = parser.parse_args()
    if args.out.exists():
        parser.error('Output directory exists; refusing overwrite')
    if bool(args.native_report) != bool(args.deck):
        parser.error('--native-report and --deck are required together')
    spec = json.loads(args.spec.read_text(encoding='utf-8-sig'));deck = json.loads(args.deck_plan.read_text(encoding='utf-8-sig'))
    plan, conflicts = compile_plan(spec, deck, args.index)
    plan['source_files'] = source_files(args.spec, args.deck_plan, spec, deck)
    plan['plan_sha256'] = plan_digest(plan)
    if args.native_report:
        bind_native_report(plan, args.native_report, args.deck)
    args.out.mkdir(parents=True)
    for name, data in [('composition-plan.json', plan), ('conflicts.json', conflicts)]:
        (args.out / name).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'status': plan['status'], 'blocking_conflicts': conflicts['remaining'], 'plan_sha256': plan['plan_sha256']}))
    return 2 if conflicts['blocked'] else 0


if __name__ == '__main__':
    raise SystemExit(main())
