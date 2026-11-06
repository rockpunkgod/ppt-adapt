#!/usr/bin/env python3
"""Validate director v1 plans; no Office, renderer or third-party dependencies."""
import argparse
import json
import math
import re
from pathlib import Path
from urllib.parse import urlparse


def validate(spec, plan, spec_path, plan_path, page=None):
    errors, warnings = [], []

    def fail(where, message):
        errors.append(f'{where}: {message}')

    def obj(value, where):
        if not isinstance(value, dict):
            fail(where, 'must be an object')
            return {}
        return value

    def arr(value, where):
        if not isinstance(value, list):
            fail(where, 'must be an array')
            return []
        return value

    def string(value, where):
        if not isinstance(value, str) or not value.strip():
            fail(where, 'must be a nonempty string')
            return ''
        return value

    def number(value, where, minimum=0, positive=False, integer=False):
        if (isinstance(value, bool) or not isinstance(value, (int, float))
                or not math.isfinite(value) or value < minimum
                or (positive and value <= 0) or (integer and not isinstance(value, int))):
            fail(where, 'must be a finite ' + ('integer ' if integer else 'number ')
                 + ('> 0' if positive else f'>= {minimum}'))
            return None
        return value

    def strings(value, where):
        return [string(v, f'{where}[{i}]') for i, v in enumerate(arr(value, where))]

    def unique(values, where):
        if len(values) != len(set(values)):
            fail(where, 'duplicate identifiers')

    spec = obj(spec, 'spec')
    plan = obj(plan, 'plan')
    for label, data in [('spec', spec), ('plan', plan)]:
        if type(data.get('version')) is not int or data.get('version') != 1:
            fail(label + '.version', 'must be integer 1')
        string(data.get('deck_id'), label + '.deck_id')
    if spec.get('deck_id') != plan.get('deck_id'):
        fail('deck_id', 'spec and plan must match')
    declared = string(plan.get('design_spec'), 'plan.design_spec')
    if declared and (plan_path.parent / declared).resolve() != spec_path.resolve():
        fail('plan.design_spec', 'does not resolve to supplied spec')
    for key in ('audience', 'purpose', 'decision'):
        string(spec.get(key), 'spec.' + key)
    canvas = obj(spec.get('canvas_cm'), 'spec.canvas_cm')
    width = number(canvas.get('width'), 'canvas.width', positive=True)
    height = number(canvas.get('height'), 'canvas.height', positive=True)
    tokens = obj(spec.get('tokens'), 'spec.tokens')
    fonts = obj(tokens.get('fonts'), 'tokens.fonts')
    for role in ('heading', 'body'):
        family = obj(fonts.get(role), 'fonts.' + role)
        for script in ('ea', 'latin'):
            string(family.get(script), f'fonts.{role}.{script}')
    sizes = obj(tokens.get('font_sizes_pt'), 'font_sizes_pt')
    for role in ('title', 'body', 'caption', 'footer'):
        number(sizes.get(role), 'font_sizes_pt.' + role, positive=True)
    colors = obj(tokens.get('colors'), 'colors')
    for role in ('paper', 'ink', 'accent1', 'accent2', 'muted', 'night'):
        if role not in colors:
            fail('colors.' + role, 'required role missing')
    for role, color in colors.items():
        if not isinstance(color, str) or not re.fullmatch(r'#?[0-9a-fA-F]{6}', color):
            fail('colors.' + role, 'must be six hex digits, optionally prefixed #')
    spacing = obj(tokens.get('spacing_cm'), 'spacing_cm')
    margin = number(spacing.get('margin'), 'spacing_cm.margin')
    number(spacing.get('gap'), 'spacing_cm.gap')
    if margin is not None and width and height and 2 * margin >= min(width, height):
        fail('spacing_cm.margin', 'leaves no usable canvas')
    quality = obj(spec.get('quality'), 'quality')
    min_body = number(quality.get('min_body_pt'), 'quality.min_body_pt', positive=True)
    max_groups = number(quality.get('max_motion_groups_per_page'),
                        'quality.max_motion_groups_per_page', integer=True)
    if (min_body is not None and isinstance(sizes.get('body'), (int, float))
            and sizes['body'] < min_body):
        fail('font_sizes_pt.body', 'below quality.min_body_pt')

    narrative = obj(plan.get('narrative'), 'plan.narrative')
    for key in ('opening', 'throughline', 'closing_ask'):
        string(narrative.get(key), 'narrative.' + key)
    evidence = arr(plan.get('evidence'), 'plan.evidence')
    evidence_ids, evidence_json = [], {}
    for i, raw in enumerate(evidence):
        item = obj(raw, f'evidence[{i}]')
        evidence_ids.append(string(item.get('id'), f'evidence[{i}].id'))
        string(item.get('kind'), f'evidence[{i}].kind')
        path = string(item.get('path'), f'evidence[{i}].path')
        if path:
            if path.startswith(('http://', 'https://')):
                parsed = urlparse(path)
                if not parsed.netloc or parsed.username or parsed.password:
                    fail(f'evidence[{i}].path', 'invalid URL or embedded credentials')
                else:
                    warnings.append(f'evidence[{i}]: URL existence/content not checked')
            else:
                local_path = plan_path.parent / path
                if not local_path.exists():
                    fail(f'evidence[{i}].path', 'local evidence does not exist')
                elif local_path.is_file() and local_path.suffix.lower() == '.json':
                    try:
                        evidence_json[evidence_ids[-1]] = json.loads(local_path.read_text(encoding='utf-8-sig'))
                    except (OSError, ValueError, TypeError) as exc:
                        fail(f'evidence[{i}].path', 'invalid JSON evidence: ' + str(exc))
    unique(evidence_ids, 'evidence')
    evidence_set = set(evidence_ids)

    def data_resolves(ref, refs):
        if ref in evidence_set:
            return ref in refs
        # Field paths are resolved against actual page-linked local JSON only.
        # Optional evidence-id#path removes ambiguity across multiple sources.
        owner, sep, selector = ref.partition('#')
        selector = selector if sep else ref
        if not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_-]*(?:\[\d+\]|\.[A-Za-z_][A-Za-z0-9_-]*)*', selector):
            return False
        parts = re.findall(r'[A-Za-z_][A-Za-z0-9_-]*|\d+', selector)
        matches = 0
        for eid in refs:
            if eid not in evidence_json or (sep and owner != eid):
                continue
            value = evidence_json[eid]
            try:
                for part in parts:
                    value = value[int(part)] if isinstance(value, list) and part.isdigit() else value[part]
                matches += 1
            except (KeyError, IndexError, TypeError, ValueError):
                continue
        return matches == 1
    pages = arr(plan.get('pages'), 'plan.pages')
    if not pages:
        fail('plan.pages', 'at least one page required')
    normalized = [obj(p, f'pages[{i}]') for i, p in enumerate(pages)]
    page_ids = [string(p.get('id'), f'pages[{i}].id') for i, p in enumerate(normalized)]
    unique(page_ids, 'pages')
    if page is not None and page not in page_ids:
        fail('--page', 'unknown page id')
    orders = []
    for i, p in enumerate(normalized):
        orders.append(number(p.get('order'), f'pages[{i}].order', minimum=1, integer=True))
    if all(o is not None for o in orders) and sorted(orders) != list(range(1, len(pages) + 1)):
        fail('pages.order', 'must be unique continuous 1..N')

    variants = obj(spec.get('variants'), 'spec.variants')
    if 'default' in variants:
        fail('variants.default', 'default is implicit, cannot redefine')
    variant_members = {}
    for name, raw in variants.items():
        v = obj(raw, 'variants.' + name)
        members = strings(v.get('pages'), f'variants.{name}.pages')
        unique(members, f'variants.{name}.pages')
        if not members:
            fail('variants.' + name, 'empty page scope')
        for pid in members:
            if pid not in page_ids:
                fail('variants.' + name, 'unknown page ' + pid)
            if pid in variant_members:
                fail('variants.' + name, 'page in multiple variants: ' + pid)
            variant_members[pid] = name
        string(v.get('reason'), f'variants.{name}.reason')
        bg = string(v.get('background_role'), f'variants.{name}.background_role')
        if bg and bg not in colors:
            fail('variants.' + name, 'unknown background color role')
        text_role = string(v.get('text_role'), f'variants.{name}.text_role')
        if text_role and text_role not in colors:
            warnings.append(f'variants.{name}: semantic text_role requires executor resolution: {text_role}')

    checked = []
    for i, p in enumerate(normalized):
        pid = page_ids[i]
        label = 'pages.' + (pid or str(i))
        for key in ('role', 'chapter', 'message', 'layout_family'):
            string(p.get(key), label + '.' + key)
        if 'page_job' in p:
            string(p['page_job'], label + '.page_job')
        refs = strings(p.get('evidence_refs'), label + '.evidence_refs')
        if not refs:
            fail(label + '.evidence_refs', 'at least one evidence reference required')
        for ref in refs:
            if ref not in evidence_set:
                fail(label + '.evidence_refs', 'unknown evidence ' + ref)
        for j, requirement in enumerate(arr(p.get('required_mechanisms'), label + '.required_mechanisms')):
            loc = f'{label}.required_mechanisms[{j}]'
            if isinstance(requirement, str):
                string(requirement, loc)
            elif isinstance(requirement, dict):
                string(requirement.get('query'), loc + '.query')
                if 'recipe' in requirement and not isinstance(requirement['recipe'], str):
                    fail(loc + '.recipe', 'must be a string')
                if 'reuse_scope' in requirement:
                    string(requirement['reuse_scope'], loc + '.reuse_scope')
                if 'excluded_capabilities' in requirement:
                    strings(requirement['excluded_capabilities'], loc + '.excluded_capabilities')
            else:
                fail(loc, 'must be a nonempty string or mechanism object')
        variant = p.get('variant', 'default')
        if not isinstance(variant, str) or variant != variant_members.get(pid, 'default'):
            fail(label + '.variant', 'must exactly match declared spec page scope')
        if page is not None and pid != page:
            continue
        checked.append(pid)
        objects = [obj(v, label + '.objects') for v in arr(p.get('objects'), label + '.objects')]
        ids, morphs = [], []
        for j, o in enumerate(objects):
            loc = f'{label}.objects[{j}]'
            ids.append(string(o.get('id'), loc + '.id'))
            kind = string(o.get('kind'), loc + '.kind')
            bounds = arr(o.get('bounds_cm'), loc + '.bounds_cm')
            if len(bounds) != 4:
                fail(loc + '.bounds_cm', 'requires [x,y,w,h]')
            else:
                b = [number(v, loc + '.bounds_cm', positive=k > 1) for k, v in enumerate(bounds)]
                if all(v is not None for v in b) and width and height:
                    x, y, w, h = b
                    if x + w > width + 1e-6 or y + h > height + 1e-6:
                        fail(loc + '.bounds_cm', 'outside canvas')
            if kind in ('text', 'title', 'heading', 'body', 'caption', 'footer') or 'text' in o:
                string(o.get('text'), loc + '.text')
                role = o.get('font_role')
                if role not in ('heading', 'title', 'body', 'caption', 'footer'):
                    fail(loc + '.font_role', 'text needs a known font role')
                else:
                    size = number(o.get('font_size_pt', sizes.get('title' if role == 'heading' else role)),
                                  loc + '.font_size_pt', positive=True)
                    if role == 'body' and size is not None and min_body is not None and size < min_body:
                        fail(loc + '.font_size_pt', 'body text below minimum')
            elif 'font_size_pt' in o:
                number(o['font_size_pt'], loc + '.font_size_pt', positive=True)
            if kind in ('chart', 'table') and 'data_ref' not in o:
                fail(loc + '.data_ref', 'chart/table needs evidence-backed data')
            if 'data_ref' in o:
                ref = string(o['data_ref'], loc + '.data_ref')
                if not data_resolves(ref, refs):
                    fail(loc + '.data_ref', 'must resolve uniquely in evidence referenced by this page')
            if 'morph_key' in o:
                morphs.append(string(o['morph_key'], loc + '.morph_key'))
        unique(ids, label + '.objects')
        unique(morphs, label + '.morph_key')
        for raw in arr(p.get('relations'), label + '.relations'):
            relation = obj(raw, label + '.relations')
            if 'from' in relation or 'to' in relation:
                for key in ('from', 'to'):
                    ref = string(relation.get(key), label + '.relations.' + key)
                    if ref not in ids:
                        fail(label + '.relations.' + key, 'unknown object')
            elif 'objects' in relation:
                members = strings(relation['objects'], label + '.relations.objects')
                if len(members) < 2:
                    fail(label + '.relations.objects', 'relation requires at least two objects')
                for ref in members:
                    if ref not in ids:
                        fail(label + '.relations.objects', 'unknown object ' + ref)
                string(relation.get('type'), label + '.relations.type')
            elif relation.get('type') == 'morph-pair':
                target = string(relation.get('target_page'), label + '.relations.target_page')
                target_page = next((v for v in normalized if v.get('id') == target), None)
                if target_page is None or target == pid:
                    fail(label + '.relations.target_page', 'must reference another existing page')
                keys = strings(relation.get('keys'), label + '.relations.keys')
                if not keys:
                    fail(label + '.relations.keys', 'morph-pair requires identity keys')
                target_objects = target_page.get('objects', []) if target_page else []
                target_keys = [v.get('morph_key') for v in target_objects if isinstance(v, dict)] if isinstance(target_objects, list) else []
                for key in keys:
                    if key not in morphs or target_keys.count(key) != 1:
                        fail(label + '.relations.keys', 'key must occur once on both pages: ' + key)
            else:
                fail(label + '.relations', 'requires from/to, type+objects, or morph-pair')
        motion = obj(p.get('motion'), label + '.motion')
        transition = motion.get('transition')
        if transition not in ('none', 'morph'):
            fail(label + '.motion.transition', 'must be none or morph')
        duration = number(motion.get('duration_ms'), label + '.motion.duration_ms')
        if transition == 'none' and duration not in (None, 0):
            fail(label + '.motion.duration_ms', 'none transition requires 0')
        groups = [obj(g, label + '.motion.groups') for g in arr(motion.get('groups'), label + '.motion.groups')]
        if max_groups is not None and len(groups) > max_groups:
            fail(label + '.motion.groups', 'exceeds maximum motion groups')
        gids = [string(g.get('id'), label + '.motion.groups.id') for g in groups]
        unique(gids, label + '.motion.groups')
        graph = {}
        for gid, g in zip(gids, groups):
            for field in ('duration_ms', 'delay_ms'):
                if field in g:
                    number(g[field], label + '.motion.groups.' + field)
            for target in strings(g.get('targets', []), label + '.motion.groups.targets'):
                if target not in ids:
                    fail(label + '.motion.groups.targets', 'unknown object ' + target)
            for target in strings(g.get('object_ids', []), label + '.motion.groups.object_ids'):
                if target not in ids:
                    fail(label + '.motion.groups.object_ids', 'unknown object ' + target)
            dependencies = strings(g.get('depends_on', []), label + '.motion.groups.depends_on')
            graph[gid] = dependencies
            for dependency in dependencies:
                if dependency not in gids:
                    fail(label + '.motion.groups.depends_on', 'unknown group ' + dependency)
        # Iterative topological removal avoids recursion limits on malformed large input.
        pending = {k: set(v) & set(gids) for k, v in graph.items()}
        while pending:
            ready = {k for k, v in pending.items() if not v}
            if not ready:
                fail(label + '.motion.groups', 'dependency cycle')
                break
            pending = {k: v - ready for k, v in pending.items() if k not in ready}

    if len(normalized) >= 4:
        if len({str(p.get('layout_family')) for p in normalized}) == 1 and not any(p.get('rhythm_reason') for p in normalized):
            warnings.append('pages: one layout family throughout; explain rhythm_reason or reassess hierarchy')
        ordered = sorted(normalized, key=lambda p: p.get('order') if type(p.get('order')) is int else 0)
        for start in range(len(ordered) - 3):
            run = ordered[start:start + 4]
            if len({str(p.get('role')) for p in run}) == 1 and not any(p.get('rhythm_reason') for p in run):
                warnings.append('pages: four consecutive same-role pages; review rhythm or record rhythm_reason')
                break
    return {'valid': not errors, 'errors': errors, 'warnings': warnings, 'checked_pages': checked,
            'scope': 'offline-plan-contract-only', 'version': 1}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--spec', required=True, type=Path)
    parser.add_argument('--plan', required=True, type=Path)
    parser.add_argument('--page')
    parser.add_argument('--report', type=Path)
    args = parser.parse_args()
    try:
        spec = json.loads(args.spec.read_text(encoding='utf-8-sig'))
        plan = json.loads(args.plan.read_text(encoding='utf-8-sig'))
        result = validate(spec, plan, args.spec, args.plan, args.page)
    except (OSError, ValueError, TypeError, OverflowError) as exc:
        result = {'valid': False, 'errors': [f'input: {exc}'], 'warnings': [], 'checked_pages': []}
    if args.report:
        try:
            with args.report.open('x', encoding='utf-8') as handle:
                json.dump(result, handle, ensure_ascii=False, indent=2)
        except OSError as exc:
            result['valid'] = False
            result['errors'].append(f'report: {exc}')
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result['valid'] else 1


if __name__ == '__main__':
    raise SystemExit(main())
