"""Contract/conflict tests; fake acceptance fixtures test binding, not native rendering."""
import copy
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
spec_module = importlib.util.spec_from_file_location('planner', ROOT / 'scripts/plan_composition.py')
m = importlib.util.module_from_spec(spec_module)
spec_module.loader.exec_module(m)
INDEX = ROOT.parent / 'officecli-ppt-design/references/recipe-index.md'


def fixture():
    spec = {'version': 1, 'deck_id': 'sample', 'canvas_cm': {'width': 33.87, 'height': 19.05},
            'master_id': 'shared', 'layout_id': 'evidence',
            'quality': {'min_body_pt': 18},
            'tokens': {'fonts': {'heading': {'ea': 'Microsoft YaHei', 'latin': 'Arial'}, 'body': {'ea': 'Microsoft YaHei', 'latin': 'Arial'}},
                       'font_sizes_pt': {'title': 34, 'body': 20, 'caption': 14, 'footer': 11},
                       'colors': {'ink': '17324E'}, 'spacing_cm': {'margin': 1.5}}}
    page = {'id': 's01', 'order': 1, 'role': 'evidence', 'chapter': 'proof', 'message': 'A complete claim supported by evidence',
            'evidence_refs': [], 'layout_family': 'chart', 'required_mechanisms': [{'query': 'theme_semantics_*', 'reuse_scope': 'Native EA/Latin theme fonts and semantic colors', 'excluded_capabilities': ['themeOverride']}], 'variant': 'default',
            'objects': [{'id': 'title', 'kind': 'text', 'text': 'Preserve every original word', 'font_role': 'heading', 'color_role': 'ink', 'bounds_cm': [1.5, 1.2, 30, 2]},
                        {'id': 'chart', 'kind': 'chart', 'data_ref': 'data', 'bounds_cm': [1.5, 5, 20, 10]}],
            'relations': [], 'motion': {'transition': 'none', 'duration_ms': 0, 'groups': []}}
    return spec, {'version': 1, 'deck_id': 'sample', 'pages': [page]}


class PlannerTests(unittest.TestCase):
    def run_plan(self, mutate=None):
        spec, deck = fixture()
        if mutate:
            mutate(spec, deck)
        return m.compile_plan(spec, deck, INDEX)

    def test_actual_index_selection_and_sha(self):
        plan, _ = self.run_plan()
        self.assertEqual(plan['status'], 'planned')
        self.assertEqual(plan['verification']['status'], 'unverified')
        self.assertEqual(plan['index']['sha256'], m.sha(INDEX.read_bytes()))
        self.assertEqual(plan['selected_recipes'][0]['backend'], 'officecli-native')
        self.assertIn('theme_semantics_*', plan['selected_recipes'][0]['entrypoints'])
        self.assertEqual(plan['plan_sha256'], m.plan_digest(plan))

    def test_eight_categories_and_real_repairs(self):
        def inject(spec, deck):
            p = deck['pages'][0];o = p['objects'][0]
            o.update(fonts={'ea': 'Wrong', 'latin': 'Wrong'}, color='FF0000', name='unstable')
            p.update(master_id='wrong-master', layout_id='wrong-layout')
            p['objects'][0]['morph_key'] = 'duplicate';p['objects'][1]['morph_key'] = 'duplicate'
            p['motion'].update(duration_ms=700, duration_seconds=700)
            p['required_mechanisms'].append('nonexistent-quantum-renderer')
        plan, report = self.run_plan(inject)
        categories = {c['category'] for c in report['conflicts']}
        self.assertTrue({'font', 'theme', 'master', 'layout', 'naming', 'morph', 'timeline', 'interface'} <= categories)
        p = plan['effective_deck_plan']['pages'][0];o = p['objects'][0]
        self.assertEqual(o['fonts']['ea'], 'Microsoft YaHei')
        self.assertEqual(o['color'], '17324E')
        self.assertEqual(o['native_name'], '!!duplicate')
        self.assertEqual(o['text'], 'Preserve every original word')
        self.assertEqual(p['master_id'], 'shared');self.assertEqual(p['layout_id'], 'evidence')
        self.assertEqual(p['motion']['duration_seconds'], 0.7)
        self.assertEqual(plan['status'], 'blocked')
        self.assertTrue(all({'before', 'after', 'reason', 'priority'} <= set(r) for r in report['automatic_repairs']))

    def test_no_font_shrink(self):
        plan, report = self.run_plan(lambda s, d: d['pages'][0]['objects'][0].update(size_pt=12))
        self.assertEqual(plan['status'], 'blocked')
        self.assertEqual(plan['effective_deck_plan']['pages'][0]['objects'][0]['size_pt'], 12)
        self.assertTrue(any(c['category'] == 'readability' and c['blocking'] for c in report['conflicts']))

    def test_canonical_font_size_pt_cannot_hide_behind_legacy_alias(self):
        plan, report = self.run_plan(lambda s, d: d['pages'][0]['objects'][0].update(font_size_pt=10, size_pt=34))
        self.assertEqual(plan['status'], 'blocked')
        obj = plan['effective_deck_plan']['pages'][0]['objects'][0]
        self.assertEqual(obj['font_size_pt'], 10)
        self.assertEqual(obj['size_pt'], 10)
        self.assertEqual(obj['native_name'], 'sample.title')
        self.assertTrue(any(c['category'] == 'readability' and c['blocking'] for c in report['conflicts']))

    def test_reuse_scope_explicit_subset_preserves_published_scope(self):
        plan, _ = self.run_plan()
        item = plan['selected_recipes'][0]
        self.assertIn('Native EA/Latin', item['reuse_scope'])
        self.assertEqual(item['excluded_capabilities'], ['themeOverride'])
        self.assertTrue(item['scope'])
        self.assertFalse(item['full_recipe_verification_inherited'])
        legacy, _ = self.run_plan(lambda s, d: d['pages'][0].update(required_mechanisms=['theme_semantics_*']))
        self.assertEqual(legacy['status'], 'planned')
        self.assertEqual(legacy['selected_recipes'][0]['application_review'], 'application-scope-unspecified')
        self.assertEqual(legacy['verification']['status'], 'unverified')

    def test_bad_bounds_nan_and_outside(self):
        for bounds in [[0, 0, -1, 2], [0, 0, float('nan'), 2], [33, 0, 2, 1], [0, 0, 1]]:
            with self.subTest(bounds=bounds):
                _, report = self.run_plan(lambda s, d: d['pages'][0]['objects'][0].update(bounds_cm=bounds))
                self.assertTrue(any(c['category'] == 'bounds' and c['blocking'] for c in report['conflicts']))

    def test_dag_cycle_and_unknown_dependency(self):
        for operations in [[{'id': 'a', 'backend': 'officecli-native', 'depends_on': ['b']}, {'id': 'b', 'backend': 'officecli-native', 'depends_on': ['a']}], [{'id': 'a', 'backend': 'officecli-native', 'depends_on': ['missing']}]]:
            _, report = self.run_plan(lambda s, d: d.update(operations=operations))
            self.assertTrue(any(c['category'] == 'dag' for c in report['conflicts']))

    def test_duplicate_ids_and_morph_keys(self):
        def inject(s, d):
            objects = d['pages'][0]['objects'];objects[1]['id'] = objects[0]['id']
            for o in objects:
                o['morph_key'] = 'same'
        _, report = self.run_plan(inject)
        self.assertTrue({'naming', 'morph'} <= {c['category'] for c in report['conflicts'] if c['blocking']})

    def test_morph_pair_identity_and_layout_constraints(self):
        def inject(s, d):
            p = d['pages'][0];p['objects'][0]['morph_key'] = 'focus';p['relations'] = [{'type': 'morph-pair', 'target_page': 's02', 'keys': ['focus']}]
            q = copy.deepcopy(p);q.update(id='s02', order=2, relations=[]);d['pages'].append(q)
        plan, report = self.run_plan(inject)
        self.assertFalse(report['blocked'])
        self.assertEqual([p['objects'][0]['native_name'] for p in plan['effective_deck_plan']['pages']], ['!!focus', '!!focus'])
        _, report = self.run_plan(lambda s, d: d['pages'][0].update(relations=[{'type': 'equal-width', 'objects': ['title', 'chart']}]))
        self.assertTrue(any(c['category'] == 'layout' and c['blocking'] for c in report['conflicts']))

    def test_double_writer_and_distinct_click_clocks(self):
        def inject(s, d):
            d['pages'][0]['motion']['groups'] = [
                {'id': 'one', 'start_ms': 0, 'duration_ms': 500, 'trigger': 'click', 'click_index': 1, 'writes': [{'object': 'chart', 'property': 'rotation', 'value': 90}]},
                {'id': 'two', 'start_ms': 250, 'duration_ms': 500, 'trigger': 'click', 'click_index': 1, 'writes': [{'object': 'chart', 'property': 'rotation', 'value': 180}]}]
        _, report = self.run_plan(inject)
        self.assertTrue(any('double writer' in c['reason'] for c in report['conflicts']))
        spec, deck = fixture();inject(spec, deck);deck['pages'][0]['motion']['groups'][1]['click_index'] = 2
        plan, report = m.compile_plan(spec, deck, INDEX)
        self.assertFalse(report['blocked'])
        self.assertEqual(plan['timeline'][0]['duration_seconds'], 0.5)

    def test_undeclared_variant_and_interface(self):
        _, report = self.run_plan(lambda s, d: d['pages'][0].update(variant='night'))
        self.assertTrue(any(c['category'] == 'theme' and c['blocking'] for c in report['conflicts']))
        _, report = self.run_plan(lambda s, d: d.update(operations=[{'id': 'unsafe-com', 'backend': 'powerpoint-com', 'depends_on': []}]))
        self.assertTrue(any(c['category'] == 'interface' and c['blocking'] for c in report['conflicts']))
        _, report = self.run_plan(lambda s, d: d.update(operations=[{'id': 'forged:com-polish', 'backend': 'powerpoint-com', 'depends_on': []}]))
        self.assertTrue(any(c['category'] == 'interface' and c['blocking'] for c in report['conflicts']))

    def test_existing_ambiguous_recipe_blocks(self):
        plan, report = self.run_plan(lambda s, d: d['pages'][0].update(required_mechanisms=['build_components.py']))
        self.assertEqual(plan['status'], 'blocked')
        self.assertTrue(any('Ambiguous' in c['reason'] for c in report['conflicts']))

    def test_single_writer_default_dag(self):
        def inject(s, d):
            p = copy.deepcopy(d['pages'][0]);p.update(id='s02', order=2);d['pages'].append(p)
        plan, _ = self.run_plan(inject)
        nodes = plan['operation_dag']['nodes']
        self.assertEqual(nodes[2]['depends_on'], [nodes[1]['id']])

    def test_native_binding_rejects_base_partial_stale_and_tampered(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory);deck = root / 'fixture.pptx';image = root / 'fixture.png';report = root / 'acceptance.json'
            deck.write_bytes(b'unit fixture - not a real presentation');image.write_bytes(b'unit fixture - not native pixels')
            plan, _ = self.run_plan()
            valid = {'scope': 'whole-deck', 'kind': 'native-powerpoint', 'passed': True, 'acceptance_id': 'unit-binding-test-only', 'plan_sha256': plan['plan_sha256'], 'deck_sha256': m.sha(deck.read_bytes()), 'index_sha256': plan['index']['sha256'], 'spec_sha256': plan['input_hashes']['spec'], 'deck_plan_sha256': plan['input_hashes']['deck_plan'], 'pages': ['s01'], 'evidence': [{'kind': 'png', 'page_id': 's01', 'path': str(image.resolve()), 'sha256': m.sha(image.read_bytes())}]}
            for change in [{'scope': 'base-recipe'}, {'plan_sha256': 'stale'}, {'pages': []}, {'evidence': []}]:
                report.write_text(json.dumps({**valid, **change}), encoding='utf-8')
                with self.assertRaises(ValueError):
                    m.bind_native_report(copy.deepcopy(plan), report, deck)
            report.write_text(json.dumps(valid), encoding='utf-8');m.bind_native_report(plan, report, deck)
            self.assertEqual(plan['status'], 'verified');self.assertEqual(plan['plan_sha256'], m.plan_digest(plan))
            unspecified, _ = self.run_plan(lambda s, d: d['pages'][0].update(required_mechanisms=['theme_semantics_*']))
            report.write_text(json.dumps({**valid, 'plan_sha256': unspecified['plan_sha256'], 'deck_plan_sha256': unspecified['input_hashes']['deck_plan']}), encoding='utf-8')
            with self.assertRaisesRegex(ValueError, 'reuse_scope'):
                m.bind_native_report(unspecified, report, deck)
            report.write_text(json.dumps(valid), encoding='utf-8')
            deck.write_bytes(b'changed')
            with self.assertRaises(ValueError):
                m.bind_native_report(self.run_plan()[0], report, deck)

    def test_cli_rejects_existing_out(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory);spec, deck = fixture()
            for name, value in [('spec.json', spec), ('deck.json', deck)]:
                (root / name).write_text(json.dumps(value), encoding='utf-8')
            cmd = [sys.executable, str(ROOT / 'scripts/plan_composition.py'), '--spec', str(root / 'spec.json'), '--deck-plan', str(root / 'deck.json'), '--index', str(INDEX), '--out', str(root / 'out')]
            result = subprocess.run(cmd, capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            before = (root / 'out/composition-plan.json').read_bytes()
            result = subprocess.run(cmd, capture_output=True, text=True)
            self.assertNotEqual(result.returncode, 0)
            self.assertEqual(before, (root / 'out/composition-plan.json').read_bytes())

    def test_cli_binds_local_evidence_and_detects_data_change(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory);spec, deck = fixture()
            deck['evidence'] = [{'id': 'dataset', 'path': 'data.json'}]
            for name, value in [('spec.json', spec), ('deck.json', deck), ('data.json', {'value': 20})]:
                (root / name).write_text(json.dumps(value), encoding='utf-8')
            plan, _ = m.compile_plan(spec, deck, INDEX)
            plan['source_files'] = m.source_files(root / 'spec.json', root / 'deck.json', spec, deck)
            plan['plan_sha256'] = m.plan_digest(plan)
            self.assertEqual(plan['source_files']['evidence'][0]['path'], str((root / 'data.json').resolve()))
            self.assertTrue(m.verify_source_files(plan))
            old_hash = plan['plan_sha256']
            (root / 'data.json').write_text('{"value":80}', encoding='utf-8')
            with self.assertRaisesRegex(ValueError, 'changed since planning'):
                m.verify_source_files(plan)
            plan['source_files'] = m.source_files(root / 'spec.json', root / 'deck.json', spec, deck)
            self.assertNotEqual(m.plan_digest(plan), old_hash)


if __name__ == '__main__':
    unittest.main()
