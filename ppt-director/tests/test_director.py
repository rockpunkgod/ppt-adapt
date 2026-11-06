import copy
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = SKILL_ROOT / 'scripts/validate_deck_plan.py'
sys.dont_write_bytecode = True
module_spec = importlib.util.spec_from_file_location('director_validator', SCRIPT)
validator = importlib.util.module_from_spec(module_spec)
module_spec.loader.exec_module(validator)


class DirectorTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        (self.base / 'source.txt').write_text('User-provided original demonstration facts.', encoding='utf-8')
        self.spec = {'version': 1, 'deck_id': 'demo', 'canvas_cm': {'width': 33.87, 'height': 19.05},
                     'audience': 'product leaders', 'purpose': 'compare stages', 'decision': 'approve pilot',
                     'tokens': {'fonts': {r: {'ea': 'Microsoft YaHei', 'latin': 'Arial'} for r in ('heading', 'body')},
                                'font_sizes_pt': {'title': 34, 'body': 20, 'caption': 14, 'footer': 11},
                                'colors': {r: '#112233' for r in ('paper', 'ink', 'accent1', 'accent2', 'muted', 'night')},
                                'spacing_cm': {'margin': 1.5, 'gap': 0.5}},
                     'variants': {}, 'quality': {'min_body_pt': 18, 'max_motion_groups_per_page': 4}}
        page = {'id': 's01', 'order': 1, 'role': 'decision', 'chapter': 'pilot', 'message': 'Run a bounded pilot',
                'evidence_refs': ['e1'], 'layout_family': 'comparison', 'required_mechanisms': [],
                'objects': [{'id': 'body', 'kind': 'text', 'font_role': 'body', 'text': 'Pilot scope', 'bounds_cm': [1.5, 2, 20, 3]}],
                'relations': [], 'motion': {'transition': 'none', 'duration_ms': 0, 'groups': []}}
        self.plan = {'version': 1, 'deck_id': 'demo', 'design_spec': 'design-spec.json',
                     'narrative': {'opening': 'Current constraint', 'throughline': 'Bound the risk', 'closing_ask': 'Approve the pilot'},
                     'evidence': [{'id': 'e1', 'kind': 'user', 'path': 'source.txt'}], 'pages': [page]}

    def check(self, page=None):
        return validator.validate(self.spec, self.plan, self.base / 'design-spec.json', self.base / 'deck-plan.json', page)

    def rejected(self, fragment):
        result = self.check()
        self.assertFalse(result['valid'], result)
        self.assertTrue(any(fragment in e for e in result['errors']), result)

    def test_valid_single_page_and_extensions(self):
        self.plan['extra'] = {'allowed': True}
        self.assertTrue(self.check()['valid'])
        self.assertTrue(self.check('s01')['valid'])

    def test_narrative_and_evidence(self):
        for mutation, fragment in [(lambda: self.plan['narrative'].update(throughline=''), 'throughline'),
                                    (lambda: self.plan['pages'][0].update(evidence_refs=[]), 'evidence_refs'),
                                    (lambda: self.plan['evidence'][0].update(path='missing.csv'), 'does not exist')]:
            original = copy.deepcopy(self.plan)
            mutation()
            self.rejected(fragment)
            self.plan = original

    def test_declared_variant_both_directions(self):
        self.spec['variants']['night'] = {'pages': ['s01'], 'background_role': 'night', 'text_role': 'on_dark', 'reason': 'Closing decision emphasis'}
        self.rejected('variant')
        self.plan['pages'][0]['variant'] = 'night'
        self.assertTrue(self.check()['valid'])
        self.spec['variants']['night']['reason'] = ''
        self.rejected('reason')

    def test_bounds_font_and_nonfinite(self):
        item = self.plan['pages'][0]['objects'][0]
        item['bounds_cm'] = [30, 0, 5, 3]
        self.rejected('outside canvas')
        item['bounds_cm'] = [1, 2, 20, 3]
        item['font_size_pt'] = 12
        self.rejected('below minimum')
        item['font_size_pt'] = float('nan')
        self.rejected('finite')

    def test_semantic_object_and_relation(self):
        self.plan['pages'][0]['objects'].append({'id': 'table', 'kind': 'table', 'bounds_cm': [2, 5, 20, 5]})
        self.rejected('data_ref')
        self.plan['pages'][0]['objects'][-1]['data_ref'] = 'e1'
        self.plan['pages'][0]['relations'] = [{'from': 'body', 'to': 'missing'}]
        self.rejected('unknown object')

    def test_motion_cycle_and_limit(self):
        motion = self.plan['pages'][0]['motion']
        motion['groups'] = [{'id': 'a', 'depends_on': ['b']}, {'id': 'b', 'depends_on': ['a']}]
        self.rejected('dependency cycle')
        motion['groups'] = [{'id': str(i)} for i in range(5)]
        self.rejected('exceeds maximum')

    def test_page_mode_keeps_global_gate(self):
        self.plan['narrative']['opening'] = ''
        self.assertFalse(self.check('s01')['valid'])
        self.assertFalse(self.check('unknown')['valid'])

    def test_duplicate_order_and_morph(self):
        self.plan['pages'][0]['objects'][0]['morph_key'] = 'same'
        another = copy.deepcopy(self.plan['pages'][0]['objects'][0])
        another['id'] = 'other'
        self.plan['pages'][0]['objects'].append(another)
        self.rejected('morph_key')
        self.plan['pages'].append(copy.deepcopy(self.plan['pages'][0]))
        self.rejected('pages.order')

    def test_rhythm_is_warning_not_forced_template(self):
        for i in range(1, 4):
            p = copy.deepcopy(self.plan['pages'][0])
            p.update(id=f's0{i+1}', order=i+1)
            self.plan['pages'].append(p)
        result = self.check()
        self.assertTrue(result['valid'], result)
        self.assertEqual(len(result['warnings']), 2)

    def test_malformed_types_report_errors(self):
        self.spec['tokens']['fonts'] = []
        self.plan['pages'] = [None]
        self.assertFalse(self.check()['valid'])

    def test_json_data_selectors_validate_values_exist(self):
        (self.base / 'data.json').write_text(json.dumps({'phases': [{'amount': 18}]}), encoding='utf-8')
        self.plan['evidence'][0]['path'] = 'data.json'
        item = self.plan['pages'][0]['objects'][0]
        item['data_ref'] = 'phases[0].amount'
        self.assertTrue(self.check()['valid'])
        item['data_ref'] = 'phases[1].amount'
        self.rejected('resolve uniquely')
        item['data_ref'] = 'e1#phases[0].amount'
        self.assertTrue(self.check()['valid'])

    def test_morph_pair_and_object_ids(self):
        first = self.plan['pages'][0]
        first['objects'][0]['morph_key'] = 'focus'
        second = copy.deepcopy(first)
        second.update(id='s02', order=2)
        self.plan['pages'].append(second)
        first['relations'] = [{'type': 'morph-pair', 'target_page': 's02', 'keys': ['focus']}]
        self.assertTrue(self.check()['valid'])
        second['objects'][0]['morph_key'] = 'other'
        self.rejected('both pages')
        first['relations'] = []
        first['motion']['groups'] = [{'id': 'enter', 'object_ids': ['missing']}]
        self.rejected('unknown object')

    def test_mechanism_reuse_scope_compatible_and_exact(self):
        self.plan['pages'][0]['required_mechanisms'] = ['theme_semantics_*', {
            'query': '遮罩数字条、矩形焦点窗口', 'recipe': 'existing-entry',
            'reuse_scope': 'Only object identity and Morph interpolation',
            'excluded_capabilities': ['No fixed-window mask claim']}]
        self.assertTrue(self.check()['valid'])
        original = copy.deepcopy(self.plan['pages'][0]['required_mechanisms'][1])
        for field, value in [('query', ''), ('recipe', 7), ('reuse_scope', ''),
                             ('excluded_capabilities', 'mask'), ('excluded_capabilities', [''])]:
            self.plan['pages'][0]['required_mechanisms'][1] = {**original, field: value}
            self.rejected('required_mechanisms')
        for value in [None, 1, '', {}, {'reuse_scope': 'no query'}]:
            self.plan['pages'][0]['required_mechanisms'] = [value]
            self.rejected('required_mechanisms')

    def test_cli_refuses_report_overwrite(self):
        spec_path = self.base / 'design-spec.json'
        plan_path = self.base / 'deck-plan.json'
        spec_path.write_text(json.dumps(self.spec), encoding='utf-8')
        plan_path.write_text(json.dumps(self.plan), encoding='utf-8')
        report = self.base / 'report.json'
        cmd = [sys.executable, str(SCRIPT), '--spec', str(spec_path), '--plan', str(plan_path), '--report', str(report)]
        first = subprocess.run(cmd, capture_output=True, text=True)
        self.assertEqual(first.returncode, 0, first.stdout + first.stderr)
        original = report.read_bytes()
        second = subprocess.run(cmd, capture_output=True, text=True)
        self.assertEqual(second.returncode, 1)
        self.assertEqual(original, report.read_bytes())


if __name__ == '__main__':
    unittest.main(verbosity=2)
