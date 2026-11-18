import json
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from unittest.mock import patch

import test_scheduling as fixtures
from report_portal import app
from report_portal.result_store import result_page, write_result_store


class ResultBrowsing(unittest.TestCase):
    def test_numeric_sort_search_and_bounded_pages(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            rows = [{'n': n, '名称."': '设备%_' if n == 10 else 'Straße'} for n in range(123)]
            write_result_store(root, rows)
            page = result_page(root, {'page':['2'], 'page_size':['50']})
            self.assertEqual((page['total'], len(page['items']), page['items'][0]['n']), (123, 50, 50))
            descending = result_page(root, {'sort':['n'], 'direction':['desc'], 'page_size':['20']})
            self.assertEqual([r['n'] for r in descending['items']], list(range(122, 102, -1)))
            literal = result_page(root, {'q':['%_']})
            self.assertEqual(literal['items'], [rows[10]])
            unicode = result_page(root, {'q':['STRASSE'], 'page_size':['20']})
            self.assertEqual(unicode['filtered'], 122)
            self.assertEqual(result_page(root, {'q':['no match']})['items'], [])
            self.assertEqual(result_page(root, {'page':['999']})['page'], 3)
            # Unusual column names are mapped to an array index, never interpolated into SQL.
            self.assertEqual(result_page(root, {'sort':['名称."']})['total'], 123)
            for query in [{'page':['0']}, {'page_size':['501']}, {'sort':['n); DROP TABLE result_rows;--']}, {'direction':['bad']}]:
                with self.assertRaises(ValueError):
                    result_page(root, query)

    def test_empty_headers_and_legacy_conversion_once(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            write_result_store(root, [], ['time', 'value'])
            self.assertEqual(result_page(root, {})['columns'], ['time', 'value'])
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root/'data.json').write_text('[{"value": 123}]', encoding='utf-8')
            self.assertEqual(result_page(root, {})['items'], [{'value':123}])
            with patch('report_portal.result_store.write_result_store', side_effect=AssertionError('rebuilt')):
                self.assertEqual(result_page(root, {})['total'], 1)


class PortalBrowsing(unittest.TestCase):
    setUp = fixtures.PortalScheduling.setUp
    tearDown = fixtures.PortalScheduling.tearDown
    request = fixtures.PortalScheduling.request
    create = fixtures.PortalScheduling.create
    wait_run = fixtures.PortalScheduling.wait_run

    def test_history_search_reaches_older_than_100_and_pages_are_stable(self):
        report = self.create()
        with app.connection() as conn:
            conn.executemany('INSERT INTO runs(report_id,status,started_at,parameters,trigger_type) VALUES(?,?,?,?,?)', [
                (report['id'], 'failed' if i % 2 else 'success', '2026-11-18T00:00:00+00:00',
                 json.dumps({'key':'old needle' if i == 0 else i}), 'manual' if i % 2 else 'scheduled') for i in range(125)])
        self.assertEqual(len(self.request('GET', '/api/runs')['items']), 100)
        one = self.request('GET', '/api/runs?page=1&page_size=20')
        two = self.request('GET', '/api/runs?page=2&page_size=20')
        self.assertEqual((one['total'], one['success'], one['failed']), (125, 63, 62))
        self.assertFalse({r['id'] for r in one['items']} & {r['id'] for r in two['items']})
        old = self.request('GET', '/api/runs?page=1&q=old%20needle')
        self.assertEqual(old['total'], 1)
        self.assertEqual(old['items'][0]['id'], 1)
        filtered = self.request('GET', '/api/runs?page=999&status=success&trigger_type=scheduled')
        self.assertEqual((filtered['total'], filtered['page']), (63, 2))
        self.request('GET', '/api/runs?page=-1', expected=400)
        self.request('GET', '/api/runs?status=unknown', expected=400)

    def test_summary_omits_large_layout_and_results_keep_legacy_api(self):
        report = self.create()
        summary = self.request('GET', '/api/reports?summary=1')['items']
        self.assertTrue(all('layout' not in r and 'sql' not in r for r in summary))
        self.assertIn('layout', self.request('GET', f"/api/reports/{report['id']}"))
        self.request('POST', f"/api/reports/{report['id']}/run", {}, 202)
        run = self.wait_run(report['id'])
        self.assertEqual(self.request('GET', f"/api/runs/{run['id']}/data"), [{'amount':12}])
        page = self.request('GET', f"/api/runs/{run['id']}/data?page=1")
        self.assertEqual(page['items'], [{'amount':12}])
        self.assertGreaterEqual(run['metrics']['duration_ms'], run['metrics']['query_ms'])

    def test_sources_reused_across_cells_and_formats_but_not_runs(self):
        source = {'type':'sql', 'sql':'SELECT :amount AS n'}
        layout = {'rows':1, 'cols':3, 'cells':{f'r0c{i}':{'element':'value', 'data_source':source} for i in range(3)}}
        first = self.create(layout=layout, output_formats=['html','xlsx','docx'])
        second = self.create(name='second', layout=layout, output_formats=['html','xlsx','docx'])
        with ThreadPoolExecutor(max_workers=2) as pool:
            jobs = [pool.submit(app.execute_report, report['id'], {'amount':value}) for report, value in [(first, 31), (second, 87)]]
            runs = [job.result(timeout=15) for job in jobs]
        for run, value in zip(runs, [31,87]):
            self.assertEqual(run['status'], 'success')
            self.assertEqual(run['metrics']['source_queries'], 1)
            self.assertEqual(run['metrics']['source_requests'], 9)
            self.assertEqual(run['metrics']['cache_hits'], 8)
            self.assertEqual(result_page(Path(run['output_dir']), {})['items'], [{'amount':value}])
            from openpyxl import load_workbook
            workbook = load_workbook(Path(run['output_dir'])/'report.xlsx', read_only=True)
            self.assertEqual(workbook['版式']['A1'].value, value)
            workbook.close()

    def test_scheduler_does_not_decode_unscheduled_layouts(self):
        self.create(schedule_type='manual')
        with patch.object(app, 'row_to_report', side_effect=AssertionError('decoded idle definition')):
            self.scheduler.dispatch_due()


if __name__ == '__main__':
    unittest.main()
