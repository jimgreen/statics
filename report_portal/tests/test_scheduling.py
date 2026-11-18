"""Isolated scheduler/API regression tests; never use the deployed database."""
import json
import tempfile
import threading
import time
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import patch
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from report_portal import app


def dt(value):
    return datetime.fromisoformat(value)


class ScheduleTimes(unittest.TestCase):
    def test_daily_timezone_and_rollover(self):
        report = dict(schedule_type="daily", schedule_time="08:30", schedule_timezone="Asia/Shanghai", enabled=True)
        self.assertEqual(app.next_schedule_run(report, dt("2026-11-18T00:00:00+00:00")), dt("2026-11-18T00:30:00+00:00"))
        self.assertEqual(app.next_schedule_run(report, dt("2026-11-18T00:30:00+00:00")), dt("2026-11-19T00:30:00+00:00"))
        report["enabled"] = False
        self.assertIsNone(app.next_schedule_run(report))

    def test_dst_gap_and_repeated_hour(self):
        report = dict(schedule_type="daily", schedule_time="02:30", schedule_timezone="Pacific/Auckland", enabled=True)
        # September gap: 02:30 is moved to 03:30 local.
        self.assertEqual(app.next_schedule_run(report, dt("2026-09-26T12:00:00+00:00")), dt("2026-09-26T14:30:00+00:00"))
        # April fold: only the first 02:30 is selected, never a second same-day run.
        first = app.next_schedule_run(report, dt("2026-04-04T12:00:00+00:00"))
        self.assertEqual(first, dt("2026-04-04T13:30:00+00:00"))
        self.assertEqual(app.next_schedule_run(report, first), dt("2026-04-05T14:30:00+00:00"))

    def test_manual_and_interval(self):
        base = dt("2026-11-18T00:00:00+00:00")
        self.assertIsNone(app.next_schedule_run(dict(schedule_type="manual"), base))
        self.assertEqual(app.next_schedule_run(dict(schedule_type="interval", schedule_minutes=7), base), base+timedelta(minutes=7))


class PortalScheduling(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.scheduler = app.Scheduler()
        root = Path(self.temp.name)
        self.patches = [patch.object(app, 'DB_PATH', root/'portal.db'), patch.object(app, 'OUTPUT_DIR', root/'outputs'), patch.object(app, 'SCHEDULER', self.scheduler)]
        for item in self.patches:
            item.start()
        app.init_db()
        self.server = app.ThreadingHTTPServer(('127.0.0.1', 0), app.PortalHandler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.url = 'http://127.0.0.1:'+str(self.server.server_address[1])
        # Deterministic ticks via dispatch_due(), with real worker threads.
        self.scheduler.running = True

    def tearDown(self):
        self.scheduler.stop()
        with self.scheduler.lock:
            threads = list(self.scheduler.active.values())
        for thread in threads:
            thread.join(5)
        self.server.shutdown()
        self.server.server_close()
        self.thread.join(3)
        for item in reversed(self.patches):
            item.stop()
        self.temp.cleanup()

    def request(self, method, path, body=None, expected=200):
        req = Request(self.url+path, data=json.dumps(body).encode() if body is not None else None,
                      method=method, headers={'Content-Type':'application/json'})
        try:
            response = urlopen(req, timeout=10)
        except HTTPError as exc:
            response = exc
        with response:
            status = response.status
            data = response.read()
        self.assertEqual(status, expected, data.decode(errors='replace'))
        return json.loads(data)

    def create(self, **extra):
        payload = dict(name='fixture', sql='SELECT :amount AS amount', parameters=[dict(name='amount', type='number', default=12, required=True)],
                       output_formats=['html','csv'], schedule_type='daily', schedule_time='08:30', schedule_timezone='Asia/Shanghai', enabled=True)
        payload.update(extra)
        return self.request('POST', '/api/reports', payload, 201)

    def wait_run(self, report_id):
        deadline = time.monotonic()+5
        while time.monotonic()<deadline:
            with self.scheduler.lock:
                active = report_id in self.scheduler.active
            if not active:
                runs=[r for r in app.run_rows() if r['report_id']==report_id]
                if runs:
                    return runs[0]
            time.sleep(.02)
        self.fail('worker did not finish')

    def due(self, report, now=None):
        now = now or app.utc_now()
        with app.connection() as conn:
            conn.execute('UPDATE reports SET next_run_at=? WHERE id=?', (app.iso(now-timedelta(minutes=1)), report['id']))
        return now

    def test_save_manual_run_and_non_schedule_edit_preserve_plan(self):
        report = self.create()
        saved = self.request('PUT', f"/api/reports/{report['id']}", {'description':'edited'})
        self.assertEqual(saved['next_run_at'], report['next_run_at'])
        self.request('POST', f"/api/reports/{report['id']}/run", {'parameters':{'amount':45}}, 202)
        run = self.wait_run(report['id'])
        self.assertEqual((run['status'], run['trigger_type'], run['parameters']), ('success','manual',{'amount':45}))
        self.assertEqual(app.get_report(report['id'])['next_run_at'], report['next_run_at'])
        self.assertEqual(self.request('GET', f"/api/runs/{run['id']}/data"), [{'amount':45}])
        changed = self.request('PUT', f"/api/reports/{report['id']}", {'schedule_time':'09:15'})
        self.assertNotEqual(changed['next_run_at'], report['next_run_at'])
        self.assertEqual(self.request('GET', f"/api/runs/{run['id']}")['schedule_snapshot']['schedule_time'], '08:30')
        disabled = self.request('PUT', f"/api/reports/{report['id']}", {'enabled':False})
        self.assertIsNone(disabled['next_run_at'])

    def test_scheduled_default_parameters_once_and_failure_not_retried(self):
        report = self.create()
        now = self.due(report)
        self.scheduler.dispatch_due(now)
        run = self.wait_run(report['id'])
        self.assertEqual((run['status'], run['trigger_type'], run['parameters']), ('success','scheduled',{'amount':12}))
        self.assertIsNotNone(run['scheduled_for'])
        self.scheduler.dispatch_due(now)
        self.assertEqual(len([r for r in app.run_rows() if r['report_id']==report['id']]),1)
        self.request('PUT', f"/api/reports/{report['id']}", {'sql':'SELECT * FROM missing_table'})
        now = self.due(report)
        self.scheduler.dispatch_due(now)
        failed = self.wait_run(report['id'])
        self.assertEqual(failed['status'],'failed')
        self.scheduler.dispatch_due(now+timedelta(seconds=2))
        self.assertEqual(len([r for r in app.run_rows() if r['report_id']==report['id']]),2)
        self.assertGreater(app.parse_iso(app.get_report(report['id'])['next_run_at']),now)

    def test_saving_while_scheduler_dispatches_preserves_advanced_time(self):
        report = self.create()
        now = self.due(report)
        validating, dispatched = threading.Event(), threading.Event()
        original_validate = app.validate_report_payload
        failures = []

        def validate(*args, **kwargs):
            result = original_validate(*args, **kwargs)
            validating.set()
            dispatched.wait(.3)
            return result

        def edit():
            try:
                self.request('PUT', f"/api/reports/{report['id']}", {'description':'concurrent edit'})
            except Exception as exc:
                failures.append(exc)

        def dispatch():
            try:
                self.scheduler.dispatch_due(now)
            except Exception as exc:
                failures.append(exc)
            finally:
                dispatched.set()

        with patch.object(app, 'validate_report_payload', validate):
            editor = threading.Thread(target=edit)
            editor.start()
            self.assertTrue(validating.wait(2))
            dispatcher = threading.Thread(target=dispatch)
            dispatcher.start()
            editor.join(5)
            dispatcher.join(5)
        self.assertFalse(editor.is_alive())
        self.assertFalse(dispatcher.is_alive())
        self.assertEqual(failures, [])
        self.assertEqual(self.wait_run(report['id'])['status'], 'success')
        saved = app.get_report(report['id'])
        self.assertEqual(saved['description'], 'concurrent edit')
        self.assertGreater(app.parse_iso(saved['next_run_at']), now)

    def test_history_formats_and_unique_output_directories(self):
        report = self.create()
        self.request('POST', f"/api/reports/{report['id']}/run", {}, 202)
        run = self.wait_run(report['id'])
        self.request('PUT', f"/api/reports/{report['id']}", {'output_formats':['csv']})
        historical = self.request('GET', f"/api/runs/{run['id']}")
        self.assertEqual(set(historical['output_formats']), {'html', 'csv'})
        with urlopen(self.url+f"/api/runs/{run['id']}/file?format=html") as response:
            self.assertEqual(response.status, 200)
        old_content = (Path(run['output_dir'])/'data.json').read_bytes()
        # A different database can allocate the same run ID into this output root.
        second = app.write_outputs(app.get_report(report['id']), [{'amount':99}], run['id'], [])
        self.assertNotEqual(second, Path(run['output_dir']))
        self.assertEqual((Path(run['output_dir'])/'data.json').read_bytes(), old_content)

    def test_interval_coalesces_missed_ticks_and_pause_allows_manual(self):
        report = self.create(schedule_type='interval', schedule_minutes=10)
        now = app.utc_now()
        with app.connection() as conn:
            conn.execute('UPDATE reports SET next_run_at=? WHERE id=?',(app.iso(now-timedelta(minutes=25)),report['id']))
        self.scheduler.pause()
        self.scheduler.dispatch_due(now)
        self.assertEqual(app.run_rows(),[])
        self.request('POST',f"/api/reports/{report['id']}/run",{},202)
        self.assertEqual(self.wait_run(report['id'])['trigger_type'],'manual')
        self.scheduler.resume()
        self.scheduler.dispatch_due(now)
        self.assertEqual(self.wait_run(report['id'])['trigger_type'],'scheduled')
        self.assertEqual(app.parse_iso(app.get_report(report['id'])['next_run_at']),now+timedelta(minutes=5))

    def test_validation_and_legacy_interval_client(self):
        for extra in [dict(schedule_time='24:00'), dict(schedule_timezone='Invalid/Zone'), dict(schedule_type='interval',schedule_minutes=0), dict(parameters=[dict(name='amount',required=True)])]:
            payload=dict(name='bad',sql='SELECT 1',schedule_type='daily',schedule_time='08:00')
            payload.update(extra)
            self.request('POST','/api/reports',payload,400)
        report = self.create(schedule_type='manual')
        saved=self.request('PUT',f"/api/reports/{report['id']}",{'schedule_minutes':5})
        self.assertEqual(saved['schedule_type'],'interval')
        saved=self.request('PUT',f"/api/reports/{report['id']}",{'schedule_minutes':0})
        self.assertEqual(saved['schedule_type'],'manual')
        self.assertIsNone(saved['next_run_at'])

    def test_duplicate_preserves_plan_but_disables_it(self):
        report=self.create()
        copied=self.request('POST',f"/api/reports/{report['id']}/duplicate",{},201)
        self.assertEqual(copied['schedule_time'], report['schedule_time'])
        self.assertEqual(copied['schedule_timezone'],'Asia/Shanghai')
        self.assertFalse(copied['enabled'])
        self.assertIsNone(copied['next_run_at'])

    def test_legacy_schema_migration_is_idempotent(self):
        report=self.create(schedule_type='interval',schedule_minutes=5)
        with app.connection() as conn:
            conn.execute('DROP INDEX IF EXISTS reports_due_idx')
            for field in ('schedule_type','schedule_time','schedule_timezone'):
                conn.execute(f'ALTER TABLE reports DROP COLUMN {field}')
        app.init_db()
        self.assertEqual(app.get_report(report['id'])['schedule_type'],'interval')
        app.init_db()
        self.assertEqual(app.get_report(report['id'])['schedule_type'],'interval')

    def test_stop_start_uses_distinct_loop_events(self):
        scheduler=app.Scheduler()
        scheduler.start()
        old_thread, old_event=scheduler.thread,scheduler.stop_event
        scheduler.stop()
        scheduler.start()
        self.assertIsNot(scheduler.stop_event,old_event)
        self.assertTrue(old_event.is_set())
        scheduler.stop()
        old_thread.join(3)
        scheduler.thread.join(3)
        self.assertFalse(old_thread.is_alive())
        self.assertFalse(scheduler.thread.is_alive())

    def test_duplicate_active_run_rejected(self):
        report=self.create()
        entered, release=threading.Event(),threading.Event()
        def slow(*args):
            entered.set()
            release.wait(3)
        with patch.object(app,'execute_report',slow):
            try:
                self.request('POST',f"/api/reports/{report['id']}/run",{},202)
                self.assertTrue(entered.wait(2))
                self.request('POST',f"/api/reports/{report['id']}/run",{},409)
            finally:
                release.set()
                with self.scheduler.lock:
                    threads=list(self.scheduler.active.values())
                for thread in threads:
                    thread.join(3)


if __name__ == '__main__':
    unittest.main()
