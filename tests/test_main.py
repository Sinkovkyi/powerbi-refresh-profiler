"""Regression tests for main.py on a small synthetic trace.

Run from the project folder:  python -m unittest discover tests
"""

import contextlib
import io
import os
import sys
import tempfile
import unittest
from datetime import datetime, timedelta, timezone

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import main  # noqa: E402

FIXTURE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures", "sample_refresh_trace.xml")
SHEETS = ["Summary Info", "Timeline", "System Events", "Errors", "Execution Metrics", "Unfinished"]


class SampleTraceTest(unittest.TestCase):
    """sample_refresh_trace.xml: one refresh (9.0 s) with 4 Job Graph operations on 2 threads,
    a lock wait, a commit, a service batch of the same session and failed queries of another session."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        output = os.path.join(cls.tmp.name, "refresh_analysis.xlsx")
        with contextlib.redirect_stdout(io.StringIO()):
            main.build_refresh_analysis(FIXTURE, output)
        cls.sheets = {name: pd.read_excel(output, sheet_name=name) for name in SHEETS}

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_workbook_schema(self):
        expected = {
            "Timeline": main.TIMELINE_COLUMNS,
            "System Events": main.SYSTEM_EVENTS_COLUMNS,
            "Errors": main.ERRORS_COLUMNS,
            "Execution Metrics": main.EXEC_METRICS_COLUMNS,
            "Unfinished": main.UNFINISHED_COLUMNS,
        }
        for sheet, columns in expected.items():
            self.assertEqual(list(self.sheets[sheet].columns), columns, sheet)

    def test_summary_comes_from_the_refresh_command(self):
        summary = self.sheets["Summary Info"].iloc[0]
        self.assertEqual(summary["Database"], "Contoso Sales")
        self.assertEqual(summary["User"], "Power BI Service (analyst@contoso.com)")

    def test_timeline_from_job_graph(self):
        timeline = self.sheets["Timeline"]
        self.assertEqual(len(timeline), 4)
        self.assertEqual(set(timeline["Thread_ID"]), {"Thread 1", "Thread 2"})
        self.assertEqual(set(timeline["Details"]), {
            "Process Prepare Partitions", "Process Partition", "Process AttributeHierarchy", "Process Relationship"})
        self.assertIn("Sales (prepare)", set(timeline["ObjectName"]))
        self.assertIn("Relationship", set(timeline["TableName"]))
        partition = timeline[timeline["Details"] == "Process Partition"].iloc[0]
        self.assertAlmostEqual(partition["Duration_Sec"], 3.8)
        self.assertAlmostEqual(partition["CPU_Duration_Sec"], 1.0)

    def test_system_events_do_not_overlap(self):
        events = self.sheets["System Events"].set_index("Action")
        self.assertEqual(set(events.index), {"Lock Wait", "Preparation", "Commit", "Finalization"})
        self.assertTrue(pd.isna(events.loc["Lock Wait", "Duration_Sec"]))
        self.assertAlmostEqual(events.loc["Preparation", "Duration_Sec"], 0.5)
        self.assertAlmostEqual(events.loc["Commit", "Duration_Sec"], 1.5)
        self.assertAlmostEqual(events.loc["Finalization", "Duration_Sec"], 0.5)

    def test_parts_add_up_to_engine_duration(self):
        timeline = self.sheets["Timeline"]
        start = pd.to_datetime(timeline["StartTime"]).min()
        end = pd.to_datetime(timeline["EndTime"]).max()
        span = (end - start).total_seconds()
        system = self.sheets["System Events"]["Duration_Sec"].sum()
        engine = self.sheets["Execution Metrics"]["Duration Sec"].iloc[0]
        self.assertAlmostEqual(span + system, engine, places=6)

    def test_execution_metrics_of_the_refresh_only(self):
        metrics = self.sheets["Execution Metrics"]
        self.assertEqual(len(metrics), 1)
        row = metrics.iloc[0]
        self.assertEqual(row["RequestID"], "11111111-0000-4000-8000-000000000011")
        self.assertAlmostEqual(row["Duration Sec"], 9.0)
        self.assertAlmostEqual(row["Peak RAM GB"], 1.0)
        self.assertEqual(row["Refresh Parallelism"], 2)

    def test_unfinished_objects(self):
        unfinished = self.sheets["Unfinished"]
        self.assertEqual(len(unfinished), 1)
        row = unfinished.iloc[0]
        self.assertEqual(row["EventSubclass"], 46)
        self.assertIn("relationship", row["Details"])
        # started 10:00:07, the trace ends 10:00:09.5
        self.assertAlmostEqual(row["Open_Until_Trace_End_Sec"], 2.5, places=3)

    def test_summary_busy_and_idle_time(self):
        summary = self.sheets["Summary Info"].iloc[0]
        self.assertAlmostEqual(summary["Timeline_Span_Sec"], 6.5)
        self.assertAlmostEqual(summary["Busy_Time_Sec"], 6.5)
        self.assertAlmostEqual(summary["Idle_Time_Sec"], 0.0)
        self.assertAlmostEqual(summary["Sum_Object_Duration_Sec"], 0.2 + 3.8 + 2.0 + 0.5)

    def test_other_sessions_are_not_in_errors(self):
        self.assertEqual(len(self.sheets["Errors"]), 0)


class ActivityStatsTest(unittest.TestCase):
    """Busy / idle split of the Timeline: parallel work is counted once, gaps are idle."""

    @staticmethod
    def op(start, end):
        base = datetime(2026, 9, 8, 10, 0, 0)
        return {"StartTime": base + timedelta(seconds=start), "EndTime": base + timedelta(seconds=end),
                "Duration_Sec": end - start}

    def test_parallel_work_counts_once_and_gap_is_idle(self):
        ops = [self.op(0, 4), self.op(1, 3), self.op(3, 5), self.op(8, 10)]
        stats = main.compute_activity_stats(ops)
        self.assertAlmostEqual(stats["Timeline_Span_Sec"], 10.0)
        self.assertAlmostEqual(stats["Busy_Time_Sec"], 7.0)
        self.assertAlmostEqual(stats["Idle_Time_Sec"], 3.0)
        self.assertAlmostEqual(stats["Sum_Object_Duration_Sec"], 4 + 2 + 2 + 2)

    def test_single_operation(self):
        stats = main.compute_activity_stats([self.op(2, 3)])
        self.assertAlmostEqual(stats["Busy_Time_Sec"], 1.0)
        self.assertAlmostEqual(stats["Idle_Time_Sec"], 0.0)


class SeveralAttemptsTest(unittest.TestCase):
    """A trace with earlier failed attempts: only the Rollback of the refresh request counts."""

    def test_rollbacks_of_earlier_attempts_do_not_make_finalization_negative(self):
        base = datetime(2026, 10, 14, 10, 0, 0)
        at = lambda sec: base + timedelta(seconds=sec)
        events = [{"RequestID": "NOW", "StartTime": at(11), "EndTime": at(70), "Duration_Sec": 59.0}]
        system = [
            {"Action": "Rollback", "Duration_Sec": 0.3, "StartTime": at(-100), "RequestID": "OLD1"},
            {"Action": "Rollback", "Duration_Sec": 0.4, "StartTime": at(-50), "RequestID": "OLD2"},
            {"Action": "Rollback", "Duration_Sec": 0.1, "StartTime": at(71), "RequestID": "NOW"},
        ]
        command = {"RequestID": "NOW", "Duration_Sec": 71.0}
        main.add_refresh_phases(events, system, command, {"NOW": at(0)})
        phases = {e["Action"]: e["Duration_Sec"] for e in system if e["Action"] in ("Preparation", "Finalization")}
        self.assertAlmostEqual(phases["Preparation"], 11.0)
        # 71 - 11 - 59 - 0.1 = 0.9; with the old Rollbacks it would be 0.9 - 0.7 = 0.2, with more of them below 0
        self.assertAlmostEqual(phases["Finalization"], 0.9)


class TimestampTest(unittest.TestCase):
    """Profiler timestamp formats; must parse on every supported Python version (3.10+)."""

    def test_formats(self):
        plus2 = timezone(timedelta(hours=2))
        cases = {
            "2026-09-08T14:06:52+02:00": datetime(2026, 9, 8, 14, 6, 52, tzinfo=plus2),
            "2026-09-08T12:06:52.780Z": datetime(2026, 9, 8, 12, 6, 52, 780000, tzinfo=timezone.utc),
            "2026-09-08T12:06:52.78Z": datetime(2026, 9, 8, 12, 6, 52, 780000, tzinfo=timezone.utc),
            "2026-09-08T12:06:52.4370000+02:00": datetime(2026, 9, 8, 12, 6, 52, 437000, tzinfo=plus2),
            "2026-09-08T12:06:52.4370000": datetime(2026, 9, 8, 12, 6, 52, 437000),
        }
        for text, expected in cases.items():
            self.assertEqual(main.parse_pbi_time(text), expected, text)
        self.assertIsNone(main.parse_pbi_time(""))


if __name__ == "__main__":
    unittest.main()
