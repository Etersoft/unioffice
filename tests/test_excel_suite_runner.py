"""Reject false green legacy-suite evidence without requiring a Wine fixture."""
import importlib.util
from pathlib import Path
import unittest


spec = importlib.util.spec_from_file_location(
    'excel_suite_runner', Path(__file__).resolve().parents[1] / 'scripts/run-excel-suite.py')
runner = importlib.util.module_from_spec(spec)
spec.loader.exec_module(runner)


class SuiteEvidenceTests(unittest.TestCase):
    def test_existing_suite_results(self):
        for number, markers, counter in ((1, 32, 32), (2, 13, 13), (3, 22, 22),
                                         (4, 33, 34), (5, 32, 32), (6, 11, 12), (7, 13, 13)):
            with self.subTest(number=number):
                self.assertEqual([], runner.result_errors(
                    number, [str(counter), '0'], b'[SUCCESS]\n' * markers, 0))

    def test_empty_run_cannot_pass(self):
        self.assertTrue(runner.result_errors(1, ['0', '0'], b'', 0))

    def test_completion_cannot_hide_failed_assertions(self):
        report = b'[SUCCESS]\n' * 32 + b'!![FAILED]!! regression\n'
        self.assertTrue(runner.result_errors(1, ['32', '0'], report, 0))

    def test_truncated_report_cannot_pass(self):
        self.assertTrue(runner.result_errors(1, ['32', '0'], b'[SUCCESS]\n' * 31, 0))

    def test_invalid_completion_cannot_pass(self):
        for totals in ([], ['32'], ['32', '0', 'extra'], ['NaN', '0'], ['-1', '0'],
                       ['0', '0'], ['32', '1'], ['31', '0']):
            with self.subTest(totals=totals):
                self.assertTrue(runner.result_errors(1, totals, b'[SUCCESS]\n' * 32, 0))

    def test_failed_process_cannot_pass_with_valid_report(self):
        self.assertTrue(runner.result_errors(1, ['32', '0'], b'[SUCCESS]\n' * 32, 1))

    def test_legacy_counter_exceptions_are_limited(self):
        self.assertTrue(runner.result_errors(1, ['33', '0'], b'[SUCCESS]\n' * 32, 0))
        self.assertTrue(runner.result_errors(4, ['34', '0'], b'[SUCCESS]\n' * 34, 0))


if __name__ == '__main__':
    unittest.main()
