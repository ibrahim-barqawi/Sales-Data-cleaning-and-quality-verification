import ast
import contextlib
import io
from pathlib import Path
import runpy
import sys
import unittest
from unittest.mock import patch
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / '02_Verification.py'
tree = ast.parse(SCRIPT.read_text())
function = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == 'verify_dataframe')
ns = {'pd': pd, 're': __import__('re')}
exec(compile(ast.Module(body=[function], type_ignores=[]), str(SCRIPT), 'exec'), ns)

class VerificationTests(unittest.TestCase):
    def setUp(self):
        self.rules = {k: [] for k in ['string_columns', 'numeric_columns', 'date_columns', 'mandatory_columns', 'unique_columns', 'non_negative_columns', 'date_relationships']}
        self.rules['allowed_values'] = {}

    def verify(self, df):
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            result = ns['verify_dataframe'](df, 'Test', self.rules)
        return result, out.getvalue()

    def test_missing_numeric(self):
        self.rules['numeric_columns'] = ['Price']
        self.assertGreater(self.verify(pd.DataFrame({'Other': [1]}))[0], 0)

    def test_nonnumeric_price(self):
        self.rules['numeric_columns'] = ['Price']
        result, out = self.verify(pd.DataFrame({'Price': ['invalid']}))
        self.assertGreater(result, 0)
        self.assertNotIn('near-zero', out)

    def test_price_boundaries(self):
        self.rules['numeric_columns'] = ['Price', 'Quantity']
        result, out = self.verify(pd.DataFrame({'Price': [-1, 0, .05, .1, .11], 'Quantity': [1, 0, 2, 3, 4]}))
        self.assertEqual(result, 0)
        self.assertIn('contains 1 zero price(s)', out)
        self.assertIn('contains 2 near-zero price(s)', out)
        self.assertIn("'Quantity' contains 1 zero value(s)", out)

    def test_missing_allowed_column(self):
        self.rules['allowed_values'] = {'Country': ['UK']}
        self.assertGreater(self.verify(pd.DataFrame({'Other': [1]}))[0], 0)

    def test_missing_date_relationship_column(self):
        self.rules['date_relationships'] = [('Start', 'End')]
        self.assertGreater(self.verify(pd.DataFrame({'Other': [1]}))[0], 0)

    def test_zero_verified_sheets(self):
        original_read = pd.read_excel
        def read(path, *args, **kwargs):
            if kwargs.get('sheet_name', 'missing') is None:
                return {'Unknown': pd.DataFrame({'Other': [1]})}
            return original_read(path, *args, **kwargs)
        out = io.StringIO()
        with patch.object(sys, 'argv', [str(SCRIPT), '--sample']), patch.object(Path, 'exists', return_value=True), patch.object(pd, 'ExcelFile') as book, patch.object(pd, 'read_excel', side_effect=read), contextlib.redirect_stdout(out):
            book.return_value.sheet_names = ['Unknown']
            with self.assertRaises(SystemExit) as result:
                runpy.run_path(str(SCRIPT), run_name='__main__')
        self.assertEqual(result.exception.code, 1)
        self.assertIn('NOT VERIFIED', out.getvalue())
        self.assertNotIn('OVERALL WORKBOOK RESULT: PASS', out.getvalue())

if __name__ == '__main__':
    unittest.main()
