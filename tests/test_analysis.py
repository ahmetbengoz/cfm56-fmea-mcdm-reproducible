import importlib.util
from pathlib import Path
import unittest

import numpy as np


MODULE_PATH = Path(__file__).resolve().parents[1] / "src" / "analysis.py"
SPEC = importlib.util.spec_from_file_location("analysis", MODULE_PATH)
analysis = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(analysis)


class AnalysisTests(unittest.TestCase):
    def test_parse_period_is_inclusive(self):
        self.assertEqual(analysis.parse_period("2012-2024"), (2012, 2024, 13))
        self.assertEqual(analysis.parse_period("2005"), (2005, 2005, 1))

    def test_weight_vectors_sum_to_one(self):
        matrix = analysis.minmax(np.array([[1, 5, 3], [2, 7, 4], [3, 9, 8]], dtype=float))
        self.assertTrue(np.isclose(analysis.critic_weights(matrix).sum(), 1.0))
        self.assertTrue(np.isclose(analysis.entropy_weights(matrix).sum(), 1.0))

    def test_primary_decision_matrix_is_complete(self):
        decision = analysis.load_decision_data()
        self.assertEqual(len(decision), 15)
        self.assertTrue(decision["Failure category"].is_unique)
        self.assertTrue(decision["O_window_adjusted"].between(0, 1).all())
        self.assertTrue(decision["O_raw"].between(0, 1).all())

    def test_fan_blade_is_primary_critic_topsis_rank_one(self):
        decision = analysis.load_decision_data()
        _, result = analysis.evaluate_specification(decision, "O_window_adjusted")
        fan_blade = result.loc[result["Failure category"] == "Fan blade issue"].iloc[0]
        self.assertEqual(fan_blade["TOPSIS_CRITIC_rank"], 1)
        self.assertEqual(fan_blade["VIKOR_CRITIC_rank"], 1)


if __name__ == "__main__":
    unittest.main()
