"""Rounding on the 0.05 line and the "a 0 leaves the cell as it is" option."""
import unittest

import pandas as pd

from gradefiller.automation import (
    ColumnMapping, FillJob, RoundUp, apply_round_up, build_fill_plan, step_label,
)


class RoundingTests(unittest.TestCase):
    def check(self, rule, cases):
        for given, expected in cases.items():
            with self.subTest(given=given):
                self.assertEqual(apply_round_up(given, rule), expected)

    def test_usual_rule(self):  # .1-.49 -> .5, .5-.99 -> next whole, below .1 -> .0
        self.check(RoundUp(), {
            "12": "12", "12.05": "12", "12.09": "12", "12.1": "12.5", "12.2": "12.5",
            "12.49": "12.5", "12.5": "13", "12.95": "13", "0.02": "0", "abc": "abc",
        })

    def test_half_steps(self):  # arrows 0.25 -> 0.5 -> 0.9
        self.check(RoundUp(5, 10, 18), {
            "12.2": "12", "12.24": "12", "12.25": "12.5", "12.6": "12.9",
            "12.89": "12.9", "12.9": "12.9", "12.95": "12.95",  # past the last arrow: stays
        })

    def test_one_arrow_is_plain_rounding(self):  # 0.5 -> 1 only: below 0.5 drops
        self.check(RoundUp(10, 10, 20), {"7.4": "7", "7.5": "8"})

    def test_labels(self):
        self.assertEqual([step_label(n) for n in (0, 1, 5, 10, 19, 20)], ["0", "0.05", "0.25", "0.5", "0.95", "1"])

    def test_points_in_order(self):
        with self.assertRaises(ValueError):
            RoundUp(10, 5, 20)
        with self.assertRaises(ValueError):
            RoundUp(0, 10, 21)


class ZeroTests(unittest.TestCase):
    def plan(self, zero_as_blank):
        df = pd.DataFrame({"ID": ["1", "2", "3"], "Quiz": ["0", "", "7"]})
        job = FillJob("x.xlsx", "S", ColumnMapping(0, "ID", "ID"),
                      [ColumnMapping(1, "Quiz", "Quiz")], zero_as_blank=zero_as_blank)
        return build_fill_plan(job, df, ["1", "2", "3"])

    def test_zero_left_alone_by_default(self):
        self.assertTrue(FillJob("x", "S", ColumnMapping(0, "ID", "ID")).zero_as_blank)
        plan = self.plan(True)
        self.assertEqual([(s.sid, [w.value for w in s.writes]) for s in plan.students], [("3", ["7"])])
        self.assertEqual((plan.zero_cells, plan.blank_cells), (1, 1))

    def test_zero_typed_when_unticked(self):
        plan = self.plan(False)
        self.assertEqual([[w.value for w in s.writes] for s in plan.students], [["0"], ["7"]])


if __name__ == "__main__":
    unittest.main()
