import json
import unittest
from pathlib import Path

from jdwork import normalize as norm


class NormHelpersTest(unittest.TestCase):
    def test_clean_text_preserves_leading_zero(self):
        self.assertEqual(norm.clean_text("\t0020429\t"), "0020429")
        self.assertEqual(norm.clean_text(10015.0), "10015")

    def test_clean_number(self):
        self.assertEqual(norm.clean_number("1,234.50"), 1234.5)
        self.assertIsNone(norm.clean_number("--"))
        with self.assertRaises(ValueError):
            norm.clean_number("not-a-number")

    def test_column_letters(self):
        self.assertEqual(norm.col_letter(1), "A")
        self.assertEqual(norm.col_letter(26), "Z")
        self.assertEqual(norm.col_letter(27), "AA")

    def test_formula_rendering(self):
        mapping = {"数量": 10, "进价": 16, "单品进价小计": 21}
        formula = norm.render_formula(
            "={this:数量}*{this:进价}+SUM({range:单品进价小计})",
            mapping,
            2,
            100,
            {"erp_product": "'C:\\norm\\[ERP商品.xlsx]Sheet1'!$A:$M"},
        )
        self.assertEqual(formula, "=J2*P2+SUM($U$2:$U$100)")

    def test_current_config_validates(self):
        with Path("config/norm.json").open(encoding="utf-8") as handle:
            norm.validate_config(json.load(handle))


if __name__ == "__main__":
    unittest.main()
