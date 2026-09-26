import json
import unittest
from pathlib import Path

from jdwork import normalize


class NormalizeHelpersTest(unittest.TestCase):
    def test_clean_text_preserves_leading_zero(self):
        self.assertEqual(normalize.clean_text("\t0020429\t"), "0020429")
        self.assertEqual(normalize.clean_text(10015.0), "10015")

    def test_clean_number(self):
        self.assertEqual(normalize.clean_number("1,234.50"), 1234.5)
        self.assertIsNone(normalize.clean_number("--"))
        with self.assertRaises(ValueError):
            normalize.clean_number("not-a-number")

    def test_column_letters(self):
        self.assertEqual(normalize.col_letter(1), "A")
        self.assertEqual(normalize.col_letter(26), "Z")
        self.assertEqual(normalize.col_letter(27), "AA")

    def test_formula_rendering(self):
        mapping = {"数量": 10, "进价": 16, "单品进价小计": 21}
        formula = normalize.render_formula(
            "={this:数量}*{this:进价}+SUM({range:单品进价小计})",
            mapping,
            2,
            100,
            {"erp_product": "'C:\\data\\normalize\\[ERP商品.xlsx]Sheet1'!$A:$M"},
        )
        self.assertEqual(formula, "=J2*P2+SUM($U$2:$U$100)")

    def test_current_config_validates(self):
        with Path("config/normalize.json").open(encoding="utf-8") as handle:
            normalize.validate_config(json.load(handle))

    def test_common_config_validates_and_composes(self):
        with Path("config/config.json").open(encoding="utf-8") as handle:
            common = json.load(handle)
        normalize.validate_config(common)
        config = normalize.load_config(Path("config/normalize.json"))
        self.assertEqual(len(config["sources"]["shops"]), 5)
        self.assertIn("erp_product", config["rules"])

    def test_rejects_old_output_path_key(self):
        with Path("config/normalize.json").open(encoding="utf-8") as handle:
            config = json.load(handle)
        config["paths"]["norm"] = "../norm"
        with self.assertRaisesRegex(ValueError, "paths.normalize"):
            normalize.validate_config(config)


if __name__ == "__main__":
    unittest.main()
