import json
import copy
import unittest
from pathlib import Path

from jdwork import normalize


class NormalizeHelpersTest(unittest.TestCase):
    def test_grouped_columns_preserve_runtime_order(self):
        common = json.loads(Path("config/config.json").read_text(encoding="utf-8"))
        config = json.loads(Path("config/normalize.json").read_text(encoding="utf-8"))
        composed = normalize._compose_config(config, common)
        for group in config["rules"]:
            target = group["table"]
            if target == "shop_product":
                target = composed["sources"]["shops"][0]["product_table"]
            runtime = composed["rules"][target]
            for kind in ("text", "number"):
                self.assertEqual(runtime[f"{kind}_columns"], [c["column"] for c in group["columns"] if c["type"] == "format" and c["value"] == kind])
            self.assertEqual(runtime["lookups"], [{"column": c["column"], "formula": c["value"]} for c in group["columns"] if c["type"] == "function"])

    def test_invalid_grouped_rules(self):
        valid = {"rules": [{"table": "erp_product", "columns": [{"column": "code", "type": "format", "value": "text"}]}]}
        cases = []
        duplicate_table = copy.deepcopy(valid)
        duplicate_table["rules"].append(copy.deepcopy(valid["rules"][0]))
        cases.append(duplicate_table)
        duplicate_column = copy.deepcopy(valid)
        duplicate_column["rules"][0]["columns"] *= 2
        cases.append(duplicate_column)
        for field, value in (("type", "lookup"), ("value", "date"), ("column", "")):
            invalid = copy.deepcopy(valid)
            invalid["rules"][0]["columns"][0][field] = value
            cases.append(invalid)
        cases.append({"rules": [{"table": "erp_product", "column": "code", "type": "format", "value": "text"}]})
        for config in cases:
            with self.subTest(config=config), self.assertRaises(ValueError):
                normalize.validate_config(config)

    def test_grouped_rules_reject_unknown_references(self):
        common = json.loads(Path("config/config.json").read_text(encoding="utf-8"))
        for table, formula in (("unknown", "=1"), ("erp_product", "=SUM({source:unknown})")):
            config = {"rules": [{"table": table, "columns": [{"column": "metric", "type": "function", "value": formula}]}]}
            with self.subTest(table=table), self.assertRaises(ValueError):
                normalize._compose_config(config, common)

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

    def test_rejects_legacy_normalize_config(self):
        config = {"sources": {}, "rules": {}}
        with self.assertRaisesRegex(ValueError, "只能包含 rules"):
            normalize.validate_config(config)


if __name__ == "__main__":
    unittest.main()
