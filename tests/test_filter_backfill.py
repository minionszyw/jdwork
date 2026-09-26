import unittest
import json
import tempfile
from pathlib import Path

from jdwork import backfill
from jdwork import filtering as filter_module


class FilterRulesTest(unittest.TestCase):
    def test_find_latest_input_uses_batch_id(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            output = root / "filter"
            output.mkdir()
            (output / "filter-20260101000000.xlsx").touch()
            (output / "filter-20260201000000.xlsx").touch()
            (output / "filter-20260201000000.tmp.xlsx").touch()
            config_path = root / "config.json"
            config_path.write_text(json.dumps({"paths": {"output": "filter"}, "output": {"prefix": "filter-"}}), encoding="utf-8")
            self.assertEqual(backfill.find_latest_input(json.loads(config_path.read_text(encoding="utf-8")), config_path), output / "filter-20260201000000.xlsx")
    def test_operators(self):
        self.assertTrue(filter_module.apply_operator("上架", "eq", "上架"))
        self.assertTrue(filter_module.apply_operator(9, "lt", 10))
        self.assertTrue(filter_module.apply_operator("0011428", "not_in", ["0020096"]))
        self.assertTrue(filter_module.apply_operator("abc-001", "contains", "001"))
        self.assertTrue(filter_module.apply_operator(None, "is_empty", None))

    def test_and_or_rules(self):
        row = {"商品状态": "上架", "库存": 5}
        self.assertTrue(filter_module.matches_rule(row, {"logic": "and", "conditions": [{"field": "商品状态", "operator": "eq", "value": "上架"}, {"field": "库存", "operator": "lt", "value": 10}]}))
        self.assertTrue(filter_module.matches_rule(row, {"logic": "or", "conditions": [{"field": "商品状态", "operator": "eq", "value": "自主下架"}, {"field": "库存", "operator": "lt", "value": 10}]}))

    def test_missing_encoding_rule_only_matches_missing_codes(self):
        with Path("config/filter.json").open(encoding="utf-8") as handle:
            config = json.load(handle)
        rule = next(item for item in config["filters"] if item["key"] == "missing_encoding")
        self.assertFalse(filter_module.matches_rule({"商家SKU": "1001", "货号": "1002"}, rule))
        self.assertTrue(filter_module.matches_rule({"商家SKU": "--", "货号": "1002"}, rule))
        self.assertTrue(filter_module.matches_rule({"商家SKU": "1001", "货号": "--"}, rule))

    def test_text_backfill_preserves_leading_zero_differences(self):
        self.assertNotEqual(backfill.normalized("00123", True), backfill.normalized("123", True))
        self.assertEqual(backfill.normalized("00123", True), "00123")

    def test_duplicate_backfill_values_merge(self):
        changes = {}
        row = {"店铺": "济喜堂", "SKUID": "sku-1", "货号": "1056471", "商品状态": "自主下架"}
        backfill.merge_changes(changes, row, ["商品状态"])
        backfill.merge_verification(changes, row, ["货号"])
        backfill.merge_changes(changes, row, ["商品状态"])
        backfill.merge_verification(changes, row, ["货号"])
        self.assertEqual(changes[("济喜堂", "sku-1")]["商品状态"], ["自主下架", "自主下架"])

    def test_duplicate_backfill_conflict(self):
        changes = {}
        row = {"店铺": "济喜堂", "SKUID": "sku-1", "货号": "1056471", "商品状态": "自主下架"}
        backfill.merge_changes(changes, row, ["商品状态"])
        backfill.merge_changes(changes, {**row, "商品状态": "上架"}, ["商品状态"])
        self.assertEqual(changes[("济喜堂", "sku-1")]["商品状态"], ["自主下架", "上架"])


if __name__ == "__main__":
    unittest.main()
