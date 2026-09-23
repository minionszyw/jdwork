import unittest

import backfill
import filter as filter_module


class FilterRulesTest(unittest.TestCase):
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
