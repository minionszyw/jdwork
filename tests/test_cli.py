import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from jdwork import cli
from jdwork.config import resolve_config_path
from jdwork.filtering import values_from_range, validate_filter_config, validate_backfill_config


class CliTest(unittest.TestCase):
    def test_default_config_paths(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self.assertEqual(resolve_config_path(None, "normalize", root), root / "config" / "normalize.json")
            self.assertEqual(resolve_config_path("other/filter.json", "filter", root), root / "other" / "filter.json")

    def test_commands_dispatch_with_defaults(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with patch("jdwork.config.Path.cwd", return_value=root), patch("jdwork.normalize.check") as check:
                self.assertEqual(cli.main(["normalize", "--check"]), 0)
                check.assert_called_once_with(root / "config" / "normalize.json")
            with patch("jdwork.config.Path.cwd", return_value=root), patch("jdwork.filtering.run") as run:
                self.assertEqual(cli.main(["filter", "--batch-id", "20260923150000"]), 0)
                run.assert_called_once_with(root / "config" / "filter.json", "20260923150000")
            with patch("jdwork.config.Path.cwd", return_value=root), patch("jdwork.backfill.run") as run:
                self.assertEqual(cli.main(["backfill", "--dry-run"]), 0)
                run.assert_called_once_with(root / "config" / "backfill.json", None, True)

    def test_backfill_input_is_optional(self):
        self.assertIsNone(cli.build_parser().parse_args(["backfill"]).input)


class RangeShapeTest(unittest.TestCase):
    def test_single_row_stays_one_row(self):
        self.assertEqual(values_from_range((("shop", "sku", 5),), 1, 3), [("shop", "sku", 5)])
        self.assertEqual(values_from_range(("shop", "sku", 5), 1, 3), [("shop", "sku", 5)])

    def test_single_cell_and_column(self):
        self.assertEqual(values_from_range((("sku",),), 1, 1), [("sku",)])
        self.assertEqual(values_from_range((("a",), ("b",)), 2, 1), [("a",), ("b",)])


class FilterConfigTest(unittest.TestCase):
    def test_current_filter_config_validates(self):
        with Path("config/filter.json").open(encoding="utf-8") as handle:
            validate_filter_config(json.load(handle))

    def test_rejects_filter_process_settings(self):
        with self.assertRaisesRegex(ValueError, "paths"):
            validate_filter_config({"paths": {"filter": "../data/filter"}, "filters": []})

    def test_rejects_unknown_operator(self):
        config = {"filters": [{"key": "x", "name": "X", "conditions": [{"field": "状态", "operator": "wrong"}]}]}
        with self.assertRaisesRegex(ValueError, "未知操作符"):
            validate_filter_config(config)

    def test_rejects_unsafe_backfill_field(self):
        config = {"filters": [], "backfill": {"fields": ["商品状态", "店铺"]}}
        with self.assertRaisesRegex(ValueError, "定位/元数据"):
            validate_backfill_config(config["backfill"])

    def test_allows_encoding_fields_with_stable_verification(self):
        config = {
            "filters": [],
            "backfill": {"fields": ["商家SKU", "货号"], "verify_fields": ["商品编码"]},
        }
        validate_backfill_config(config["backfill"])

    def test_rejects_backfill_verify_overlap(self):
        config = {
            "filters": [],
            "backfill": {"fields": ["货号"], "verify_fields": ["货号"]},
        }
        with self.assertRaisesRegex(ValueError, "校验字段"):
            validate_backfill_config(config["backfill"])

    def test_rejects_empty_backfill_fields(self):
        config = {"filters": [], "backfill": {"fields": []}}
        with self.assertRaisesRegex(ValueError, "非空"):
            validate_backfill_config(config["backfill"])


if __name__ == "__main__":
    unittest.main()
