from __future__ import annotations

import datetime as dt
import gc
import os
import re
from pathlib import Path
from typing import Any

from .config import read_json, source_path
from . import normalize as norm

OPERATORS = {"eq", "ne", "lt", "lte", "gt", "gte", "in", "not_in", "contains", "not_contains", "is_empty", "not_empty"}


def values_from_range(value: Any, row_count: int, col_count: int) -> list[tuple]:
    """Return a two-dimensional COM range without collapsing a single row."""
    if row_count <= 0:
        return []
    if row_count == 1 and col_count == 1:
        if isinstance(value, tuple) and len(value) == 1:
            item = value[0]
            return [(item[0] if isinstance(item, tuple) else item,)]
        return [(value,)]
    raw = norm.as_tuple(value)
    if row_count == 1:
        if len(raw) == 1 and isinstance(raw[0], tuple):
            return [raw[0]]
        return [raw]
    return [item if isinstance(item, tuple) else (item,) for item in raw]


def read_sheet(
    runner: norm.ExcelRunner,
    path: Path,
    sheet_name: str | None,
    default_sheet: str,
    header_row: int = 1,
) -> tuple[Any, Any, list[str], list[dict[str, Any]]]:
    book = runner.open(path, read_only=True)
    ws = norm.sheet_for(book, sheet_name, default_sheet)
    header_values, mapping = norm.headers(ws, header_row)
    last_row = norm.used_last_row(ws, header_row)
    last_col = max(mapping.values(), default=0)
    if last_row < header_row + 1 or last_col == 0:
        return book, ws, header_values, []
    values = ws.Range(ws.Cells(header_row + 1, 1), ws.Cells(last_row, last_col)).Value
    rows = []
    for values_row in values_from_range(values, last_row - header_row, last_col):
        row = {header_values[index]: values_row[index] if index < len(values_row) else None for index in range(len(header_values)) if header_values[index]}
        rows.append(row)
    return book, ws, header_values, rows


def comparable(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    text = norm.clean_text(value)
    if text is None or text == "":
        return None
    try:
        number = float(text.replace(",", ""))
        return number
    except ValueError:
        return text


def equal_value(left: Any, right: Any) -> bool:
    return comparable(left) == comparable(right)


def apply_operator(actual: Any, operator: str, expected: Any) -> bool:
    if operator in {"is_empty", "not_empty"}:
        empty = actual is None or norm.clean_text(actual) == ""
        return empty if operator == "is_empty" else not empty
    if operator in {"in", "not_in"}:
        if not isinstance(expected, list):
            raise ValueError(f"{operator} 的 value 必须是数组")
        result = any(equal_value(actual, item) for item in expected)
        return result if operator == "in" else not result
    if operator in {"contains", "not_contains"}:
        left = "" if actual is None else str(actual)
        result = str(expected) in left
        return result if operator == "contains" else not result
    if operator not in {"eq", "ne", "lt", "lte", "gt", "gte"}:
        raise ValueError(f"未知筛选操作符: {operator}")
    left = comparable(actual)
    right = comparable(expected)
    if operator == "eq":
        return left == right
    if operator == "ne":
        return left != right
    if left is None or right is None:
        return False
    try:
        if operator == "lt":
            return left < right
        if operator == "lte":
            return left <= right
        if operator == "gt":
            return left > right
        return left >= right
    except TypeError as exc:
        raise ValueError(f"无法比较值: {actual!r} {operator} {expected!r}") from exc


def matches_rule(row: dict[str, Any], rule: dict[str, Any]) -> bool:
    conditions = rule.get("conditions", [])
    results = []
    for condition in conditions:
        field = condition.get("field")
        if field not in row:
            raise ValueError(f"筛选字段不存在: {field}")
        results.append(apply_operator(row[field], condition.get("operator", "eq"), condition.get("value")))
    if rule.get("logic", "and") == "or":
        return any(results)
    if rule.get("logic", "and") != "and":
        raise ValueError(f"未知规则逻辑: {rule.get('logic')}")
    return all(results)


def validate_filter_config(config: dict[str, Any]) -> None:
    if not isinstance(config.get("filters"), list):
        raise ValueError("filter.json 必须包含 filters 数组")
    keys = set()
    for index, rule in enumerate(config["filters"]):
        if not isinstance(rule, dict) or not rule.get("key") or not rule.get("name"):
            raise ValueError(f"filters[{index}] 必须包含 key 和 name")
        if rule["key"] in keys:
            raise ValueError(f"筛选规则 key 重复: {rule['key']}")
        keys.add(rule["key"])
        if rule.get("logic", "and") not in {"and", "or"}:
            raise ValueError(f"规则 {rule['key']} 的 logic 必须是 and 或 or")
        if not isinstance(rule.get("conditions", []), list):
            raise ValueError(f"规则 {rule['key']} 的 conditions 必须是数组")
        for position, condition in enumerate(rule.get("conditions", [])):
            if not isinstance(condition, dict) or not isinstance(condition.get("field"), str) or not condition["field"]:
                raise ValueError(f"规则 {rule['key']} 的条件 {position} 必须包含 field")
            operator = condition.get("operator", "eq")
            if operator not in OPERATORS:
                raise ValueError(f"规则 {rule['key']} 使用未知操作符: {operator}")
            if operator in {"in", "not_in"} and not isinstance(condition.get("value"), list):
                raise ValueError(f"规则 {rule['key']} 的 {operator} 条件必须使用数组 value")
    backfill = config.get("backfill", {})
    if not isinstance(backfill, dict):
        raise ValueError("backfill 必须是对象")
    fields = backfill.get("fields", [])
    if not isinstance(fields, list) or any(not isinstance(field, str) or not field for field in fields):
        raise ValueError("backfill.fields 必须是非空字符串数组")
    if len(fields) != len(set(fields)) or set(fields) & {"店铺", "类型", "SKUID", "货号"}:
        raise ValueError("backfill.fields 不能重复或包含定位/元数据字段")
    if backfill.get("key_fields", ["SKUID"]) != ["SKUID"]:
        raise ValueError("当前版本回填键必须为 ['SKUID']")
    verify_fields = backfill.get("verify_fields", ["货号"])
    if not isinstance(verify_fields, list) or any(not isinstance(field, str) or not field for field in verify_fields):
        raise ValueError("backfill.verify_fields 必须是非空字符串数组")


def validate_backfill_fields(config: dict[str, Any], norm_config: dict[str, Any]) -> None:
    writable = set(config.get("backfill", {}).get("fields", []))
    for shop in norm_config["sources"].get("shops", []):
        selected = shop.get("rule", "shop_product")
        calculated = set(norm.rule_fields(norm_config["rules"][selected]))
        overlap = sorted(writable & calculated)
        if overlap:
            raise ValueError(f"店铺 {shop['name']} 的回填字段包含公式列: {', '.join(overlap)}")


def load_filter_config(path: Path) -> tuple[dict[str, Any], dict[str, Any], Path, Path]:
    config = read_json(path)
    validate_filter_config(config)
    base = path.parent.resolve()
    norm_config_path = source_path(base, config.get("paths", {}).get("norm_config", "norm.json"))
    norm_config = norm.load_config(norm_config_path)
    validate_backfill_fields(config, norm_config)
    norm_base = norm_config_path.parent.resolve()
    norm_dir = source_path(norm_base, norm_config.get("paths", {}).get("norm", "norm")).resolve()
    output_dir = source_path(base, config.get("paths", {}).get("output", "filter")).resolve()
    return config, norm_config, norm_dir, output_dir


def batch_id(value: str | None) -> str:
    result = value or dt.datetime.now().strftime("%Y%m%d%H%M%S")
    if not re.fullmatch(r"\d{14}", result):
        raise ValueError("batch_id 必须是 14 位数字，例如 20260922171715")
    return result


def shop_input(norm_config: dict[str, Any], norm_dir: Path, shop: dict[str, Any]) -> tuple[Path, str | None]:
    output = shop.get("output") or f"{Path(shop['product']).stem}.xlsx"
    return norm_dir / output, shop.get("product_sheet")


def filter_rows(config: dict[str, Any], norm_config: dict[str, Any], norm_dir: Path, runner: norm.ExcelRunner) -> tuple[list[str], list[tuple[Any, ...]]]:
    rules = [rule for rule in config["filters"] if rule.get("enabled", True)]
    output_headers: list[str] | None = None
    output_rows: list[tuple[Any, ...]] = []
    for shop in norm_config["sources"].get("shops", []):
        if not shop.get("enabled", True):
            continue
        path, sheet = shop_input(norm_config, norm_dir, shop)
        if not path.exists():
            raise FileNotFoundError(f"店铺 {shop['name']} 的标准化文件不存在: {path}")
        book, _, headers, rows = read_sheet(
            runner,
            path,
            sheet,
            config.get("excel", {}).get("default_sheet", "Sheet1"),
            int(config.get("excel", {}).get("header_row", 1)),
        )
        try:
            if output_headers is not None and output_headers[2:] != headers:
                raise ValueError(f"{path.name} 的字段顺序与其他店铺商品表不一致")
            for rule in rules:
                missing = sorted({condition.get("field") for condition in rule.get("conditions", [])} - set(headers))
                if missing:
                    raise ValueError(f"规则 {rule['key']} 在 {path.name} 缺少字段: {', '.join(missing)}")
                for row in rows:
                    if matches_rule(row, rule):
                        if output_headers is None:
                            output_headers = ["店铺", "类型"] + headers
                        output_rows.append((shop["name"], rule["name"], *[row.get(header) for header in headers]))
        finally:
            runner.close(book, False)
    return output_headers or ["店铺", "类型"], output_rows


def write_output(runner: norm.ExcelRunner, output_path: Path, headers: list[str], rows: list[tuple[Any, ...]], overwrite: bool) -> None:
    if output_path.exists() and not overwrite:
        raise FileExistsError(f"筛选输出已存在: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = output_path.with_name(f".{output_path.stem}.tmp.xlsx")
    if temporary.exists():
        temporary.unlink()
    book = runner.excel.Workbooks.Add()
    runner.open_books.append(book)
    try:
        ws = book.Worksheets(1)
        all_values = [tuple(headers)] + rows
        end_row = len(all_values)
        end_col = len(headers)
        for index, field in enumerate(headers, 1):
            if field in {"店铺", "类型", "SKUID", "商品编码", "商家SKU", "货号"}:
                try:
                    ws.Range(ws.Cells(1, index), ws.Cells(end_row, index)).NumberFormat = "@"
                except Exception:
                    pass
        ws.Range(ws.Cells(1, 1), ws.Cells(end_row, end_col)).Value = tuple(all_values)
        book.SaveAs(str(temporary.resolve()), FileFormat=51)
        runner.close(book, False)
        book = None
        ws = None
        gc.collect()
        os.replace(temporary, output_path)
    except Exception:
        if book in runner.open_books:
            runner.close(book, False)
        if temporary.exists():
            temporary.unlink()
        raise


def check(config_path: Path) -> None:
    config, norm_config, norm_dir, _ = load_filter_config(config_path)
    runner = norm.ExcelRunner()
    try:
        filter_rows(config, norm_config, norm_dir, runner)
    finally:
        runner.shutdown()
    print("筛选配置和标准化输入检查通过")


def run(config_path: Path, requested_batch_id: str | None) -> Path:
    config, norm_config, norm_dir, output_dir = load_filter_config(config_path)
    runner = norm.ExcelRunner()
    try:
        headers, rows = filter_rows(config, norm_config, norm_dir, runner)
        identifier = batch_id(requested_batch_id)
        filename = f"{config.get('output', {}).get('prefix', 'filter-')}{identifier}.xlsx"
        output_path = output_dir / filename
        write_output(runner, output_path, headers, rows, bool(config.get("output", {}).get("overwrite", False)))
        print(f"筛选完成: {output_path} ({len(rows)} 行)")
        return output_path
    finally:
        runner.shutdown()
