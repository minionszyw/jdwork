from __future__ import annotations

import datetime as dt
import gc
import os
import re
from pathlib import Path
from typing import Any

from .config import read_json, source_path
from . import normalize

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
    raw = normalize.as_tuple(value)
    if row_count == 1:
        if len(raw) == 1 and isinstance(raw[0], tuple):
            return [raw[0]]
        return [raw]
    return [item if isinstance(item, tuple) else (item,) for item in raw]


def read_sheet(
    runner: normalize.ExcelRunner,
    path: Path,
    sheet_name: str | None,
    default_sheet: str,
    header_row: int = 1,
) -> tuple[Any, Any, list[str], list[dict[str, Any]]]:
    book = runner.open(path, read_only=True)
    ws = normalize.sheet_for(book, sheet_name, default_sheet)
    header_values, mapping = normalize.headers(ws, header_row)
    last_row = normalize.used_last_row(ws, header_row)
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
    text = normalize.clean_text(value)
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
        empty = actual is None or normalize.clean_text(actual) == ""
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
    forbidden = sorted(set(config) & {"paths", "excel", "output", "backfill"})
    if forbidden:
        raise ValueError(f"filter.json 不应包含: {', '.join(forbidden)}")
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
    # Backfill is validated by backfill.json and is intentionally independent.
    return


def validate_backfill_config(backfill: dict[str, Any]) -> None:
    fields = backfill.get("fields", [])
    if not isinstance(fields, list) or not fields or any(not isinstance(field, str) or not field for field in fields):
        raise ValueError("backfill.fields 必须是非空字符串数组")
    if len(fields) != len(set(fields)) or set(fields) & {"店铺", "类型", "SKUID"}:
        raise ValueError("backfill.fields 不能重复或包含店铺、类型、SKUID 定位/元数据字段")
    if backfill.get("key_fields", ["SKUID"]) != ["SKUID"]:
        raise ValueError("当前版本回填键必须为 ['SKUID']")
    verify_fields = backfill.get("verify_fields", ["货号"])
    if not isinstance(verify_fields, list) or not verify_fields or any(not isinstance(field, str) or not field for field in verify_fields):
        raise ValueError("backfill.verify_fields 必须是非空字符串数组")
    if set(fields) & set(verify_fields):
        overlap = ", ".join(sorted(set(fields) & set(verify_fields)))
        raise ValueError(f"backfill.fields 不能包含校验字段: {overlap}")


def validate_backfill_fields(config: dict[str, Any], normalize_config: dict[str, Any]) -> None:
    writable = set(config.get("backfill", {}).get("fields", []))
    for shop in normalize_config["sources"].get("shops", []):
        selected = shop.get("product_table", shop.get("rule", "shop_product"))
        calculated = set(normalize.rule_fields(normalize_config["rules"][selected]))
        overlap = sorted(writable & calculated)
        if overlap:
            raise ValueError(f"店铺 {shop['name']} 的回填字段包含公式列: {', '.join(overlap)}")


def load_filter_config(path: Path) -> tuple[dict[str, Any], dict[str, Any], Path, Path]:
    config = read_json(path)
    validate_filter_config(config)
    base = path.parent.resolve()
    common = read_json(base / "config.json")
    normalize_config_path = base / "normalize.json"
    normalize_config = normalize.load_config(normalize_config_path)
    paths = common.get("paths", {})
    normalize_dir = source_path(base, paths.get("normalize", "../data/normalize")).resolve()
    output_dir = source_path(base, paths.get("filter", "../data/filter")).resolve()
    config["sheet"] = common.get("sheet", {"default": "Sheet1"})
    return config, normalize_config, normalize_dir, output_dir


def batch_id(value: str | None) -> str:
    result = value or dt.datetime.now().strftime("%Y%m%d%H%M%S")
    if not re.fullmatch(r"\d{14}", result):
        raise ValueError("batch_id 必须是 14 位数字，例如 20260922171715")
    return result


def shop_input(normalize_config: dict[str, Any], normalize_dir: Path, shop: dict[str, Any]) -> tuple[Path, str | None]:
    output = shop.get("output") or f"{Path(shop['product']).stem}.xlsx"
    return normalize_dir / output, shop.get("product_sheet")


def filter_rows(config: dict[str, Any], normalize_config: dict[str, Any], normalize_dir: Path, runner: normalize.ExcelRunner) -> tuple[list[str], list[tuple[Any, ...]]]:
    rules = [rule for rule in config["filters"] if rule.get("enabled", True)]
    output_headers: list[str] | None = None
    output_rows: list[tuple[Any, ...]] = []
    for shop in normalize_config["sources"].get("shops", []):
        if not shop.get("enabled", True):
            continue
        path, sheet = shop_input(normalize_config, normalize_dir, shop)
        if not path.exists():
            raise FileNotFoundError(f"店铺 {shop['name']} 的标准化文件不存在: {path}")
        book, _, headers, rows = read_sheet(
            runner,
            path,
            sheet,
            config.get("sheet", {}).get("default", "Sheet1"),
            1,
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


def write_output(runner: normalize.ExcelRunner, output_path: Path, headers: list[str], rows: list[tuple[Any, ...]], overwrite: bool) -> None:
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
    config, normalize_config, normalize_dir, _ = load_filter_config(config_path)
    runner = normalize.ExcelRunner()
    try:
        filter_rows(config, normalize_config, normalize_dir, runner)
    finally:
        runner.shutdown()
    print("筛选配置和标准化输入检查通过")


def run(config_path: Path, requested_batch_id: str | None) -> Path:
    config, normalize_config, normalize_dir, output_dir = load_filter_config(config_path)
    runner = normalize.ExcelRunner()
    try:
        headers, rows = filter_rows(config, normalize_config, normalize_dir, runner)
        identifier = batch_id(requested_batch_id)
        filename = f"filter-{identifier}.xlsx"
        output_path = output_dir / filename
        write_output(runner, output_path, headers, rows, False)
        print(f"筛选完成: {output_path} ({len(rows)} 行)")
        return output_path
    finally:
        runner.shutdown()
