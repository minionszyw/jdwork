from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Any

from .config import read_json, source_path
from . import filtering as filter_module
from . import normalize


def read_filter_rows(runner: normalize.ExcelRunner, path: Path, sheet: str | None, header_row: int = 1) -> tuple[Any, list[str], list[dict[str, Any]]]:
    book, _, headers, rows = filter_module.read_sheet(runner, path, sheet, "Sheet1", header_row)
    return book, headers, rows


def normalized(value: Any, text_field: bool = False) -> Any:
    return normalize.clean_text(value) if text_field else filter_module.comparable(value)


def merge_changes(changes: dict[tuple[str, str], dict[str, Any]], row: dict[str, Any], fields: list[str]) -> None:
    shop = normalize.clean_text(row.get("店铺"))
    skuid = normalize.clean_text(row.get("SKUID"))
    if not shop or not skuid:
        raise ValueError("筛选结果缺少店铺或 SKUID")
    target = changes.setdefault((shop, skuid), {})
    for field in fields:
        value = row.get(field)
        target.setdefault(field, []).append(value)


def merge_verification(changes: dict[tuple[str, str], dict[str, Any]], row: dict[str, Any], fields: list[str]) -> None:
    shop = normalize.clean_text(row.get("店铺"))
    skuid = normalize.clean_text(row.get("SKUID"))
    target = changes[(shop, skuid)]
    for field in fields:
        marker = f"__verify__{field}"
        value = row.get(field)
        target.setdefault(marker, []).append(value)


def load_backfill_config(path: Path) -> tuple[dict[str, Any], dict[str, Any], Path, dict[str, Any]]:
    config = read_json(path)
    filter_module.validate_backfill_config(config)
    base = path.parent.resolve()
    common = read_json(base / "config.json")
    normalize_path = base / "normalize.json"
    normalize_config = normalize.load_config(normalize_path)
    filter_module.validate_backfill_fields(config, normalize_config)
    raw_dir = source_path(base, common.get("paths", {}).get("raw", "../data/raw")).resolve()
    return config, normalize_config, raw_dir, common.get("sheet", {"default": "Sheet1"})


def locate_shop(normalize_config: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {shop["name"]: shop for shop in normalize_config["sources"].get("shops", [])}


def raw_row_map(ws: Any, mapping: dict[str, int], key: str, verify: list[str], header_row: int) -> tuple[dict[str, int], dict[str, dict[str, Any]]]:
    last = normalize.used_last_row(ws, header_row)
    key_col = mapping[key]
    verify_cols = {field: mapping[field] for field in verify}
    values = ws.Range(ws.Cells(header_row + 1, 1), ws.Cells(last, max(mapping.values()))).Value
    rows = filter_module.values_from_range(values, last - header_row, max(mapping.values()))
    result: dict[str, int] = {}
    checks: dict[str, dict[str, Any]] = {}
    for offset, values_row in enumerate(rows):
        excel_row = header_row + 1 + offset
        value = values_row[key_col - 1] if key_col - 1 < len(values_row) else None
        identity = normalize.clean_text(value)
        if not identity:
            continue
        if identity in result:
            raise ValueError(f"raw 表中 SKUID 重复，无法安全回填: {identity}")
        result[identity] = excel_row
        checks[identity] = {field: values_row[col - 1] for field, col in verify_cols.items()}
    return result, checks


def process_store(runner: normalize.ExcelRunner, raw_dir: Path, shop: dict[str, Any], updates: dict[str, Any], verify_fields: list[str], text_fields: set[str], dry_run: bool, backup: bool, header_row: int) -> int:
    raw_path = source_path(raw_dir, shop, "product")
    book = runner.open(raw_path, read_only=dry_run, password=shop.get("product_password"))
    try:
        ws = normalize.sheet_for(book, shop.get("product_sheet"), "Sheet1")
        _, mapping = normalize.headers(ws, header_row)
        editable_fields = sorted({field for values in updates.values() for field in values if not field.startswith("__verify__")})
        normalize.require_columns(mapping, ["SKUID"] + verify_fields + editable_fields, ws.Name)
        row_map, checks = raw_row_map(ws, mapping, "SKUID", verify_fields, header_row)
        changed = 0
        for skuid, fields in updates.items():
            if skuid not in row_map:
                raise ValueError(f"{shop['name']} raw 表找不到 SKUID: {skuid}")
            row_number = row_map[skuid]
            for field in verify_fields:
                expected_values = fields.get(f"__verify__{field}", [])
                distinct_expected = {repr(normalized(value, True)): value for value in expected_values}
                if len(distinct_expected) > 1:
                    raise ValueError(f"{shop['name']} / {skuid} 的校验字段 {field} 存在冲突")
                expected = next(iter(distinct_expected.values()), None)
                if normalized(checks[skuid][field], True) != normalized(expected, True):
                    raise ValueError(f"{shop['name']} / {skuid} 的校验字段 {field} 不一致")
            for field, new_value in fields.items():
                if field.startswith("__verify__"):
                    continue
                current = ws.Cells(row_number, mapping[field]).Value
                text_field = field in text_fields
                values = {repr(normalized(value, text_field)): value for value in new_value}
                candidates = [value for value in values.values() if normalized(current, text_field) != normalized(value, text_field)]
                if len(candidates) > 1:
                    raise ValueError(f"{shop['name']} / {skuid} 的字段 {field} 存在冲突修改")
                if not candidates:
                    continue
                new_value = candidates[0]
                if dry_run:
                    print(f"预览: {shop['name']} / {skuid} / {field}: {current!r} -> {new_value!r}")
                else:
                    if text_field:
                        ws.Cells(row_number, mapping[field]).NumberFormat = "@"
                    ws.Cells(row_number, mapping[field]).Value = new_value
                    print(f"回填: {shop['name']} / {skuid} / {field}: {current!r} -> {new_value!r}")
                changed += 1
        if not dry_run and changed:
            if backup:
                backup_path = raw_path.with_name(raw_path.name + ".bak")
                shutil.copy2(raw_path, backup_path)
            book.Save()
        return changed
    finally:
        runner.close(book, False)


def find_latest_input(config_path: Path) -> Path:
    """Select the filter workbook with the greatest filename batch id."""
    base = config_path.parent.resolve()
    common = read_json(base / "config.json")
    output_dir = source_path(base, common.get("paths", {}).get("filter", "../data/filter")).resolve()
    prefix = "filter-"
    pattern = re.compile(rf"^{re.escape(prefix)}(\d{{14}})\.xlsx$")
    candidates: list[tuple[str, Path]] = []
    if output_dir.exists():
        for path in output_dir.iterdir():
            if not path.is_file():
                continue
            match = pattern.fullmatch(path.name)
            if match:
                candidates.append((match.group(1), path))
    if not candidates:
        raise FileNotFoundError(
            f"找不到筛选文件: {output_dir}\\{prefix}{{14 位 batch_id}}.xlsx"
        )
    return max(candidates, key=lambda item: item[0])[1].resolve()


def run(config_path: Path, input_path: Path | None, dry_run: bool) -> int:
    config, normalize_config, raw_dir, sheet = load_backfill_config(config_path)
    if input_path is None:
        input_path = find_latest_input(config_path)
    else:
        input_path = input_path.resolve()
    print(f"使用筛选文件: {input_path}")
    fields = config.get("fields", [])
    verify_fields = config.get("verify_fields", ["货号"])
    if not fields:
        raise ValueError("backfill.fields 不能为空")
    runner = normalize.ExcelRunner()
    try:
        book, headers, rows = read_filter_rows(
            runner,
            input_path,
            sheet.get("default", "Sheet1"),
            1,
        )
        try:
            required = ["店铺", "类型", "SKUID", *verify_fields, *fields]
            missing = [field for field in required if field not in headers]
            if missing:
                raise ValueError(f"筛选文件缺少字段: {', '.join(sorted(set(missing)))}")
        finally:
            runner.close(book, False)

        changes: dict[tuple[str, str], dict[str, Any]] = {}
        for row in rows:
            merge_changes(changes, row, fields)
            merge_verification(changes, row, verify_fields)
        shops = locate_shop(normalize_config)
        grouped: dict[str, dict[str, dict[str, Any]]] = {}
        for (shop, skuid), update in changes.items():
            if shop not in shops:
                raise ValueError(f"normalize.json 中找不到店铺: {shop}")
            grouped.setdefault(shop, {})[skuid] = update
        total = 0
        raw_header_row = 1
        for shop_name, updates in grouped.items():
            shop = shops[shop_name]
            rule = normalize_config["rules"][shop.get("product_table", shop.get("rule", "shop_product"))]
            text_fields = set(rule.get("text_columns", []))
            total += process_store(runner, raw_dir, shop, updates, verify_fields, text_fields, dry_run, bool(config.get("backup", True)), raw_header_row)
        print(f"{'预览' if dry_run else '回填'}完成，共 {total} 个字段变更")
        return total
    finally:
        runner.shutdown()
