from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

from .config import read_json, source_path
from . import filtering as filter_module
from . import normalize as norm


def read_filter_rows(runner: norm.ExcelRunner, path: Path, sheet: str | None, header_row: int = 1) -> tuple[Any, list[str], list[dict[str, Any]]]:
    book, _, headers, rows = filter_module.read_sheet(runner, path, sheet, "Sheet1", header_row)
    return book, headers, rows


def normalized(value: Any) -> Any:
    return filter_module.comparable(value)


def merge_changes(changes: dict[tuple[str, str], dict[str, Any]], row: dict[str, Any], fields: list[str]) -> None:
    shop = norm.clean_text(row.get("店铺"))
    skuid = norm.clean_text(row.get("SKUID"))
    if not shop or not skuid:
        raise ValueError("筛选结果缺少店铺或 SKUID")
    target = changes.setdefault((shop, skuid), {})
    for field in fields:
        value = row.get(field)
        target.setdefault(field, []).append(value)


def merge_verification(changes: dict[tuple[str, str], dict[str, Any]], row: dict[str, Any], fields: list[str]) -> None:
    shop = norm.clean_text(row.get("店铺"))
    skuid = norm.clean_text(row.get("SKUID"))
    target = changes[(shop, skuid)]
    for field in fields:
        marker = f"__verify__{field}"
        value = row.get(field)
        target.setdefault(marker, []).append(value)


def load_backfill_config(path: Path) -> tuple[dict[str, Any], dict[str, Any], Path]:
    config = read_json(path)
    filter_module.validate_filter_config(config)
    base = path.parent.resolve()
    norm_path = source_path(base, config.get("paths", {}).get("norm_config", "norm.json"))
    norm_config = norm.load_config(norm_path)
    filter_module.validate_backfill_fields(config, norm_config)
    raw_dir = source_path(norm_path.parent.resolve(), norm_config.get("paths", {}).get("raw", "raw")).resolve()
    return config, norm_config, raw_dir


def locate_shop(norm_config: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {shop["name"]: shop for shop in norm_config["sources"].get("shops", [])}


def raw_row_map(ws: Any, mapping: dict[str, int], key: str, verify: list[str], header_row: int) -> tuple[dict[str, int], dict[str, dict[str, Any]]]:
    last = norm.used_last_row(ws, header_row)
    key_col = mapping[key]
    verify_cols = {field: mapping[field] for field in verify}
    values = ws.Range(ws.Cells(header_row + 1, 1), ws.Cells(last, max(mapping.values()))).Value
    rows = filter_module.values_from_range(values, last - header_row, max(mapping.values()))
    result: dict[str, int] = {}
    checks: dict[str, dict[str, Any]] = {}
    for offset, values_row in enumerate(rows):
        excel_row = header_row + 1 + offset
        value = values_row[key_col - 1] if key_col - 1 < len(values_row) else None
        identity = norm.clean_text(value)
        if not identity:
            continue
        if identity in result:
            raise ValueError(f"raw 表中 SKUID 重复，无法安全回填: {identity}")
        result[identity] = excel_row
        checks[identity] = {field: values_row[col - 1] for field, col in verify_cols.items()}
    return result, checks


def process_store(runner: norm.ExcelRunner, raw_dir: Path, shop: dict[str, Any], updates: dict[str, Any], verify_fields: list[str], dry_run: bool, backup: bool, header_row: int) -> int:
    raw_path = source_path(raw_dir, shop, "product")
    book = runner.open(raw_path, read_only=dry_run, password=shop.get("product_password"))
    try:
        ws = norm.sheet_for(book, shop.get("product_sheet"), "Sheet1")
        _, mapping = norm.headers(ws, header_row)
        editable_fields = sorted({field for values in updates.values() for field in values if not field.startswith("__verify__")})
        norm.require_columns(mapping, ["SKUID"] + verify_fields + editable_fields, ws.Name)
        row_map, checks = raw_row_map(ws, mapping, "SKUID", verify_fields, header_row)
        changed = 0
        for skuid, fields in updates.items():
            if skuid not in row_map:
                raise ValueError(f"{shop['name']} raw 表找不到 SKUID: {skuid}")
            row_number = row_map[skuid]
            for field in verify_fields:
                expected_values = fields.get(f"__verify__{field}", [])
                distinct_expected = {repr(normalized(value)): value for value in expected_values}
                if len(distinct_expected) > 1:
                    raise ValueError(f"{shop['name']} / {skuid} 的校验字段 {field} 存在冲突")
                expected = next(iter(distinct_expected.values()), None)
                if expected is not None and normalized(checks[skuid][field]) != normalized(expected):
                    raise ValueError(f"{shop['name']} / {skuid} 的校验字段 {field} 不一致")
            for field, new_value in fields.items():
                if field.startswith("__verify__"):
                    continue
                current = ws.Cells(row_number, mapping[field]).Value
                values = {repr(normalized(value)): value for value in new_value}
                candidates = [value for value in values.values() if normalized(current) != normalized(value)]
                if len(candidates) > 1:
                    raise ValueError(f"{shop['name']} / {skuid} 的字段 {field} 存在冲突修改")
                if not candidates:
                    continue
                new_value = candidates[0]
                if dry_run:
                    print(f"预览: {shop['name']} / {skuid} / {field}: {current!r} -> {new_value!r}")
                else:
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


def run(config_path: Path, input_path: Path, dry_run: bool) -> int:
    config, norm_config, raw_dir = load_backfill_config(config_path)
    fields = config.get("backfill", {}).get("fields", [])
    verify_fields = config.get("backfill", {}).get("verify_fields", ["货号"])
    if not fields:
        raise ValueError("backfill.fields 不能为空")
    runner = norm.ExcelRunner()
    try:
        book, headers, rows = read_filter_rows(
            runner,
            input_path,
            config.get("output", {}).get("sheet"),
            int(config.get("output", {}).get("header_row", 1)),
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
        shops = locate_shop(norm_config)
        grouped: dict[str, dict[str, dict[str, Any]]] = {}
        for (shop, skuid), update in changes.items():
            if shop not in shops:
                raise ValueError(f"norm.json 中找不到店铺: {shop}")
            grouped.setdefault(shop, {})[skuid] = update
        total = 0
        raw_header_row = int(norm_config.get("excel", {}).get("header_row", 1))
        for shop_name, updates in grouped.items():
            total += process_store(runner, raw_dir, shops[shop_name], updates, verify_fields, dry_run, bool(config.get("backfill", {}).get("backup", True)), raw_header_row)
        print(f"{'预览' if dry_run else '回填'}完成，共 {total} 个字段变更")
        return total
    finally:
        runner.shutdown()
