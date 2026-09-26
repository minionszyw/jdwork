from __future__ import annotations

import gc
import math
import os
import re
import time
from pathlib import Path
from typing import Any, Callable

from .config import read_json, source_path
from .excel import ExcelRunner, as_tuple, clean_text, headers, require_columns, sheet_for, used_last_row

XL_OPEN_XML_WORKBOOK = 51
TABLE_TYPES = {"erp", "shop_product", "shop_sales"}

def clean_number(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError("布尔值不是有效数字")
    if isinstance(value, (int, float)):
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError("数字不是有限值")
        return value
    text = clean_text(value)
    if not text or text.lower() in {"null", "none", "--", "-"}:
        return None
    try:
        number = float(text.replace(",", ""))
    except ValueError as exc:
        raise ValueError(f"无法转换为数字: {value!r}") from exc
    if not math.isfinite(number):
        raise ValueError(f"数字不是有限值: {value!r}")
    return int(number) if number.is_integer() else number

def col_letter(col: int) -> str:
    if col < 1:
        raise ValueError("列号必须大于 0")
    result = ""
    while col:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def rule_fields(rule: dict[str, Any]) -> list[str]:
    return [item["column"] for item in rule.get("lookups", []) + rule.get("calculations", [])]


def grouped_actions(config: dict[str, Any]) -> list[dict[str, str]]:
    """Validate table groups and flatten them for the shared runtime model."""
    groups = config.get("rules")
    if not isinstance(groups, list):
        raise ValueError("normalize.json 必须包含 rules 数组")
    actions: list[dict[str, str]] = []
    seen_tables: set[str] = set()
    for index, group in enumerate(groups):
        prefix = f"rules[{index}]"
        if not isinstance(group, dict) or set(group) != {"table", "columns"}:
            raise ValueError(f"{prefix} 只能包含 table 和 columns")
        table = group["table"]
        if not isinstance(table, str) or not table.strip():
            raise ValueError(f"{prefix}.table 必须是非空字符串")
        if table in seen_tables:
            raise ValueError(f"table 重复: {table}")
        seen_tables.add(table)
        if not isinstance(group["columns"], list):
            raise ValueError(f"{prefix}.columns 必须是数组")
        seen_columns: set[str] = set()
        for position, column in enumerate(group["columns"]):
            location = f"{prefix}.columns[{position}]"
            if not isinstance(column, dict) or set(column) != {"column", "type", "value"}:
                raise ValueError(f"{location} 只能包含 column、type、value")
            name = column["column"]
            if not isinstance(name, str) or not name.strip():
                raise ValueError(f"{location}.column 必须是非空字符串")
            if name in seen_columns:
                raise ValueError(f"{table} 的 column 重复: {name}")
            seen_columns.add(name)
            kind, value = column["type"], column["value"]
            if kind not in ("format", "function"):
                raise ValueError(f"{location}.type 必须是 format 或 function")
            if kind == "format" and value not in ("text", "number"):
                raise ValueError(f"{location}.value 必须是 text 或 number")
            if kind == "function" and (not isinstance(value, str) or not value.strip()):
                raise ValueError(f"{location}.value 必须是非空 Excel 公式")
            actions.append({"table": table, **column})
    return actions


def validate_config(config: dict[str, Any]) -> None:
    if "tables" in config:
        validate_common_config(config)
        return
    if set(config) != {"rules"}:
        raise ValueError("normalize.json 只能包含 rules")
    grouped_actions(config)

def validate_common_config(config: dict[str, Any]) -> None:
    paths = config.get("paths")
    if not isinstance(paths, dict) or any(not paths.get(key) for key in ("raw", "normalize", "filter")):
        raise ValueError("公共配置 paths 必须包含 raw、normalize、filter")
    sheet = config.get("sheet")
    if not isinstance(sheet, dict) or not sheet.get("default"):
        raise ValueError("公共配置 sheet.default 必须是非空字符串")
    tables = config.get("tables")
    if not isinstance(tables, list) or not tables:
        raise ValueError("公共配置必须包含非空 tables 数组")
    if set(config) != {"paths", "tables", "sheet"}:
        raise ValueError("config.json 只能包含 paths、tables、sheet")
    names: set[str] = set()
    shops: dict[str, list[str]] = {}
    for index, item in enumerate(tables):
        prefix = f"tables[{index}]"
        if not isinstance(item, dict):
            raise ValueError(f"{prefix} 必须是对象")
        missing = [key for key in ("table", "name", "type", "file", "shop") if key not in item or not isinstance(item[key], str)]
        if missing:
            raise ValueError(f"{prefix} 缺少字段: {', '.join(missing)}")
        if set(item) != {"table", "name", "type", "file", "shop"}:
            raise ValueError(f"{prefix} 只能包含 table、name、type、file、shop")
        table = str(item["table"])
        if table in names:
            raise ValueError(f"table 重复: {table}")
        names.add(table)
        kind = item["type"]
        if kind not in TABLE_TYPES:
            raise ValueError(f"{prefix}.type 必须是 erp、shop_product 或 shop_sales")
        if kind.startswith("shop_"):
            shop = item.get("shop")
            if not isinstance(shop, str) or not shop:
                raise ValueError(f"{prefix} 的 shop_product/shop_sales 必须配置 shop")
            shops.setdefault(shop, []).append(kind)
    for shop, kinds in shops.items():
        if len(kinds) != len(set(kinds)):
            raise ValueError(f"店铺 {shop} 的商品/销售表类型重复")
def set_column_values(ws: Any, col: int, first_row: int, last_row: int, converter: Callable[[Any], Any], field: str) -> None:
    if last_row < first_row:
        return
    values = ws.Range(ws.Cells(first_row, col), ws.Cells(last_row, col)).Value
    output = []
    for offset, item in enumerate(as_tuple(values)):
        value = item[0] if isinstance(item, tuple) else item
        try:
            converted = converter(value)
        except ValueError as exc:
            raise ValueError(f"{ws.Name}!{field} 第 {first_row + offset} 行: {exc}") from exc
        output.append((converted,))
    ws.Range(ws.Cells(first_row, col), ws.Cells(last_row, col)).Value = tuple(output)


def set_number_format(ws: Any, col: int, first: int, last: int, value: str, field: str) -> None:
    try:
        ws.Range(ws.Cells(first, col), ws.Cells(last, col)).NumberFormat = value
    except Exception as exc:
        print(f"警告: 无法设置 {ws.Name}!{field} 的格式 {value!r}: {exc}")


def format_columns(ws: Any, mapping: dict[str, int], rule: dict[str, Any], first: int, last: int) -> None:
    text_columns = rule.get("text_columns", [])
    number_columns = rule.get("number_columns", [])
    require_columns(mapping, text_columns + number_columns, ws.Name)
    for name in text_columns:
        col = mapping[name]
        set_number_format(ws, col, first, last, "@", name)
        set_column_values(ws, col, first, last, clean_text, name)
    for name in number_columns:
        col = mapping[name]
        set_column_values(ws, col, first, last, clean_number, name)


def external_ref(path: Path, sheet: str, cell_range: str) -> str:
    text = f"{path.resolve().parent}\\[{path.name}]{sheet}".replace("'", "''")
    return f"'{text}'!{cell_range}"


def render_formula(template: str, mapping: dict[str, int], row: int, last: int, sources: dict[str, str]) -> str:
    def current_cell(match: re.Match[str]) -> str:
        name = match.group(1)
        if name not in mapping:
            raise ValueError(f"公式引用了不存在的字段: {name}")
        return f"{col_letter(mapping[name])}{row}"

    def current_range(match: re.Match[str]) -> str:
        name = match.group(1)
        if name not in mapping:
            raise ValueError(f"公式范围引用了不存在的字段: {name}")
        letter = col_letter(mapping[name])
        return f"${letter}${row}:${letter}${last}"

    def source(match: re.Match[str]) -> str:
        name = match.group(1)
        if name not in sources:
            raise ValueError(f"公式引用了不可用的数据源: {name}")
        return sources[name]

    formula = re.sub(r"\{this:([^{}]+)\}", current_cell, template)
    formula = re.sub(r"\{range:([^{}]+)\}", current_range, formula)
    formula = re.sub(r"\{source:([^{}]+)\}", source, formula)
    if re.search(r"\{(?:this|range|source):", formula):
        raise ValueError(f"公式包含未解析占位符: {formula}")
    return formula


def reserve_rule_columns(ws: Any, mapping: dict[str, int], rule: dict[str, Any], header_row: int) -> None:
    formulas = rule.get("lookups", []) + rule.get("calculations", [])
    by_name = {item["column"]: item for item in formulas}
    ordered = rule.get("output_columns") or [item["column"] for item in formulas]
    next_col = max(mapping.values(), default=0) + 1
    for name in ordered:
        if name not in by_name:
            raise ValueError(f"output_columns 包含未知字段: {name}")
        if name not in mapping:
            ws.Cells(header_row, next_col).Value = name
            mapping[name] = next_col
            next_col += 1


def write_formulas(ws: Any, mapping: dict[str, int], rules: list[dict[str, Any]], first: int, last: int, sources: dict[str, str]) -> None:
    for rule in rules:
        if last < first:
            continue
        name = rule["column"]
        col = mapping[name]
        formula = render_formula(rule["formula"], mapping, first, last, sources)
        cell = ws.Cells(first, col)
        cell.Formula = formula
        if last > first:
            ws.Range(cell, ws.Cells(last, col)).FillDown()
        if rule.get("number_format"):
            set_number_format(ws, col, first, last, rule["number_format"], name)


def actual_sheet_name(runner: ExcelRunner, path: Path, requested: str | None, default: str, password: str | None = None) -> str:
    book = runner.open(path, read_only=True, password=password)
    try:
        return sheet_for(book, requested, default).Name
    finally:
        runner.close(book, False)


def make_sources(
    raw_dir: Path,
    normalize_dir: Path,
    sources: dict[str, Any],
    shop: dict[str, Any] | None,
    default_sheet: str,
    exclude: set[str] | None = None,
) -> dict[str, str]:
    refs: dict[str, str] = {}
    excluded = exclude or set()
    for name, item in sources.items():
        if name == "shops" or name in excluded:
            continue
        root = normalize_dir if item.get("output") else raw_dir
        filename = item.get("output") or item["file"]
        path = source_path(root, filename)
        refs[name] = external_ref(path, item.get("sheet") or default_sheet, item.get("reference_range", "$A:$XFD"))
    if shop and shop.get("sales"):
        sales = source_path(raw_dir, shop, "sales")
        reference = external_ref(sales, shop.get("sales_sheet") or default_sheet, shop.get("sales_reference_range", "$B:$T"))
        refs["sales"] = reference
        if shop.get("sales_table"):
            refs[shop["sales_table"]] = reference
    return refs


def process_file(runner: ExcelRunner, input_path: Path, output_path: Path, source: dict[str, Any], rule: dict[str, Any], defaults: dict[str, Any], formula_sources: dict[str, str]) -> None:
    if output_path.exists() and not bool(defaults.get("overwrite", True)):
        raise FileExistsError(f"输出已存在且 overwrite=false: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = output_path.with_name(f".{output_path.stem}.normalize-tmp.xlsx")
    if temporary.exists():
        temporary.unlink()
    book = runner.open(input_path, read_only=False, password=source.get("password"))
    closed = False
    try:
        ws = sheet_for(book, source.get("sheet"), defaults["default_sheet"])
        header_row = int(defaults.get("header_row", 1))
        _, mapping = headers(ws, header_row)
        first = header_row + 1
        last = used_last_row(ws, header_row)
        format_columns(ws, mapping, rule, first, last)
        reserve_rule_columns(ws, mapping, rule, header_row)
        write_formulas(ws, mapping, rule.get("lookups", []), first, last, formula_sources)
        write_formulas(ws, mapping, rule.get("calculations", []), first, last, formula_sources)
        try:
            runner.excel.CalculateFullRebuild()
        except Exception as exc:
            print(f"警告: Excel 全量重算失败: {exc}")
        book.SaveAs(str(temporary.resolve()), FileFormat=XL_OPEN_XML_WORKBOOK)
        runner.close(book, False)
        book = None
        ws = None
        gc.collect()
        closed = True
        for attempt in range(10):
            try:
                os.replace(temporary, output_path)
                break
            except PermissionError as exc:
                if attempt == 9:
                    raise PermissionError(f"无法替换输出文件，请关闭正在打开的 Excel 文件: {output_path}") from exc
                gc.collect()
                time.sleep(0.5)
        print(f"完成: {input_path.name} -> {output_path.name} ({max(0, last - header_row)} 行)")
    except Exception:
        if not closed:
            runner.close(book, False)
        if temporary.exists():
            temporary.unlink()
        raise


def resolve_source_sheets(runner: ExcelRunner, raw_dir: Path, sources: dict[str, Any], default_sheet: str) -> dict[str, Any]:
    resolved = dict(sources)
    for name, item in sources.items():
        if name == "shops":
            continue
        item = dict(item)
        item["sheet"] = actual_sheet_name(runner, source_path(raw_dir, item), item.get("sheet"), default_sheet, item.get("password"))
        resolved[name] = item
    return resolved


def validate_rule_headers(rule_name: str, rule: dict[str, Any], mapping: dict[str, int]) -> None:
    available = set(mapping) | set(rule_fields(rule))
    require_columns(mapping, rule.get("text_columns", []) + rule.get("number_columns", []), rule_name)
    for item in rule.get("lookups", []) + rule.get("calculations", []):
        fields = re.findall(r"\{(?:this|range):([^{}]+)\}", item["formula"])
        missing = sorted(set(fields) - available)
        if missing:
            raise ValueError(f"规则 {rule_name}.{item['column']} 引用不存在的字段: {', '.join(missing)}")


def check_workbook(runner: ExcelRunner, path: Path, item: dict[str, Any], rule_name: str | None, rule: dict[str, Any] | None, defaults: dict[str, Any]) -> str:
    book = runner.open(path, read_only=True, password=item.get("password"))
    try:
        ws = sheet_for(book, item.get("sheet"), defaults["default_sheet"])
        _, mapping = headers(ws, int(defaults.get("header_row", 1)))
        if rule_name and rule:
            validate_rule_headers(rule_name, rule, mapping)
        return ws.Name
    finally:
        runner.close(book, False)


def check_inputs(runner: ExcelRunner, raw_dir: Path, sources: dict[str, Any], rules: dict[str, Any], defaults: dict[str, Any]) -> None:
    for name, item in sources.items():
        if name == "shops":
            continue
        path = source_path(raw_dir, item)
        if not path.exists():
            raise FileNotFoundError(path)
        rule = rules.get(name)
        sheet = check_workbook(runner, path, item, name if rule else None, rule, defaults)
        print(f"检查: {path.name} [{sheet}]")
    for shop in sources.get("shops", []):
        if not shop.get("enabled", True):
            output = shop.get("output") or f"{Path(shop['product']).stem}.xlsx"
            print(f"跳过: {shop['name']} 已停用；已有输出 {output} 不会自动删除")
            continue
        product = source_path(raw_dir, shop, "product")
        sales = source_path(raw_dir, shop, "sales") if shop.get("sales") else None
        for path in (product, sales):
            if path is None:
                continue
            if not path.exists():
                raise FileNotFoundError(path)
        selected_rule = shop.get("product_table", "shop_product")
        product_item = {"sheet": shop.get("product_sheet"), "password": shop.get("product_password")}
        product_sheet = check_workbook(runner, product, product_item, selected_rule, rules[selected_rule], defaults)
        sales_label = ""
        if sales:
            sales_item = {"sheet": shop.get("sales_sheet"), "password": shop.get("sales_password")}
            sales_sheet = check_workbook(runner, sales, sales_item, None, None, defaults)
            sales_label = f" / 销售 [{sales_sheet}]"
        print(f"检查: {shop['name']} 商品 [{product_sheet}]{sales_label}")


def _compose_config(specialized: dict[str, Any], common: dict[str, Any]) -> dict[str, Any]:
    validate_common_config(common)
    rules_config = grouped_actions(specialized)
    table_names = {item["table"] for item in common["tables"]}
    source_names = table_names | {"sales"}
    for action in rules_config:
        if action["table"] not in table_names and action["table"] not in TABLE_TYPES:
            raise ValueError(f"规则引用了不存在的 table: {action['table']}")
        if action["type"] == "function":
            aliases = re.findall(r"\{source:([^{}]+)\}", action["value"])
            unknown = sorted(set(aliases) - source_names)
            if unknown:
                raise ValueError(f"规则 {action['table']}.{action['column']} 使用未知数据源: {', '.join(unknown)}")
    sources: dict[str, Any] = {}
    shops: dict[str, dict[str, Any]] = {}
    for table in common["tables"]:
        alias = table["table"]
        item = dict(table)
        kind = table["type"]
        if kind == "erp":
            sources[alias] = item
        elif kind == "shop_product":
            shop = shops.setdefault(table["shop"], {"name": table["shop"]})
            shop.update({"product": table["file"], "product_table": alias})
            if "enabled" in item:
                shop["enabled"] = item["enabled"]
        elif kind == "shop_sales":
            shop = shops.setdefault(table["shop"], {"name": table["shop"]})
            shop.update({"sales": table["file"], "sales_table": alias})
            if "enabled" in item:
                shop["enabled"] = item["enabled"]
    sources["shops"] = list(shops.values())
    rules: dict[str, dict[str, Any]] = {}
    table_targets = {
        kind: [item["table"] for item in common["tables"] if item["type"] == kind]
        for kind in TABLE_TYPES
    }
    for action in rules_config:
        targets = table_targets.get(action["table"], [action["table"]])
        for target in targets:
            rule = rules.setdefault(target, {"text_columns": [], "number_columns": [], "lookups": [], "calculations": []})
            if action["type"] == "format":
                rule[f"{action['value']}_columns"].append(action["column"])
            else:
                formula = {"column": action["column"], "formula": action["value"]}
                rule["lookups"].append(formula)
    composed = dict(specialized)
    composed["paths"] = common.get("paths", {})
    composed["sheet"] = common.get("sheet", {"default": "Sheet1"})
    composed["sources"] = sources
    composed["rules"] = rules
    for name, item in sources.items():
        if name != "shops" and name in rules:
            item["output"] = f"{Path(item['file']).stem}.xlsx"
    return composed


def load_config(path: Path) -> dict[str, Any]:
    config = read_json(path)
    if "tables" in config:
        validate_common_config(config)
        return config
    common_path = path.parent.resolve() / "config.json"
    common = read_json(common_path)
    config = _compose_config(config, common)
    return config


def context(config_path: Path) -> tuple[dict[str, Any], Path, Path, dict[str, Any]]:
    config = load_config(config_path)
    base = config_path.parent.resolve()
    paths = config.get("paths", {})
    raw_value = paths.get("raw", "../data/raw")
    normalize_value = paths.get("normalize", "../data/normalize")
    raw_dir = source_path(base, raw_value)
    normalize_dir = source_path(base, normalize_value)
    defaults = {"default_sheet": config.get("sheet", {}).get("default", "Sheet1"), "header_row": 1, "overwrite": True}
    return config, raw_dir.resolve(), normalize_dir.resolve(), defaults


def check(config_path: Path) -> None:
    config, raw_dir, _, defaults = context(config_path)
    runner = ExcelRunner()
    try:
        check_inputs(runner, raw_dir, config["sources"], config["rules"], defaults)
    finally:
        runner.shutdown()
    print("配置和输入检查通过")


def run(config_path: Path) -> None:
    config, raw_dir, normalize_dir, defaults = context(config_path)
    rules = config["rules"]
    runner = ExcelRunner()
    try:
        check_inputs(runner, raw_dir, config["sources"], rules, defaults)
        sources = resolve_source_sheets(runner, raw_dir, config["sources"], defaults["default_sheet"])
        for name, item in sources.items():
            if name == "shops":
                continue
            output = item.get("output", f"{Path(item['file']).stem}.xlsx")
            refs = make_sources(
                raw_dir,
                normalize_dir,
                sources,
                None,
                defaults["default_sheet"],
                {name},
            )
            if name not in rules:
                continue
            process_file(runner, source_path(raw_dir, item), source_path(normalize_dir, output), item, rules[name], defaults, refs)
        for configured_shop in sources.get("shops", []):
            if not configured_shop.get("enabled", True):
                continue
            shop = dict(configured_shop)
            product = source_path(raw_dir, shop, "product")
            sales = source_path(raw_dir, shop, "sales") if shop.get("sales") else None
            shop["product_sheet"] = actual_sheet_name(runner, product, shop.get("product_sheet"), defaults["default_sheet"], shop.get("product_password"))
            if sales:
                shop["sales_sheet"] = actual_sheet_name(runner, sales, shop.get("sales_sheet"), defaults["default_sheet"], shop.get("sales_password"))
            selected_rule = shop.get("product_table", "shop_product")
            output = shop.get("output") or f"{product.stem}.xlsx"
            source = {"sheet": shop["product_sheet"], "password": shop.get("product_password")}
            process_file(runner, product, source_path(normalize_dir, output), source, rules[selected_rule], defaults, make_sources(raw_dir, normalize_dir, sources, shop, defaults["default_sheet"]))
    finally:
        runner.shutdown()
