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
REQUIRED_ERP_SOURCES = ("erp_inventory", "erp_product", "erp_ban", "erp_combo")
REQUIRED_RULES = ("erp_inventory", "erp_product", "erp_combo")
FORMULA_SOURCE_NAMES = {*REQUIRED_ERP_SOURCES, "sales"}

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


def validate_config(config: dict[str, Any]) -> None:
    errors: list[str] = []
    paths = config.get("paths", {})
    if not isinstance(paths, dict):
        errors.append("paths 必须是对象")
    elif "norm" in paths:
        errors.append("paths.norm 已停用，请改用 paths.normalize")
    sources = config.get("sources")
    rules = config.get("rules")
    if not isinstance(sources, dict):
        raise ValueError("配置缺少对象 sources")
    if not isinstance(rules, dict):
        raise ValueError("配置缺少对象 rules")
    for name in REQUIRED_ERP_SOURCES:
        item = sources.get(name)
        if not isinstance(item, dict) or not item.get("file"):
            errors.append(f"sources.{name} 必须配置 file")
        if name in REQUIRED_RULES and name not in rules:
            errors.append(f"rules 缺少 {name}")
    seen_names: set[str] = set()
    seen_outputs: set[str] = set()
    shops = sources.get("shops", [])
    if not isinstance(shops, list):
        errors.append("sources.shops 必须是数组")
        shops = []
    for index, shop in enumerate(shops):
        prefix = f"sources.shops[{index}]"
        if not isinstance(shop, dict):
            errors.append(f"{prefix} 必须是对象")
            continue
        for key in ("name", "product", "sales"):
            if not shop.get(key):
                errors.append(f"{prefix} 缺少 {key}")
        name = str(shop.get("name", ""))
        if name in seen_names:
            errors.append(f"店铺名称重复: {name}")
        seen_names.add(name)
        selected_rule = shop.get("rule", "shop_product")
        if selected_rule not in rules:
            errors.append(f"店铺 {name!r} 引用了不存在的规则 {selected_rule!r}")
        output = shop.get("output") or f"{Path(str(shop.get('product', 'unknown'))).stem}.xlsx"
        if output in seen_outputs:
            errors.append(f"店铺输出文件重复: {output}")
        seen_outputs.add(output)
    for name, rule in rules.items():
        if not isinstance(rule, dict):
            errors.append(f"rules.{name} 必须是对象")
            continue
        for list_name in ("text_columns", "number_columns", "lookups", "calculations"):
            if list_name in rule and not isinstance(rule[list_name], list):
                errors.append(f"rules.{name}.{list_name} 必须是数组")
        fields = rule_fields(rule)
        duplicates = sorted({field for field in fields if fields.count(field) > 1})
        if duplicates:
            errors.append(f"rules.{name} 公式字段重复: {', '.join(duplicates)}")
        ordered = rule.get("output_columns")
        if ordered is not None:
            if len(ordered) != len(set(ordered)):
                errors.append(f"rules.{name}.output_columns 存在重复字段")
            if set(ordered) != set(fields):
                errors.append(f"rules.{name}.output_columns 必须与公式字段一致")
        for formula_rule in rule.get("lookups", []) + rule.get("calculations", []):
            if not isinstance(formula_rule, dict) or not formula_rule.get("column") or not formula_rule.get("formula"):
                errors.append(f"rules.{name} 中每条公式必须包含 column 和 formula")
                continue
            aliases = re.findall(r"\{source:([^{}]+)\}", formula_rule["formula"])
            unknown = sorted(set(aliases) - FORMULA_SOURCE_NAMES)
            if unknown:
                errors.append(f"rules.{name}.{formula_rule['column']} 使用未知数据源: {', '.join(unknown)}")
    if errors:
        raise ValueError("配置校验失败:\n- " + "\n- ".join(errors))



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
    for name in REQUIRED_ERP_SOURCES:
        if name in excluded:
            continue
        item = sources[name]
        root = normalize_dir if item.get("output") else raw_dir
        filename = item.get("output") or item["file"]
        path = source_path(root, filename)
        refs[name] = external_ref(path, item.get("sheet") or default_sheet, item.get("reference_range", "$A:$XFD"))
    if shop:
        sales = source_path(raw_dir, shop, "sales")
        refs["sales"] = external_ref(sales, shop.get("sales_sheet") or default_sheet, shop.get("sales_reference_range", "$A:$XFD"))
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
    for name in REQUIRED_ERP_SOURCES:
        item = dict(sources[name])
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
    for name in REQUIRED_ERP_SOURCES:
        item = sources[name]
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
        sales = source_path(raw_dir, shop, "sales")
        for path in (product, sales):
            if not path.exists():
                raise FileNotFoundError(path)
        selected_rule = shop.get("rule", "shop_product")
        product_item = {"sheet": shop.get("product_sheet"), "password": shop.get("product_password")}
        sales_item = {"sheet": shop.get("sales_sheet"), "password": shop.get("sales_password")}
        product_sheet = check_workbook(runner, product, product_item, selected_rule, rules[selected_rule], defaults)
        sales_sheet = check_workbook(runner, sales, sales_item, None, None, defaults)
        print(f"检查: {shop['name']} 商品 [{product_sheet}] / 销售 [{sales_sheet}]")


def load_config(path: Path) -> dict[str, Any]:
    config = read_json(path)
    validate_config(config)
    return config


def context(config_path: Path) -> tuple[dict[str, Any], Path, Path, dict[str, Any]]:
    config = load_config(config_path)
    base = config_path.parent.resolve()
    paths = config.get("paths", {})
    raw_value = paths.get("raw", "../data/raw")
    normalize_value = paths.get("normalize", "../data/normalize")
    raw_dir = source_path(base, raw_value)
    normalize_dir = source_path(base, normalize_value)
    defaults = {"default_sheet": "Sheet1", "header_row": 1, "overwrite": True, **config.get("excel", {})}
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
        fixed_outputs = {"erp_inventory": "ERP库存.xlsx", "erp_product": "ERP商品.xlsx", "erp_combo": "ERP组合.xlsx"}
        for name in ("erp_inventory", "erp_product", "erp_combo"):
            item = sources[name]
            output = item.get("output", fixed_outputs[name])
            refs = {} if name == "erp_inventory" else make_sources(
                raw_dir,
                normalize_dir,
                sources,
                None,
                defaults["default_sheet"],
                {name},
            )
            process_file(runner, source_path(raw_dir, item), source_path(normalize_dir, output), item, rules[name], defaults, refs)
        for configured_shop in sources.get("shops", []):
            if not configured_shop.get("enabled", True):
                continue
            shop = dict(configured_shop)
            product = source_path(raw_dir, shop, "product")
            sales = source_path(raw_dir, shop, "sales")
            shop["product_sheet"] = actual_sheet_name(runner, product, shop.get("product_sheet"), defaults["default_sheet"], shop.get("product_password"))
            shop["sales_sheet"] = actual_sheet_name(runner, sales, shop.get("sales_sheet"), defaults["default_sheet"], shop.get("sales_password"))
            selected_rule = shop.get("rule", "shop_product")
            output = shop.get("output") or f"{product.stem}.xlsx"
            source = {"sheet": shop["product_sheet"], "password": shop.get("product_password")}
            process_file(runner, product, source_path(normalize_dir, output), source, rules[selected_rule], defaults, make_sources(raw_dir, normalize_dir, sources, shop, defaults["default_sheet"]))
    finally:
        runner.shutdown()


