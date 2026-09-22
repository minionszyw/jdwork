from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as win32


XL_OPEN_XML_WORKBOOK = 51
XL_CALC_AUTOMATIC = -4105


def clean_text(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).replace("\r", "").replace("\n", "").strip()


def clean_number(value: Any) -> Any:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value if not isinstance(value, float) or math.isfinite(value) else None
    text = clean_text(value)
    if not text or text.lower() in {"null", "none", "--", "-"}:
        return None
    text = text.replace(",", "")
    try:
        number = float(text)
        return int(number) if number.is_integer() else number
    except ValueError:
        return value


def as_tuple(value: Any) -> tuple:
    if isinstance(value, tuple):
        return value
    return (value,)


class ExcelRunner:
    def __init__(self) -> None:
        pythoncom.CoInitialize()
        self.excel = win32.DispatchEx("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False
        self.excel.EnableEvents = False
        self.excel.AskToUpdateLinks = False
        try:
            self.excel.Calculation = XL_CALC_AUTOMATIC
        except Exception:
            pass
        self.open_books: list[Any] = []

    def open(self, path: Path, read_only: bool = True, password: str | None = None):
        kwargs = {
            "ReadOnly": read_only,
            "UpdateLinks": 0,
            "IgnoreReadOnlyRecommended": True,
            "AddToMru": False,
            "Local": True,
        }
        if password:
            kwargs["Password"] = password
        book = self.excel.Workbooks.Open(str(path.resolve()), **kwargs)
        self.open_books.append(book)
        return book

    def close(self, book: Any, save: bool = False) -> None:
        try:
            book.Close(SaveChanges=save)
        finally:
            if book in self.open_books:
                self.open_books.remove(book)

    def shutdown(self) -> None:
        for book in reversed(self.open_books):
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            self.excel.Quit()
        finally:
            pythoncom.CoUninitialize()


def sheet_for(book: Any, requested: str | None, default: str) -> Any:
    names = [book.Worksheets(i).Name for i in range(1, book.Worksheets.Count + 1)]
    candidates = [x for x in (requested, default) if x]
    for name in candidates:
        if name in names:
            return book.Worksheets(name)
    for i in range(1, book.Worksheets.Count + 1):
        ws = book.Worksheets(i)
        ur = ws.UsedRange
        if ur.Rows.Count > 1 and ur.Columns.Count > 1:
            return ws
    return book.Worksheets(1)


def headers(ws: Any, header_row: int) -> tuple[list[str], dict[str, int]]:
    ur = ws.UsedRange
    first = ur.Column
    last = first + ur.Columns.Count - 1
    values = ws.Range(ws.Cells(header_row, first), ws.Cells(header_row, last)).Value
    raw = as_tuple(values)
    if len(raw) == 1 and isinstance(raw[0], tuple):
        raw = raw[0]
    result = [clean_text(x) or "" for x in raw]
    mapping: dict[str, int] = {}
    for offset, name in enumerate(result):
        if not name:
            continue
        if name in mapping:
            raise ValueError(f"工作表 {ws.Name!r} 存在重复表头: {name}")
        mapping[name] = first + offset
    return result, mapping


def require_columns(mapping: dict[str, int], columns: list[str], label: str) -> None:
    missing = [x for x in columns if x not in mapping]
    if missing:
        raise ValueError(f"{label} 缺少字段: {', '.join(missing)}")


def used_last_row(ws: Any, header_row: int) -> int:
    ur = ws.UsedRange
    return max(header_row, ur.Row + ur.Rows.Count - 1)


def set_column_values(ws: Any, col: int, first_row: int, last_row: int, fn) -> None:
    if last_row < first_row:
        return
    values = ws.Range(ws.Cells(first_row, col), ws.Cells(last_row, col)).Value
    rows = as_tuple(values)
    out = []
    for row in rows:
        value = row[0] if isinstance(row, tuple) else row
        out.append((fn(value),))
    ws.Range(ws.Cells(first_row, col), ws.Cells(last_row, col)).Value = tuple(out)


def format_columns(ws: Any, mapping: dict[str, int], text_cols: list[str], number_cols: list[str], first: int, last: int) -> None:
    require_columns(mapping, text_cols + number_cols, ws.Name)
    for name in text_cols:
        col = mapping[name]
        try:
            ws.Range(ws.Cells(first, col), ws.Cells(last, col)).NumberFormat = "@"
        except Exception:
            pass
        set_column_values(ws, col, first, last, clean_text)
    for name in number_cols:
        col = mapping[name]
        set_column_values(ws, col, first, last, clean_number)
        try:
            ws.Range(ws.Cells(first, col), ws.Cells(last, col)).NumberFormat = "General"
        except Exception:
            pass


def col_letter(col: int) -> str:
    result = ""
    while col:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def external_ref(path: Path, sheet: str, cell_range: str) -> str:
    text = f"{path.resolve().parent}\\[{path.name}]{sheet}".replace("'", "''")
    return f"'{text}'!{cell_range}"


def render_formula(template: str, ws: Any, mapping: dict[str, int], row: int, last: int, sources: dict[str, str]) -> str:
    def this(match):
        name = match.group(1)
        if name not in mapping:
            raise ValueError(f"公式引用了不存在的字段: {name}")
        return f"{col_letter(mapping[name])}{row}"

    def rng(match):
        name = match.group(1)
        if name not in mapping:
            raise ValueError(f"公式范围引用了不存在的字段: {name}")
        letter = col_letter(mapping[name])
        return f"${letter}${row}:${letter}${last}"

    formula = re.sub(r"\{this:([^{}]+)\}", this, template)
    formula = re.sub(r"\{range:([^{}]+)\}", rng, formula)
    formula = re.sub(r"\{source:([^{}]+)\}", lambda m: sources[m.group(1)], formula)
    return formula


def append_rule_columns(ws: Any, mapping: dict[str, int], rules: list[dict[str, Any]], header_row: int, first: int, last: int, sources: dict[str, str]) -> None:
    next_col = max(mapping.values(), default=0) + 1
    for rule in rules:
        name = rule["column"]
        if name in mapping:
            col = mapping[name]
        else:
            col = next_col
            next_col += 1
            ws.Cells(header_row, col).Value = name
            mapping[name] = col
    for rule in rules:
        name = rule["column"]
        col = mapping[name]
        if last < first or "formula" not in rule:
            continue
        formula = render_formula(rule["formula"], ws, mapping, first, last, sources)
        cell = ws.Cells(first, col)
        cell.Formula = formula
        if last > first:
            ws.Range(cell, ws.Cells(last, col)).FillDown()
        if rule.get("number_format"):
            try:
                ws.Range(ws.Cells(first, col), ws.Cells(last, col)).NumberFormat = rule["number_format"]
            except Exception:
                pass


def reserve_columns(ws: Any, mapping: dict[str, int], rules: list[dict[str, Any]], header_row: int) -> None:
    next_col = max(mapping.values(), default=0) + 1
    for rule in rules:
        name = rule["column"]
        if name not in mapping:
            ws.Cells(header_row, next_col).Value = name
            mapping[name] = next_col
            next_col += 1


def save_as_xlsx(book: Any, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()
    book.SaveAs(str(target.resolve()), FileFormat=XL_OPEN_XML_WORKBOOK)


def source_path(base: Path, section: dict[str, Any], key: str, root: Path | None = None) -> Path:
    item = section[key]
    path = Path(item["file"] if isinstance(item, dict) else item)
    return path if path.is_absolute() else (root or base) / path


def prepare_book(runner: ExcelRunner, path: Path, sheet_name: str | None, default_sheet: str, password: str | None = None):
    book = runner.open(path, read_only=False, password=password)
    ws = sheet_for(book, sheet_name, default_sheet)
    return book, ws


def actual_sheet_name(runner: ExcelRunner, path: Path, requested: str | None, default: str, password: str | None = None) -> str:
    book = runner.open(path, read_only=True, password=password)
    try:
        return sheet_for(book, requested, default).Name
    finally:
        runner.close(book, False)


def process_file(runner: ExcelRunner, input_path: Path, output_path: Path, source_cfg: dict[str, Any], rules: dict[str, Any], defaults: dict[str, Any], formula_sources: dict[str, str]) -> None:
    book, ws = prepare_book(runner, input_path, source_cfg.get("sheet"), defaults["default_sheet"], source_cfg.get("password"))
    try:
        header_row = int(defaults.get("header_row", 1))
        _, mapping = headers(ws, header_row)
        first = header_row + 1
        last = used_last_row(ws, header_row)
        format_columns(ws, mapping, rules.get("text_columns", []), rules.get("number_columns", []), first, last)
        ordered = rules.get("output_columns")
        if ordered:
            all_rules = {rule["column"]: rule for rule in rules.get("lookups", []) + rules.get("calculations", [])}
            if set(ordered) != set(all_rules):
                raise ValueError(f"{ws.Name} output_columns 与公式字段不一致")
            reserve_columns(ws, mapping, [all_rules[name] for name in ordered], header_row)
        append_rule_columns(ws, mapping, rules.get("lookups", []), header_row, first, last, formula_sources)
        append_rule_columns(ws, mapping, rules.get("calculations", []), header_row, first, last, formula_sources)
        runner.excel.CalculateFullRebuild()
        save_as_xlsx(book, output_path)
        print(f"完成: {input_path.name} -> {output_path.name} ({max(0, last - header_row)} 行)")
    finally:
        runner.close(book, False)


def make_sources(base: Path, raw_dir: Path, norm_dir: Path, source_cfg: dict[str, Any], shops: list[dict[str, Any]], current_shop: dict[str, Any] | None, default_sheet: str) -> dict[str, str]:
    refs: dict[str, str] = {}
    definitions = {
        "erp_inventory": (norm_dir / "ERP库存.xlsx", source_cfg["erp_inventory"].get("sheet") or "ERP库存", "$A:$J"),
        "erp_product": (norm_dir / "ERP商品.xlsx", source_cfg["erp_product"].get("sheet") or default_sheet, "$A:$M"),
        "erp_combo": (norm_dir / "ERP组合.xlsx", source_cfg["erp_combo"].get("sheet") or "ERP组合", "$A:$U"),
        "erp_ban": (source_path(base, source_cfg, "erp_ban", raw_dir), source_cfg["erp_ban"].get("sheet") or default_sheet, "$A:$H"),
    }
    if current_shop:
        sales = source_path(base, current_shop, "sales", raw_dir)
        definitions["sales"] = (sales, current_shop.get("sales_sheet") or default_sheet, "$B:$T")
    for name, (path, sheet, cell_range) in definitions.items():
        refs[name] = external_ref(path, sheet, cell_range)
    return refs


def load_config(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8-sig") as fh:
        return json.load(fh)


def run(config_path: Path) -> None:
    config = load_config(config_path)
    base = config_path.parent.resolve()
    paths = config.get("paths", {})
    raw_dir = (base / paths.get("raw", "raw")).resolve()
    norm_dir = (base / paths.get("norm", "norm")).resolve()
    defaults = {"default_sheet": "Sheet1", "header_row": 1, **config.get("excel", {})}
    sources = config["sources"]
    rules = config["rules"]
    runner = ExcelRunner()
    try:
        sources = dict(sources)
        inventory_input = source_path(base, sources, "erp_inventory", raw_dir)
        product_input = source_path(base, sources, "erp_product", raw_dir)
        combo_input = source_path(base, sources, "erp_combo", raw_dir)
        ban_input = source_path(base, sources, "erp_ban", raw_dir)
        for p in [inventory_input, product_input, combo_input, ban_input]:
            if not p.exists():
                raise FileNotFoundError(p)
        for key in ["erp_inventory", "erp_product", "erp_combo", "erp_ban"]:
            entry = dict(sources[key])
            entry["sheet"] = actual_sheet_name(runner, source_path(base, sources, key, raw_dir), entry.get("sheet"), defaults["default_sheet"], entry.get("password"))
            sources[key] = entry

        process_file(runner, inventory_input, norm_dir / "ERP库存.xlsx", sources["erp_inventory"], rules["erp_inventory"], defaults, {})
        process_file(runner, product_input, norm_dir / "ERP商品.xlsx", sources["erp_product"], rules["erp_product"], defaults, make_sources(base, raw_dir, norm_dir, sources, [], None, defaults["default_sheet"]))
        process_file(runner, combo_input, norm_dir / "ERP组合.xlsx", sources["erp_combo"], rules["erp_combo"], defaults, make_sources(base, raw_dir, norm_dir, sources, [], None, defaults["default_sheet"]))

        for shop in sources.get("shops", []):
            if not shop.get("enabled", True):
                continue
            product = source_path(base, shop, "product", raw_dir)
            sales = source_path(base, shop, "sales", raw_dir)
            if not product.exists() or not sales.exists():
                raise FileNotFoundError(f"店铺 {shop.get('name', product.stem)} 缺少商品或销售文件")
            shop = dict(shop)
            shop["product_sheet"] = actual_sheet_name(runner, product, shop.get("product_sheet"), defaults["default_sheet"], shop.get("product_password"))
            shop["sales_sheet"] = actual_sheet_name(runner, sales, shop.get("sales_sheet"), defaults["default_sheet"], shop.get("sales_password"))
            process_file(runner, product, norm_dir / f"{product.stem}.xlsx", {"file": shop["product"], "sheet": shop.get("product_sheet"), "password": shop.get("product_password")}, rules["shop_product"], defaults, make_sources(base, raw_dir, norm_dir, sources, [], shop, defaults["default_sheet"]))
    finally:
        runner.shutdown()


def main() -> int:
    parser = argparse.ArgumentParser(description="通过 Excel COM 将 raw 表格标准化到 norm")
    parser.add_argument("-c", "--config", default="norm.json")
    args = parser.parse_args()
    try:
        run(Path(args.config).resolve())
        return 0
    except Exception as exc:
        print(f"失败: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
