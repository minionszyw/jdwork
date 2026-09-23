"""Shared Excel COM and worksheet helpers."""

from __future__ import annotations

from pathlib import Path
from typing import Any


def clean_text(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).replace("\r", "").replace("\n", "").strip()

def as_tuple(value: Any) -> tuple:
    return value if isinstance(value, tuple) else (value,)


class ExcelRunner:
    def __init__(self) -> None:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as exc:
            raise RuntimeError("缺少 pywin32，请运行: py -m pip install jdwork") from exc
        self.pythoncom = pythoncom
        pythoncom.CoInitialize()
        try:
            self.excel = win32.DispatchEx("Excel.Application")
        except Exception:
            pythoncom.CoUninitialize()
            raise
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False
        self.excel.EnableEvents = False
        self.excel.AskToUpdateLinks = False
        self.open_books: list[Any] = []

    def open(self, path: Path, read_only: bool = True, password: str | None = None):
        kwargs: dict[str, Any] = {"ReadOnly": read_only, "UpdateLinks": 0, "IgnoreReadOnlyRecommended": True, "AddToMru": False, "Local": True}
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
            self.pythoncom.CoUninitialize()


def sheet_for(book: Any, requested: str | None, default: str) -> Any:
    names = [book.Worksheets(i).Name for i in range(1, book.Worksheets.Count + 1)]
    for name in [item for item in (requested, default) if item]:
        if name in names:
            return book.Worksheets(name)
    for i in range(1, book.Worksheets.Count + 1):
        ws = book.Worksheets(i)
        used = ws.UsedRange
        if used.Rows.Count > 1 and used.Columns.Count > 1:
            return ws
    return book.Worksheets(1)


def headers(ws: Any, header_row: int) -> tuple[list[str], dict[str, int]]:
    used = ws.UsedRange
    first = used.Column
    last = first + used.Columns.Count - 1
    values = ws.Range(ws.Cells(header_row, first), ws.Cells(header_row, last)).Value
    raw = as_tuple(values)
    if len(raw) == 1 and isinstance(raw[0], tuple):
        raw = raw[0]
    result = [clean_text(value) or "" for value in raw]
    mapping: dict[str, int] = {}
    for offset, name in enumerate(result):
        if not name:
            continue
        if name in mapping:
            raise ValueError(f"工作表 {ws.Name!r} 存在重复表头: {name}")
        mapping[name] = first + offset
    return result, mapping


def require_columns(mapping: dict[str, int], columns: list[str], label: str) -> None:
    missing = [name for name in columns if name not in mapping]
    if missing:
        raise ValueError(f"{label} 缺少字段: {', '.join(missing)}")


def used_last_row(ws: Any, header_row: int) -> int:
    used = ws.UsedRange
    return max(header_row, used.Row + used.Rows.Count - 1)

