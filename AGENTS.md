# Repository Guidelines

## Project Structure

- `norm.py` is the Windows Excel COM normalization script.
- `norm.json` contains input paths, workbook sheets, field-format rules, lookup formulas, and calculated columns.
- `requirements.txt` lists the Python dependency (`pywin32`).
- `README.md` documents installation, configuration, and operation.
- `tests/` contains Windows-independent unit tests for configuration and formula helpers.
- `raw/` contains exported source workbooks and `norm/` contains generated workbooks. Both are excluded from Git.
- `docs/` is reserved for local documentation and is excluded from Git.

## Build, Test, and Development Commands

Install dependencies:

```powershell
py -m pip install -r .\requirements.txt
```

Run the normalizer from the repository root:

```powershell
py .\norm.py --config .\norm.json
```

The script requires Windows and an installed Microsoft Excel application. It reads `.xlsx` and `.csv` files through Excel COM and writes generated `.xlsx` files to `norm/`.

Validate Python syntax without opening Excel:

```powershell
python -m py_compile norm.py
```

Validate configuration and input workbooks without generating outputs:

```powershell
python .\norm.py --check
```

Run unit tests:

```powershell
python -m unittest discover -s tests -v
```

For behavior changes, run the unit tests and then run the normalizer against representative files in `raw/`; inspect formulas and calculated values in the generated workbooks.

## Coding Style and Naming

- Use Python 3.10+ syntax, four-space indentation, and clear type annotations where practical.
- Prefer small functions with explicit inputs and meaningful error messages.
- Keep user-editable behavior in `norm.json`; avoid hard-coding paths, sheet names, or business formulas in Python.
- Add or modify a shop through `sources.shops`; select a per-shop rule with `rule` instead of branching on the shop name in Python.
- Preserve Chinese source field names exactly when they are used as configuration keys.
- Use `snake_case` for Python functions and variables; use descriptive JSON keys.

## Testing Guidelines

When changing normalization behavior, verify text fields preserve leading zeros, numeric fields are numeric, formulas remain formulas, and all configured output files are produced. Also test `--check`, safe overwrite behavior, reruns without duplicate columns, and Excel process cleanup after success or failure.

## Commit and Pull Request Guidelines

Use concise imperative commit subjects, for example `Add Excel table normalization script`. Keep commits focused on one logical change. Pull requests should describe the affected rules or configuration, list the validation command and result, and mention any Excel or Windows prerequisites. Do not commit files from `raw/`, `norm/`, `docs/`, credentials, or exported customer data.

## Security and Configuration

Treat workbook contents and any configured passwords as sensitive. Keep credentials out of committed JSON when possible. Do not modify files under `raw/`; generated output belongs under `norm/` and should be reviewed locally before sharing.
