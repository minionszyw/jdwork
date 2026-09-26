# Repository Guidelines

## Project Structure

- `pyproject.toml` defines the installable package and `jdw` console entry point.
- `config/` contains user-editable `config.json`, `normalize.json`, `filter.json`, and `backfill.json`.
- `src/jdwork/` contains the CLI, shared Excel/config helpers, normalization, filtering, and backfill modules.
- `tests/` contains Windows-independent unit tests. Excel COM integration checks use local workbooks.
- `data/raw/`, `data/normalize/`, `data/filter/`, and `docs/` are local data/output directories excluded from Git.

## Build, Test, and Development Commands

Install locally:

```powershell
py -m pip install .
```

Use editable installation while developing:

```powershell
py -m pip install -e .
```

Run the workflow from the repository root:

```powershell
jdw normalize
jdw filter
jdw backfill --input .\data\filter\filter-{batch_id}.xlsx --dry-run
```

Validation commands:

```powershell
jdw normalize --check
jdw filter --check
python -m unittest discover -s tests -v
```

The commands require Windows and Microsoft Excel. `normalize` and `filter` use Excel COM; `backfill --dry-run` never saves raw files.

## Coding Style and Naming

- Use Python 3.10+, four-space indentation, type annotations where practical, and `snake_case` names.
- Keep shared behavior in `src/jdwork/excel.py` and `src/jdwork/config.py`; avoid copying COM or path logic between commands.
- Prefer reuse (Don't Repeat Yourself). Apply the Boy Scout Rule: leave touched code clearer, remove stale imports and update nearby docs/tests.
- Keep business rules, formulas, paths, sheets, and writable fields in JSON configuration.
- Preserve Chinese source field names exactly when they are configuration keys.
- Keep every JSON configuration file expanded with indentation; do not compress objects or arrays onto one line.
- Group `normalize.json` rules by table: each group contains `table` and `columns`; each column action contains `column`, `type`, and `value`. Preserve column order and avoid duplicate tables or columns within a group.
- Keep CLI behavior in `cli.py`; command modules expose reusable `run` and `check` functions.
- Commands use the fixed files in `config/`; shared paths and the default sheet belong in `config.json`, normalization rules in `normalize.json`, filters in `filter.json`, and backfill settings in `backfill.json`.

## Configuration-Driven Architecture

- Keep Python modules generic and stable. New normalization metrics, formulas, filter rules, and writable fields belong in JSON configuration; do not add a Python module for one business rule.
- Reuse the shared formula renderer, filter operators, Excel helpers, and backfill safeguards before adding code.
- When a backfill field is editable in the filter workbook, choose a separate stable `verify_fields` value so the edited value is not used as its own verification snapshot.

## Testing Guidelines

Test configuration validation, text leading-zero preservation, numeric conversion, formula rendering, filter operators, header-row handling, duplicate-key conflicts, and backfill allowlists. For behavior changes, run unit tests plus `normalize --check` and `filter --check` against representative local files. Validate real backfill writes only against a temporary raw copy.

## Commit and Pull Request Guidelines

Use concise imperative commit subjects, such as `Add jdwork CLI package`. Keep each commit focused. Pull requests should describe changed commands/configuration, list test and Excel validation commands, and mention Windows/Excel prerequisites. Never commit raw exports, generated workbooks, credentials, or customer data.

## Security and Configuration

Treat workbook contents and configured passwords as sensitive. Keep credentials out of committed JSON. Only fields listed in `config/backfill.json` `fields` may be written back to raw files. `店铺`, `类型`, and `SKUID` are metadata/identity fields and cannot be written. Review the dry-run output before applying a batch.
