# Repository Guidelines

## Project Structure & Module Organization

This repository contains a small SDGanda data export workflow.

- `fetch_sdganda_units.py` is the main Python script. It fetches unit metadata, fetches win-rate data, combines both sources, and writes JSON/XLSX outputs.
- `README.md` documents usage, output fields, data sources, and rank display rules.
- `YYYYMMDD/` folders, such as `20260620/`, contain generated outputs: `units.json`, `win_rates.json`, `combined_units_win_rate.json`, and `sdganda_units_YYYYMMDD.xlsx`.

Keep new source files at the repository root unless the project grows enough to justify a package directory. Keep generated exports in dated `YYYYMMDD/` folders.

## Build, Test, and Development Commands

There is no build step. Use Python directly from PowerShell:

```powershell
python .\fetch_sdganda_units.py --sample 3
```

Runs a small smoke test by fetching only the first three units.

```powershell
python .\fetch_sdganda_units.py --date 20260620 --days 14
```

Writes outputs to `20260620/` using a 14-day win-rate window.

The script depends on `openpyxl` for Excel output. If it is missing, install it in your active Python environment before running the export.

## Coding Style & Naming Conventions

Use Python 3 with type hints, matching the existing script style. Prefer small functions with direct names such as `fetch_units`, `fetch_win_rate`, and `write_excel`. Use `snake_case` for variables and functions, `UPPER_CASE` for constants, and `Path` objects for filesystem paths.

Avoid broad refactors. Changes should be traceable to the requested data export behavior or documentation update.

## Testing Guidelines

No formal test suite exists yet. Before changing fetch, combine, or Excel logic, run:

```powershell
python .\fetch_sdganda_units.py --sample 3 --date 20990101
```

Verify that all four expected files are created and that the Excel sheet opens with the `units_win_rate` worksheet, frozen header row, filters, and percentage formatting for `win_rate` and `lose_rate`.

## Commit & Pull Request Guidelines

The current history uses concise imperative commit messages, for example `Add SDGanda unit data export script and README documentation`. Follow that style: start with a verb and describe the changed behavior.

Pull requests should include a short purpose summary, commands run, generated output folders touched, and any API or field-shape assumptions. Include screenshots only when Excel formatting or rendered output layout changes.

## Security & Configuration Tips

Do not commit credentials or private API tokens. This script currently uses public SDGanda API endpoints and a configurable `--api-base-url`; document any endpoint changes in `README.md`.
