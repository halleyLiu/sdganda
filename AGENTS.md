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

On Windows, confirm that `python` is a real interpreter before relying on it. In this workspace, `python.exe` may resolve to the Microsoft Store alias and exit without useful output; if that happens, use the Codex bundled Python executable from the workspace dependency runtime instead.

```powershell
python .\fetch_sdganda_units.py --date 20260620 --days 14
```

Writes outputs to `20260620/` using a 14-day win-rate window.

The script depends on `openpyxl` for Excel output. If it is missing, install it in your active Python environment before running the export.

Running the full export needs network access to `https://dbapi.sdganda.com/api`. If the first run fails with a connection-refused or sandbox/network error, retry with approved network access rather than changing script logic.

## Win-Rate Analysis Notes

For rank-bucket statistics, group units by the numeric `Rank` field instead of parsing the displayed `品阶` string. Use `Rank=4` for S-level units, `Rank=3` for A-level units, `Rank=2` for B-level units, and `Rank=0` or `Rank=1` for C-level units. This covers suffix variants such as `SS`, `SR`, `SU`, `AS`, `AR`, `AU`, `BS`, `BR`, and `BU` without fragile string matching.

When comparing dated exports, match units by `ID`, filter on the requested sample threshold such as `win_times > 50`, rank each date with the same grouping and sorting rules, then label July entries as `新进入`, `上升 N 名`, `下降 N 名`, or `持平` relative to the June rank.

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

This is an open source project, so assume all committed content may become public. Do not commit credentials, private API tokens, personal account details, private endpoint URLs, cookies, local machine paths with sensitive usernames, or unpublished data. This script currently uses public SDGanda API endpoints and a configurable `--api-base-url`; document any endpoint changes in `README.md`.
