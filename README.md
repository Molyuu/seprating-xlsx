# Excel 拆分工具 (Excel Splitter)

Splits consolidated Excel files into one file per `文件名称`, with one sheet per `工作表名称`.

Developer: Molyuu <i@molyuu.com> · Version: v0.1.1

## Run on the Mac (development)

Dependencies, the Python version, and the build tool are managed by [uv]
(`pyproject.toml`, `uv.lock`, `.python-version`). Install uv once, then:

    uv sync                              # create .venv from the lockfile (Python 3.11)
    uv run python app.py                 # GUI
    uv run python split_cli.py example         # CLI output: <file name><YYYYMMDD_HHMMSS>
    uv run python split_cli.py example -o out  # custom output folder (add --meta --stamp-filename --stamp-row)
    uv run python split_cli.py --version      # print the version
    uv run python -m unittest discover -s tests

`uv run` keeps the environment in sync automatically, so `uv sync` after pulling
is only needed to refresh `.venv`. To add or change a dependency, edit
`[project].dependencies` / `[dependency-groups].dev` in `pyproject.toml` and run
`uv sync` (or use `uv add openpyxl` / `uv add --dev pyinstaller`).

[uv]: https://docs.astral.sh/uv/

## Build the Windows .exe

PyInstaller cannot cross-compile, so build on Windows:

- **GitHub (recommended):** push this folder to a GitHub repo. The workflow
  `.github/workflows/build-windows.yml` runs the tests and uploads `ExcelSplitter.exe`
  as a build artifact (Actions tab -> latest run -> Artifacts). Publishing a GitHub
  release also builds the exe and attaches it as `ExcelSplitter-<tag>.exe`
  (for example, `ExcelSplitter-v0.1.1.exe`).
- **Any Windows PC/VM:** install uv, then run
  `uv sync` and
  `uv run pyinstaller --onefile --windowed --name ExcelSplitter --collect-data tzdata app.py`.

The result, `dist/ExcelSplitter.exe`, is a single file that needs no Python or other
installs. End users double-click it.

## Behaviour

- Header row auto-detected; blank rows split multiple tables; merged cells filled.
- Rows with empty key cells (e.g. `利润总计`) stay with the group above.
- Fallback keys: a `…公司` column for the file name and the source sheet name for the sheet.
- Metadata columns, when enabled, hold the group values (e.g. `小明`, `A`).
- Default output folder: the first source file name plus the local system timestamp
  (e.g. `会计公司汇总20261019_195603`), created beside the source file.
- Timestamp: filename suffix `_YYYYMMDD` and/or a first row `生成时间: date, timezone`.
