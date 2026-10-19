# Excel 拆分工具 (Excel Splitter)

Splits consolidated Excel files into one file per `文件名称`, with one sheet per `工作表名称`.

## Run on the Mac (development)

    pip3 install -r requirements.txt
    python3 app.py                       # GUI
    python3 split_cli.py example -o out  # CLI (add --meta --stamp-filename --stamp-row)
    python3 -m unittest discover -s tests

## Build the Windows .exe

PyInstaller cannot cross-compile, so build on Windows:

- **GitHub (recommended):** push this folder to a GitHub repo. The workflow
  `.github/workflows/build-windows.yml` runs the tests and uploads `ExcelSplitter.exe`
  as a build artifact (Actions tab -> latest run -> Artifacts).
- **Any Windows PC/VM:** install Python 3.11, then run
  `pip install -r requirements.txt pyinstaller` and
  `pyinstaller --onefile --windowed --name ExcelSplitter --collect-data tzdata app.py`.

The result, `dist/ExcelSplitter.exe`, is a single file that needs no Python or other
installs. End users double-click it.

## Behaviour

- Header row auto-detected; blank rows split multiple tables; merged cells filled.
- Rows with empty key cells (e.g. `利润总计`) stay with the group above.
- Fallback keys: a `…公司` column for the file name and the source sheet name for the sheet.
- Metadata columns, when enabled, hold the group values (e.g. `小明`, `A`).
- Timestamp: filename suffix `_YYYYMMDD` and/or a first row `生成时间: date, timezone`.
