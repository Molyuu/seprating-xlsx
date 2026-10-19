"""Command line: python split_cli.py <files-or-folders...> -o <out_dir>"""
import argparse
from pathlib import Path

from splitter.engine import Options, analyze, collect_inputs, default_keys, write_outputs


def main() -> None:
    ap = argparse.ArgumentParser(description="Split consolidated Excel files by group.")
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("-o", "--out", default="output")
    ap.add_argument("--file-key", help="column for output file name (or <输入文件名>)")
    ap.add_argument("--sheet-key", help="column for output sheet name (or <源工作表名>)")
    ap.add_argument("--meta", action="store_true", help="include 文件名称/工作表名称 columns")
    ap.add_argument("--stamp-filename", action="store_true")
    ap.add_argument("--stamp-row", action="store_true")
    ap.add_argument("--tz", default="Asia/Singapore")
    a = ap.parse_args()

    paths = collect_inputs(a.inputs)
    columns = analyze(paths).columns
    fk, sk = default_keys(columns)
    opt = Options(a.file_key or fk, a.sheet_key or sk, a.meta, a.stamp_filename, a.stamp_row, a.tz)
    res = analyze(paths, opt)
    for w in res.warnings:
        print("WARN:", w)
    for g in res.groups:
        print(f"{g.file_name} / {g.sheet_name}: {len(g.rows)} rows")
    for p in write_outputs(res.groups, Path(a.out), opt):
        print("wrote", p)


if __name__ == "__main__":
    main()
