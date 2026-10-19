"""Excel 拆分工具 - desktop GUI (Tkinter, bundles into a single Windows .exe)."""
import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from splitter.engine import (
    INPUT_FILE, SOURCE_SHEET, Options, analyze, collect_inputs, default_keys,
    plan_outputs, write_outputs,
)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel 拆分工具")
        self.geometry("900x620")
        self.minsize(720, 480)
        self.paths = []
        self.result = None

        self.file_key = tk.StringVar()
        self.sheet_key = tk.StringVar()
        self.include_meta = tk.BooleanVar(value=False)
        self.stamp_filename = tk.BooleanVar(value=False)
        self.stamp_row = tk.BooleanVar(value=False)
        self.tz = tk.StringVar(value="Asia/Singapore")
        self.out_dir = tk.StringVar()
        self.status = tk.StringVar(value="1) 选择输入文件或文件夹")
        self._build()

    # ------------------------------------------------------------------ layout
    def _build(self) -> None:
        pad = {"padx": 8, "pady": 4}
        f1 = ttk.LabelFrame(self, text="1. 输入")
        f1.pack(fill="x", **pad)
        ttk.Button(f1, text="选择文件…", command=self.pick_files).pack(side="left", **pad)
        ttk.Button(f1, text="选择文件夹…", command=self.pick_folder).pack(side="left", **pad)
        self.in_label = ttk.Label(f1, text="未选择")
        self.in_label.pack(side="left", **pad)

        f2 = ttk.LabelFrame(self, text="2. 分组依据")
        f2.pack(fill="x", **pad)
        ttk.Label(f2, text="文件名来自:").pack(side="left", **pad)
        self.cb_file = ttk.Combobox(f2, textvariable=self.file_key, state="readonly", width=22)
        self.cb_file.pack(side="left", **pad)
        ttk.Label(f2, text="工作表名来自:").pack(side="left", **pad)
        self.cb_sheet = ttk.Combobox(f2, textvariable=self.sheet_key, state="readonly", width=22)
        self.cb_sheet.pack(side="left", **pad)
        for cb in (self.cb_file, self.cb_sheet):
            cb.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        f3 = ttk.LabelFrame(self, text="3. 预览")
        f3.pack(fill="both", expand=True, **pad)
        cols = ("file", "sheet", "rows", "subtotal", "source")
        self.tree = ttk.Treeview(f3, columns=cols, show="headings")
        for c, t, w in zip(cols, ("输出文件", "工作表", "行数", "小计行", "来源"),
                           (220, 120, 60, 60, 300)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(f3, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        f4 = ttk.LabelFrame(self, text="4. 选项")
        f4.pack(fill="x", **pad)
        for text, var in (("包含 文件名称/工作表名称 列", self.include_meta),
                          ("文件名加日期", self.stamp_filename),
                          ("首行写入生成时间", self.stamp_row)):
            ttk.Checkbutton(f4, text=text, variable=var, command=self.refresh).pack(side="left", **pad)
        ttk.Label(f4, text="时区:").pack(side="left", **pad)
        ttk.Entry(f4, textvariable=self.tz, width=16).pack(side="left", **pad)

        f5 = ttk.LabelFrame(self, text="5. 输出")
        f5.pack(fill="x", **pad)
        ttk.Button(f5, text="输出文件夹…", command=self.pick_out).pack(side="left", **pad)
        ttk.Entry(f5, textvariable=self.out_dir).pack(side="left", fill="x", expand=True, **pad)
        self.go = ttk.Button(f5, text="开始拆分", command=self.run, state="disabled")
        self.go.pack(side="right", **pad)

        ttk.Label(self, textvariable=self.status, foreground="#444").pack(fill="x", padx=10, pady=(0, 8))

    # ------------------------------------------------------------------ actions
    def pick_files(self) -> None:
        files = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if files:
            self.load(list(files))

    def pick_folder(self) -> None:
        d = filedialog.askdirectory()
        if d:
            self.load([d])

    def pick_out(self) -> None:
        d = filedialog.askdirectory()
        if d:
            self.out_dir.set(d)

    def load(self, selection) -> None:
        self.paths = collect_inputs(selection)
        if not self.paths:
            messagebox.showwarning("提示", "未找到 .xlsx 文件")
            return
        self.in_label.config(text=f"{len(self.paths)} 个文件: "
                             + ", ".join(p.name for p in self.paths[:3])
                             + (" …" if len(self.paths) > 3 else ""))
        if not self.out_dir.get():
            self.out_dir.set(str(self.paths[0].parent / "拆分结果"))
        columns = analyze(self.paths).columns
        choices = [INPUT_FILE, SOURCE_SHEET] + columns
        self.cb_file["values"] = self.cb_sheet["values"] = choices
        fk, sk = default_keys(columns)
        self.file_key.set(fk)
        self.sheet_key.set(sk)
        self.refresh()

    def options(self) -> Options:
        return Options(self.file_key.get(), self.sheet_key.get(), self.include_meta.get(),
                       self.stamp_filename.get(), self.stamp_row.get(), self.tz.get().strip())

    def refresh(self) -> None:
        if not self.paths:
            return
        opt = self.options()
        self.result = analyze(self.paths, opt)
        self.tree.delete(*self.tree.get_children())
        out = Path(self.out_dir.get() or ".")
        for plan in plan_outputs(self.result.groups, out, opt):
            for sname, g in plan.sheets:
                self.tree.insert("", "end", values=(
                    plan.path.name, sname, len(g.rows), g.subtotal_rows,
                    f"{g.table.source_file.name} / {g.table.sheet}"))
        msg = f"{len(self.result.groups)} 个分组"
        if self.result.warnings:
            msg += " | 警告: " + "; ".join(self.result.warnings[:3])
        self.status.set(msg)
        self.go.config(state="normal" if self.result.groups else "disabled")

    def run(self) -> None:
        if not self.out_dir.get():
            messagebox.showwarning("提示", "请选择输出文件夹")
            return
        self.go.config(state="disabled")
        self.status.set("正在生成…")
        opt, out = self.options(), Path(self.out_dir.get())

        def work():
            try:
                res = analyze(self.paths, opt)
                files = write_outputs(res.groups, out, opt)
                self.after(0, self.done, files, out, None)
            except Exception as e:
                self.after(0, self.done, [], out, e)

        threading.Thread(target=work, daemon=True).start()

    def done(self, files, out: Path, err) -> None:
        self.go.config(state="normal")
        if err:
            self.status.set(f"失败: {err}")
            messagebox.showerror("错误", str(err))
            return
        self.status.set(f"完成: 生成 {len(files)} 个文件 → {out}")
        if messagebox.askyesno("完成", f"已生成 {len(files)} 个文件。\n打开输出文件夹?"):
            open_folder(out)


def open_folder(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)])
    else:
        subprocess.run(["xdg-open", str(path)])


if __name__ == "__main__":
    App().mainloop()
