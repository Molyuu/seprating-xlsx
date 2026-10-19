"""Excel 拆分工具 - desktop GUI (Tkinter, bundles into a single Windows .exe)."""
import os
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from splitter.engine import (
    INPUT_FILE, SOURCE_SHEET, Options, analyze, collect_inputs, default_keys,
    plan_outputs, write_outputs,
)

UI_FONT = "TkDefaultFont"
UI_FONT_CANDIDATES = (
    "Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", "Hiragino Sans GB",
)


class AccentButton(tk.Label):
    """A solid-colour button that renders consistently on Windows and macOS."""

    BG = "#007DBD"
    HOVER_BG = "#006BA3"
    PRESSED_BG = "#005A8A"
    DISABLED_BG = "#C2D6E2"
    DISABLED_FG = "#F3F7FA"

    def __init__(self, parent, text, command, font):
        self._command = command
        self._state = "disabled"
        self._default = "normal"
        self._hover = False
        self._pressed = False
        super().__init__(
            parent, text=text, font=font, bg=self.DISABLED_BG,
            fg=self.DISABLED_FG, padx=22, pady=7, cursor="arrow",
            relief="flat", borderwidth=0,
        )
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)

    def _apply_colours(self) -> None:
        if self._state == "disabled":
            bg, fg, cursor = self.DISABLED_BG, self.DISABLED_FG, "arrow"
        else:
            bg = self.PRESSED_BG if self._pressed else (
                self.HOVER_BG if self._hover else self.BG
            )
            fg, cursor = "white", "hand2"
        super().configure(background=bg, foreground=fg, cursor=cursor)

    def _set_state(self, state) -> None:
        self._state = "normal" if str(state) == "normal" else "disabled"
        if self._state == "disabled":
            self._hover = False
            self._pressed = False
        self._apply_colours()

    def configure(self, cnf=None, **kwargs):
        options = dict(cnf) if isinstance(cnf, dict) else {}
        options.update(kwargs)
        state = options.pop("state", None)
        if "default" in options:
            self._default = options.pop("default")
        if state is not None:
            self._set_state(state)
        if options:
            super().configure(**options)

    config = configure

    def __getitem__(self, key):
        if key == "state":
            return self._state
        if key == "default":
            return self._default
        return super().__getitem__(key)

    def __setitem__(self, key, value):
        if key in ("state", "default"):
            self.configure(**{key: value})
            return
        super().__setitem__(key, value)

    def invoke(self):
        if self._state == "normal":
            return self._command()
        return None

    def _pointer_over(self, event) -> bool:
        try:
            return self.winfo_containing(event.x_root, event.y_root) is self
        except (KeyError, tk.TclError):
            return 0 <= event.x < self.winfo_width() and 0 <= event.y < self.winfo_height()

    def _on_enter(self, _event) -> None:
        if self._state == "normal":
            self._hover = True
            self._apply_colours()

    def _on_leave(self, _event) -> None:
        self._hover = False
        self._pressed = False
        self._apply_colours()

    def _on_press(self, _event) -> None:
        if self._state == "normal":
            self._pressed = True
            self._apply_colours()

    def _on_release(self, event) -> None:
        was_pressed = self._pressed
        self._pressed = False
        over = self._pointer_over(event)
        self._hover = over and self._state == "normal"
        self._apply_colours()
        if was_pressed and over and self._state == "normal":
            self._command()


class App(tk.Tk):
    def __init__(self) -> None:
        global UI_FONT

        super().__init__()
        available_fonts = set(tkfont.families(self))
        UI_FONT = next(
            (family for family in UI_FONT_CANDIDATES if family in available_fonts),
            "TkDefaultFont",
        )
        self.ui_font = UI_FONT
        self.title("Excel 拆分工具")
        self.geometry("900x620")
        self.minsize(720, 480)
        self.paths = []
        self.result = None

        self.file_key = tk.StringVar()
        self.sheet_key = tk.StringVar()
        self.out_dir = tk.StringVar()
        self._build()

    # ------------------------------------------------------------------ layout
    def _build(self) -> None:
        s = ttk.Style(self)
        s.configure("Treeview", rowheight=26)
        s.configure("Treeview.Heading", padding=(6, 4))
        s.configure("TLabelframe.Label", font=(self.ui_font, 9, "bold"))

        f1 = ttk.LabelFrame(self, text="1. 输入", padding=(10, 6, 10, 10))
        f1.pack(fill="x", padx=12, pady=(8, 0))
        ttk.Button(f1, text="选择文件…", command=self.pick_files).pack(
            side="left", padx=(0, 8)
        )
        self.in_label = ttk.Label(f1, text="未选择文件")
        self.in_label.pack(side="left")

        f2 = ttk.LabelFrame(self, text="2. 分组依据", padding=(10, 6, 10, 10))
        f2.pack(fill="x", padx=12, pady=(8, 0))
        ttk.Label(f2, text="文件名来自：").pack(side="left", padx=(0, 8))
        self.cb_file = ttk.Combobox(f2, textvariable=self.file_key, state="readonly", width=22)
        self.cb_file.pack(side="left", padx=(0, 8))
        ttk.Label(f2, text="工作表名来自：").pack(side="left", padx=(24, 8))
        self.cb_sheet = ttk.Combobox(f2, textvariable=self.sheet_key, state="readonly", width=22)
        self.cb_sheet.pack(side="left")
        for cb in (self.cb_file, self.cb_sheet):
            cb.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        f3 = ttk.LabelFrame(self, text="3. 预览", padding=(10, 6, 10, 10))
        f3.pack(fill="both", expand=True, padx=12, pady=(8, 0))
        cols = ("file", "sheet", "rows", "subtotal", "source")
        self.tree = ttk.Treeview(f3, columns=cols, show="headings")
        column_specs = (
            ("file", "输出文件", 200, "w", True),
            ("sheet", "工作表", 120, "w", False),
            ("rows", "行数", 70, "e", False),
            ("subtotal", "小计行", 70, "e", False),
            ("source", "来源", 280, "w", True),
        )
        for c, t, w, anchor, stretch in column_specs:
            self.tree.heading(c, text=t, anchor=anchor)
            self.tree.column(
                c, width=w, minwidth=50, anchor=anchor, stretch=stretch
            )
        sb = ttk.Scrollbar(f3, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("hint", foreground="#888")
        self._show_hint()

        f5 = ttk.LabelFrame(self, text="4. 输出", padding=(10, 6, 10, 10))
        f5.pack(fill="x", padx=12, pady=(8, 12))
        ttk.Button(f5, text="输出文件夹…", command=self.pick_out).pack(
            side="left", padx=(0, 8)
        )
        ttk.Entry(
            f5, textvariable=self.out_dir, font=(self.ui_font, 9)
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.go = AccentButton(
            f5, text="开始拆分", command=self.run,
            font=(self.ui_font, 11, "bold"),
        )
        self.go.pack(side="right", padx=(12, 0))
        self.bind(
            "<Return>",
            lambda e: (
                self.go.invoke()
                if str(self.go["state"]) != "disabled"
                else None
            ),
        )

    # ------------------------------------------------------------------ actions
    def pick_files(self) -> None:
        files = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if files:
            self.load(list(files))

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
        if self.result and self.result.warnings:
            messagebox.showwarning("提示", "\n".join(self.result.warnings[:5]))

    def options(self) -> Options:
        return Options(self.file_key.get(), self.sheet_key.get())

    def refresh(self) -> None:
        if not self.paths:
            return
        opt = self.options()
        self.result = analyze(self.paths, opt)
        self.tree.delete(*self.tree.get_children())
        out = Path(self.out_dir.get() or ".")
        for plan in plan_outputs(self.result.groups, out, opt):
            for i, (sname, g) in enumerate(plan.sheets):
                self.tree.insert("", "end", values=(
                    plan.path.name if i == 0 else "",
                    sname, len(g.rows), g.subtotal_rows,
                    f"{g.table.source_file.name} / {g.table.sheet}"))
        self._show_hint()
        ready = bool(self.result.groups)
        # Only show the default-button ring when the action is actually available.
        self.go.config(state="normal" if ready else "disabled",
                       default="active" if ready else "normal")

    def _show_hint(self) -> None:
        if self.tree.get_children():
            return
        self.tree.insert(
            "", "end",
            values=("选择文件后在此预览拆分结果", "", "", "", ""),
            tags=("hint",),
        )

    def run(self) -> None:
        if not self.out_dir.get():
            messagebox.showwarning("提示", "请选择输出文件夹")
            return
        self.go.config(state="disabled")
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
            messagebox.showerror("错误", str(err))
            return
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
