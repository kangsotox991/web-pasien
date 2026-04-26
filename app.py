#!/usr/bin/env python3
"""
Excel Logbook Editor
Desktop application untuk membaca, mengedit, dan menyimpan file Excel.
Fitur: buka file, edit cell, import data pasien, rename sheet otomatis, simpan.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import re
from collections import OrderedDict
from copy import copy

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

BULAN_INDO = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}


class ExcelEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Logbook Editor")
        self.root.geometry("1200x700")
        self.root.minsize(900, 500)

        self.workbook = None
        self.filepath = None
        self.current_sheet_name = None
        self.modified = False

        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        # Menu bar
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Buka File Excel...", command=self.open_file, accelerator="Ctrl+O")
        file_menu.add_command(label="Simpan", command=self.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="Simpan Sebagai...", command=self.save_as, accelerator="Ctrl+Shift+S")
        file_menu.add_separator()
        file_menu.add_command(label="Keluar", command=self.on_close)
        menubar.add_cascade(label="File", menu=file_menu)

        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Import Data Pasien...", command=self.import_data_pasien)
        tools_menu.add_command(label="Rename Sheet...", command=self.rename_current_sheet)
        menubar.add_cascade(label="Tools", menu=tools_menu)

        self.root.config(menu=menubar)
        self.root.bind("<Control-o>", lambda e: self.open_file())
        self.root.bind("<Control-s>", lambda e: self.save_file())
        self.root.bind("<Control-S>", lambda e: self.save_as())
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Toolbar
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=8, pady=(8, 0))

        ttk.Button(toolbar, text="Buka File", command=self.open_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Simpan", command=self.save_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Simpan Sebagai", command=self.save_as).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)
        ttk.Button(toolbar, text="Import Data Pasien", command=self.import_data_pasien).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Rename Sheet", command=self.rename_current_sheet).pack(side=tk.LEFT, padx=2)

        self.file_label = ttk.Label(toolbar, text="Belum ada file dibuka", foreground="gray")
        self.file_label.pack(side=tk.RIGHT, padx=8)

        self.status_label = ttk.Label(toolbar, text="", foreground="green")
        self.status_label.pack(side=tk.RIGHT, padx=8)

        # Sheet tabs
        self.tab_frame = ttk.Frame(self.root)
        self.tab_frame.pack(fill=tk.X, padx=8, pady=(8, 0))

        self.tab_canvas = tk.Canvas(self.tab_frame, height=32, highlightthickness=0)
        self.tab_scrollbar = ttk.Scrollbar(self.tab_frame, orient=tk.HORIZONTAL, command=self.tab_canvas.xview)
        self.tab_inner = ttk.Frame(self.tab_canvas)

        self.tab_canvas.configure(xscrollcommand=self.tab_scrollbar.set)
        self.tab_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.tab_canvas.pack(side=tk.TOP, fill=tk.X)
        self.tab_canvas.create_window((0, 0), window=self.tab_inner, anchor="nw")
        self.tab_inner.bind("<Configure>", lambda e: self.tab_canvas.configure(scrollregion=self.tab_canvas.bbox("all")))

        # Table area
        table_frame = ttk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.tree = ttk.Treeview(table_frame, show="headings", selectmode="browse")

        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self.on_cell_double_click)

        # Status bar
        self.statusbar = ttk.Label(self.root, text="Siap", relief=tk.SUNKEN, anchor=tk.W)
        self.statusbar.pack(fill=tk.X, side=tk.BOTTOM)

        # Welcome screen
        self._show_welcome()

    def _show_welcome(self):
        for widget in self.tab_inner.winfo_children():
            widget.destroy()
        self.tree["columns"] = ()
        self.tree.delete(*self.tree.get_children())

    # ------------------------------------------------------------------ FILE OPS
    def open_file(self):
        path = filedialog.askopenfilename(
            title="Pilih File Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            self.workbook = load_workbook(path)
            self.filepath = path
            self.modified = False
            self.file_label.config(text=path.split("/")[-1].split("\\")[-1], foreground="black")
            self._render_tabs()
            self._switch_sheet(self.workbook.sheetnames[0])
            self.statusbar.config(text=f"File dibuka: {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuka file:\n{e}")

    def save_file(self):
        if not self.workbook or not self.filepath:
            self.save_as()
            return
        try:
            self.workbook.save(self.filepath)
            self.modified = False
            self.status_label.config(text="Tersimpan", foreground="green")
            self.statusbar.config(text=f"File disimpan: {self.filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan:\n{e}")

    def save_as(self):
        if not self.workbook:
            messagebox.showwarning("Peringatan", "Belum ada file yang dibuka")
            return
        path = filedialog.asksaveasfilename(
            title="Simpan Sebagai",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            self.workbook.save(path)
            self.filepath = path
            self.modified = False
            self.file_label.config(text=path.split("/")[-1].split("\\")[-1], foreground="black")
            self.status_label.config(text="Tersimpan", foreground="green")
            self.statusbar.config(text=f"File disimpan: {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan:\n{e}")

    def on_close(self):
        if self.modified:
            ans = messagebox.askyesnocancel("Simpan?", "Ada perubahan yang belum disimpan. Simpan dulu?")
            if ans is None:
                return
            if ans:
                self.save_file()
        self.root.destroy()

    # ------------------------------------------------------------------ SHEET TABS
    def _render_tabs(self):
        for w in self.tab_inner.winfo_children():
            w.destroy()
        if not self.workbook:
            return
        for name in self.workbook.sheetnames:
            style = "primary" if name == self.current_sheet_name else "default"
            btn = tk.Button(
                self.tab_inner, text=name, padx=12, pady=4,
                relief=tk.RAISED if name != self.current_sheet_name else tk.SUNKEN,
                bg="#2563eb" if name == self.current_sheet_name else "#f0f0f0",
                fg="white" if name == self.current_sheet_name else "black",
                font=("Segoe UI", 9, "bold" if name == self.current_sheet_name else "normal"),
                command=lambda n=name: self._switch_sheet(n)
            )
            btn.pack(side=tk.LEFT, padx=1, pady=2)

    def _switch_sheet(self, name):
        self.current_sheet_name = name
        self._render_tabs()
        self._render_table()

    # ------------------------------------------------------------------ TABLE
    def _render_table(self):
        self.tree.delete(*self.tree.get_children())

        if not self.workbook or not self.current_sheet_name:
            return

        ws = self.workbook[self.current_sheet_name]
        max_col = ws.max_column or 1
        max_row = ws.max_row or 1

        cols = ["#"] + [get_column_letter(c) for c in range(1, max_col + 1)]
        self.tree["columns"] = cols
        for c in cols:
            w = 50 if c == "#" else 140
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, minwidth=40, stretch=True)

        for row_idx in range(1, max_row + 1):
            values = [str(row_idx)]
            for col_idx in range(1, max_col + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                val = cell.value
                if val is None:
                    val = ""
                elif isinstance(val, str) and val.startswith("="):
                    # Show formula result if available
                    val = val
                values.append(str(val))
            self.tree.insert("", tk.END, values=values, tags=(f"row_{row_idx}",))

        self.statusbar.config(text=f"Sheet: {self.current_sheet_name} | {max_row} baris x {max_col} kolom")

    # ------------------------------------------------------------------ CELL EDIT
    def on_cell_double_click(self, event):
        if not self.workbook:
            return

        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        item = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)

        if not item or not col_id:
            return

        col_index = int(col_id.replace("#", "")) - 1  # 0-based
        if col_index == 0:  # row number column
            return

        values = self.tree.item(item, "values")
        row_num = int(values[0])
        col_num = col_index  # 1-based Excel column
        current_val = values[col_index]

        # Get cell bbox
        bbox = self.tree.bbox(item, col_id)
        if not bbox:
            return

        x, y, w, h = bbox

        entry = tk.Entry(self.tree, font=("Segoe UI", 10))
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, current_val)
        entry.select_range(0, tk.END)
        entry.focus()

        def finish_edit(save=True):
            if save:
                new_val = entry.get()
                if new_val != current_val:
                    ws = self.workbook[self.current_sheet_name]
                    ws.cell(row=row_num, column=col_num, value=new_val)
                    self.modified = True
                    self.status_label.config(text="Belum disimpan", foreground="orange")
                    self._render_table()
            entry.destroy()

        entry.bind("<Return>", lambda e: finish_edit(True))
        entry.bind("<Escape>", lambda e: finish_edit(False))
        entry.bind("<FocusOut>", lambda e: finish_edit(True))
        entry.bind("<Tab>", lambda e: finish_edit(True))

    # ------------------------------------------------------------------ RENAME SHEET
    def rename_current_sheet(self):
        if not self.workbook or not self.current_sheet_name:
            messagebox.showwarning("Peringatan", "Belum ada file yang dibuka")
            return

        new_name = simpledialog.askstring(
            "Rename Sheet",
            f"Nama baru untuk '{self.current_sheet_name}':",
            initialvalue=self.current_sheet_name,
            parent=self.root
        )
        if not new_name or new_name == self.current_sheet_name:
            return

        # Excel sheet name max 31 chars, no special chars
        new_name = re.sub(r'[\\/*?\[\]:]', '', new_name)[:31]

        ws = self.workbook[self.current_sheet_name]
        ws.title = new_name
        self.current_sheet_name = new_name
        self.modified = True
        self.status_label.config(text="Belum disimpan", foreground="orange")
        self._render_tabs()
        self.statusbar.config(text=f"Sheet renamed → {new_name}")

    # ------------------------------------------------------------------ IMPORT DATA PASIEN
    def import_data_pasien(self):
        if not self.workbook:
            messagebox.showwarning("Peringatan", "Buka file Excel template terlebih dahulu")
            return

        win = tk.Toplevel(self.root)
        win.title("Import Data Pasien")
        win.geometry("750x650")
        win.transient(self.root)
        win.grab_set()

        ttk.Label(win, text="Import Data Pasien", font=("Segoe UI", 14, "bold")).pack(pady=(12, 4))
        ttk.Label(win, text=(
            "Paste data pasien di bawah, atau load dari file .txt.\n"
            "Format per baris: DD/MM/YYYY Nama No. Reg XXXXXXX dx: diagnosa\n\n"
            "Data akan dikelompokkan per tanggal → 1 tanggal = 1 sheet.\n"
            "Sheet di-rename otomatis (misal: '1 Maret', '4 Maret', dst)."
        ), justify=tk.LEFT, foreground="gray").pack(padx=16, anchor="w")

        # Load from file button
        load_frame = ttk.Frame(win)
        load_frame.pack(fill=tk.X, padx=16, pady=(4, 0))

        def load_from_txt():
            path = filedialog.askopenfilename(
                title="Pilih File Data Pasien",
                filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"), ("All files", "*.*")],
                parent=win
            )
            if not path:
                return
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
            except UnicodeDecodeError:
                with open(path, "r", encoding="latin-1") as f:
                    content = f.read()
            text.delete("1.0", tk.END)
            text.insert("1.0", content)
            result_label.config(
                text=f"File dimuat: {path.split('/')[-1].split(chr(92))[-1]}",
                foreground="green"
            )

        ttk.Button(load_frame, text="Load dari File .txt", command=load_from_txt).pack(side=tk.LEFT)
        ttk.Label(load_frame, text="  atau paste langsung di bawah", foreground="gray").pack(side=tk.LEFT)

        text_frame = ttk.Frame(win)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=8)

        text = tk.Text(text_frame, font=("Consolas", 10), wrap=tk.NONE)
        text_vsb = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text.yview)
        text_hsb = ttk.Scrollbar(text_frame, orient=tk.HORIZONTAL, command=text.xview)
        text.configure(yscrollcommand=text_vsb.set, xscrollcommand=text_hsb.set)

        text.grid(row=0, column=0, sticky="nsew")
        text_vsb.grid(row=0, column=1, sticky="ns")
        text_hsb.grid(row=1, column=0, sticky="ew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill=tk.X, padx=16, pady=(0, 12))

        result_label = ttk.Label(btn_frame, text="", foreground="blue")
        result_label.pack(side=tk.LEFT)

        def preview():
            raw = text.get("1.0", tk.END).strip()
            if not raw:
                result_label.config(text="Tidak ada data", foreground="red")
                return
            groups = self._parse_patient_data(raw)
            if not groups:
                result_label.config(text="Format data tidak dikenali", foreground="red")
                return
            dates = list(groups.keys())
            total = sum(len(v) for v in groups.values())
            result_label.config(
                text=f"Ditemukan {len(dates)} tanggal, {total} pasien. Sheet tersedia: {len(self.workbook.sheetnames)}",
                foreground="blue"
            )

        def do_import():
            raw = text.get("1.0", tk.END).strip()
            if not raw:
                messagebox.showwarning("Peringatan", "Tidak ada data", parent=win)
                return
            groups = self._parse_patient_data(raw)
            if not groups:
                messagebox.showerror("Error", "Format data tidak dikenali", parent=win)
                return

            dates = list(groups.keys())
            sheet_count = len(self.workbook.sheetnames)

            if len(dates) > sheet_count:
                messagebox.showwarning(
                    "Peringatan",
                    f"Ada {len(dates)} tanggal tapi hanya {sheet_count} sheet.\n"
                    f"Hanya {sheet_count} tanggal pertama yang akan diproses.",
                    parent=win
                )

            count = min(len(dates), sheet_count)
            for i in range(count):
                date_key = dates[i]
                day, month, year = date_key
                sheet_name = f"{day} {BULAN_INDO[month]}"

                # Sanitize sheet name
                sheet_name = re.sub(r'[\\/*?\[\]:]', '', sheet_name)[:31]

                ws = self.workbook.worksheets[i]
                ws.title = sheet_name

            self.modified = True
            self.status_label.config(text="Belum disimpan", foreground="orange")
            self.current_sheet_name = self.workbook.sheetnames[0]
            self._render_tabs()
            self._render_table()

            messagebox.showinfo(
                "Sukses",
                f"Berhasil rename {count} sheet sesuai tanggal!",
                parent=win
            )
            win.destroy()

        ttk.Button(btn_frame, text="Preview", command=preview).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_frame, text="Import & Rename", command=do_import).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_frame, text="Batal", command=win.destroy).pack(side=tk.RIGHT, padx=4)

    def _parse_patient_data(self, raw_text):
        """Parse patient data text and group by date.
        Returns OrderedDict: {(day, month, year): [lines...]}
        """
        groups = OrderedDict()
        pattern = re.compile(
            r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})\s+(.+)'
        )

        for line in raw_text.strip().splitlines():
            line = line.strip()
            if not line:
                continue
            m = pattern.match(line)
            if m:
                day = int(m.group(1))
                month = int(m.group(2))
                year = int(m.group(3))
                rest = m.group(4).strip()
                key = (day, month, year)
                if key not in groups:
                    groups[key] = []
                groups[key].append(rest)

        return groups


def main():
    root = tk.Tk()

    # Style
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", rowheight=28, font=("Segoe UI", 10))
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

    app = ExcelEditorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
