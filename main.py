#!/usr/bin/env python3
"""
Обработчик прайс-листов SATA / RoxelPro → 1С УНФ
Версия 1.0
"""

import sys
import os

# Проверка зависимостей
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except ModuleNotFoundError:
    print("ОШИБКА: tkinter не установлен. Установите python3-tk:")
    print("  Ubuntu/Debian: sudo apt-get install python3-tk")
    print("  Windows: входит в стандартный Python")
    sys.exit(1)

try:
    import pandas as pd
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError as e:
    print(f"ОШИБКА: не хватает библиотеки: {e}")
    print("Установите: pip install pandas openpyxl")
    sys.exit(1)

import threading
import re
from datetime import datetime
from parsers.sata_parser import SataParser
from parsers.roxelpro_parser import RoxelProParser
from exporters.unf_exporter import UNFExporter


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Обработчик прайс-листов → 1С УНФ")
        self.geometry("900x700")
        self.resizable(True, True)
        self.configure(bg="#f0f0f0")

        self.input_files = []
        self.output_dir = tk.StringVar(value=os.path.expanduser("~"))
        self.brand_var = tk.StringVar(value="auto")
        self.include_photos = tk.BooleanVar(value=True)
        self.split_by_group = tk.BooleanVar(value=False)
        self.price_type = tk.StringVar(value="rub")  # rub / ue

        self._build_ui()

    # ── UI ──────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Заголовок
        hdr = tk.Frame(self, bg="#1a3a5c", height=55)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🏷 Прайс-лист → 1С УНФ", font=("Arial", 16, "bold"),
                 bg="#1a3a5c", fg="white").pack(side="left", padx=18, pady=12)

        # Основная область
        main = tk.Frame(self, bg="#f0f0f0")
        main.pack(fill="both", expand=True, padx=16, pady=12)

        # ── БЛОК 1: Входные файлы ─────────────────────────────────────────
        self._section(main, "1. Входные файлы XLSX")

        file_frame = tk.Frame(main, bg="#f0f0f0")
        file_frame.pack(fill="x", pady=(0, 8))

        self.file_listbox = tk.Listbox(file_frame, height=4, font=("Courier", 9),
                                        selectmode=tk.EXTENDED)
        self.file_listbox.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(file_frame, command=self.file_listbox.yview)
        sb.pack(side="right", fill="y")
        self.file_listbox.config(yscrollcommand=sb.set)

        btn_row = tk.Frame(main, bg="#f0f0f0")
        btn_row.pack(fill="x", pady=(0, 10))
        tk.Button(btn_row, text="➕ Добавить файлы", command=self._add_files,
                  **self._btn_style()).pack(side="left", padx=(0, 6))
        tk.Button(btn_row, text="🗑 Удалить выбранные", command=self._remove_files,
                  **self._btn_style("red")).pack(side="left")

        # ── БЛОК 2: Настройки ────────────────────────────────────────────
        self._section(main, "2. Настройки")

        cfg = tk.Frame(main, bg="#f0f0f0")
        cfg.pack(fill="x", pady=(0, 10))

        # Бренд
        tk.Label(cfg, text="Формат прайса:", bg="#f0f0f0", font=("Arial", 9)).grid(
            row=0, column=0, sticky="w", padx=(0, 8))
        brand_cb = ttk.Combobox(cfg, textvariable=self.brand_var, width=18,
                                 values=["auto", "SATA", "RoxelPro"])
        brand_cb.grid(row=0, column=1, sticky="w", padx=(0, 20))

        # Тип цены
        tk.Label(cfg, text="Цена:", bg="#f0f0f0", font=("Arial", 9)).grid(
            row=0, column=2, sticky="w", padx=(0, 8))
        ttk.Combobox(cfg, textvariable=self.price_type, width=12,
                     values=["rub", "ue"]).grid(row=0, column=3, sticky="w", padx=(0, 20))

        # Чекбоксы
        tk.Checkbutton(cfg, text="Включать фото-ссылки", variable=self.include_photos,
                       bg="#f0f0f0").grid(row=0, column=4, sticky="w", padx=(0, 10))
        tk.Checkbutton(cfg, text="Разбить по группам", variable=self.split_by_group,
                       bg="#f0f0f0").grid(row=0, column=5, sticky="w")

        # ── БЛОК 3: Папка вывода ─────────────────────────────────────────
        self._section(main, "3. Папка сохранения")
        out_row = tk.Frame(main, bg="#f0f0f0")
        out_row.pack(fill="x", pady=(0, 10))
        tk.Entry(out_row, textvariable=self.output_dir, width=65,
                 font=("Courier", 9)).pack(side="left", padx=(0, 8))
        tk.Button(out_row, text="📁 Выбрать", command=self._choose_dir,
                  **self._btn_style()).pack(side="left")

        # ── КНОПКА ЗАПУСКА ───────────────────────────────────────────────
        tk.Button(main, text="▶  Обработать и экспортировать",
                  command=self._run, font=("Arial", 11, "bold"),
                  bg="#1a7a3c", fg="white", relief="flat",
                  padx=20, pady=8).pack(pady=12)

        # ── ЛОГ ──────────────────────────────────────────────────────────
        self._section(main, "Журнал")
        self.log = scrolledtext.ScrolledText(main, height=12, font=("Courier", 9),
                                              state="disabled", bg="#1e1e1e", fg="#d4d4d4")
        self.log.pack(fill="both", expand=True)

        # Статусбар
        self.status = tk.StringVar(value="Готов к работе")
        tk.Label(self, textvariable=self.status, bg="#dde", font=("Arial", 9),
                 anchor="w").pack(fill="x", side="bottom", padx=8, pady=2)

    def _section(self, parent, title):
        tk.Label(parent, text=title, font=("Arial", 10, "bold"),
                 bg="#f0f0f0", fg="#1a3a5c").pack(anchor="w", pady=(6, 2))

    def _btn_style(self, color="blue"):
        colors = {
            "blue": ("#1a5fa8", "white"),
            "red": ("#b03030", "white"),
        }
        bg, fg = colors.get(color, ("#555", "white"))
        return dict(bg=bg, fg=fg, relief="flat", padx=10, pady=4, font=("Arial", 9))

    # ── Действия ────────────────────────────────────────────────────────────
    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="Выберите XLSX прайсы",
            filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")]
        )
        for f in files:
            if f not in self.input_files:
                self.input_files.append(f)
                self.file_listbox.insert(tk.END, os.path.basename(f))

    def _remove_files(self):
        selected = list(self.file_listbox.curselection())
        for i in reversed(selected):
            self.file_listbox.delete(i)
            self.input_files.pop(i)

    def _choose_dir(self):
        d = filedialog.askdirectory(title="Папка для сохранения")
        if d:
            self.output_dir.set(d)

    def _log(self, msg, tag=None):
        self.log.config(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self.log.insert(tk.END, line)
        self.log.see(tk.END)
        self.log.config(state="disabled")

    def _run(self):
        if not self.input_files:
            messagebox.showwarning("Нет файлов", "Добавьте хотя бы один XLSX прайс.")
            return
        threading.Thread(target=self._process, daemon=True).start()

    def _process(self):
        self.status.set("⏳ Обработка...")
        self._log("=== Начало обработки ===")

        all_records = []
        for path in self.input_files:
            fname = os.path.basename(path)
            self._log(f"Читаю: {fname}")
            try:
                brand = self.brand_var.get()
                if brand == "auto":
                    brand = detect_brand(path)
                    self._log(f"  Определён бренд: {brand}")

                if brand == "SATA":
                    parser = SataParser(path)
                elif brand == "RoxelPro":
                    parser = RoxelProParser(path)
                else:
                    self._log(f"  ⚠ Неизвестный бренд, пропуск.")
                    continue

                records = parser.parse()
                self._log(f"  ✓ Извлечено записей: {len(records)}")
                all_records.extend(records)
            except Exception as e:
                self._log(f"  ✗ ОШИБКА: {e}")

        if not all_records:
            self._log("Нет данных для экспорта.")
            self.status.set("⚠ Нет данных")
            return

        self._log(f"Итого записей: {len(all_records)}")
        self._log("Экспорт в формат 1С УНФ...")

        try:
            exporter = UNFExporter(
                records=all_records,
                output_dir=self.output_dir.get(),
                include_photos=self.include_photos.get(),
                split_by_group=self.split_by_group.get(),
                price_field="price_rub" if self.price_type.get() == "rub" else "price_ue",
            )
            output_files = exporter.export()
            for f in output_files:
                self._log(f"  💾 Сохранён: {os.path.basename(f)}")
            self._log("=== Готово! ===")
            self.status.set(f"✅ Готово. Файлов: {len(output_files)}")
            messagebox.showinfo("Готово", f"Экспорт завершён!\nФайлов: {len(output_files)}")
        except Exception as e:
            self._log(f"ОШИБКА экспорта: {e}")
            self.status.set("✗ Ошибка экспорта")
            messagebox.showerror("Ошибка", str(e))


def detect_brand(path: str) -> str:
    """Автоопределение бренда по содержимому файла."""
    try:
        xl = pd.ExcelFile(path)
        sheets = xl.sheet_names
        name = os.path.basename(path).lower()
        if "sata" in name or "Sata" in sheets:
            return "SATA"
        if "roxel" in name or "roxelpro" in name.replace("-", "").replace("_", ""):
            return "RoxelPro"
        # Читаем первые строки
        df = pd.read_excel(path, header=None, nrows=10)
        text = " ".join(str(v) for v in df.values.flatten() if str(v) != "nan").lower()
        if "sata" in text:
            return "SATA"
        if "roxelpro" in text or "roxel" in text:
            return "RoxelPro"
    except Exception:
        pass
    return "unknown"


if __name__ == "__main__":
    app = App()
    app.mainloop()
