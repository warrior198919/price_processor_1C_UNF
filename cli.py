#!/usr/bin/env python3
"""
Консольная версия обработчика прайс-листов → 1С УНФ
Использование:
  python cli.py <файл1.xlsx> [файл2.xlsx ...] [опции]

Опции:
  --brand SATA|RoxelPro|auto   Принудительный бренд (по умолчанию: auto)
  --out <папка>                 Папка для сохранения (по умолчанию: текущая)
  --price ue|rub               Тип цены (по умолчанию: rub)
  --split                      Разбить по группам товаров
  --no-photos                  Не включать ссылки на фото
"""

import sys
import os
import argparse

# Добавляем текущую папку в путь
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from parsers.sata_parser import SataParser
from parsers.roxelpro_parser import RoxelProParser
from exporters.unf_exporter import UNFExporter


def detect_brand(path: str) -> str:
    try:
        import pandas as pd
        name = os.path.basename(path).lower()
        if "sata" in name:
            return "SATA"
        if "roxel" in name:
            return "RoxelPro"
        xl = pd.ExcelFile(path)
        if "Sata" in xl.sheet_names or "SATA" in xl.sheet_names:
            return "SATA"
        df = pd.read_excel(path, header=None, nrows=10)
        text = " ".join(str(v) for v in df.values.flatten() if str(v) != "nan").lower()
        if "sata" in text:
            return "SATA"
        if "roxel" in text:
            return "RoxelPro"
    except Exception:
        pass
    return "unknown"


def main():
    parser = argparse.ArgumentParser(
        description="Конвертер прайс-листов SATA/RoxelPro → 1С УНФ"
    )
    parser.add_argument("files", nargs="+", help="Входные XLSX файлы")
    parser.add_argument("--brand", default="auto",
                        choices=["auto", "SATA", "RoxelPro"],
                        help="Формат прайса (default: auto)")
    parser.add_argument("--out", default=".",
                        help="Папка для сохранения результата")
    parser.add_argument("--price", default="rub", choices=["rub", "ue"],
                        help="Тип цены: rub (рубли) или ue (у.е.)")
    parser.add_argument("--split", action="store_true",
                        help="Разбить результат по группам товаров")
    parser.add_argument("--no-photos", action="store_true",
                        help="Не включать ссылки на фотографии")
    args = parser.parse_args()

    all_records = []

    for fpath in args.files:
        if not os.path.exists(fpath):
            print(f"[!] Файл не найден: {fpath}")
            continue

        brand = args.brand
        if brand == "auto":
            brand = detect_brand(fpath)
            print(f"[i] {os.path.basename(fpath)} → бренд: {brand}")

        if brand == "SATA":
            p = SataParser(fpath)
        elif brand == "RoxelPro":
            p = RoxelProParser(fpath)
        else:
            print(f"[!] Неизвестный формат: {fpath}. Используйте --brand")
            continue

        records = p.parse()
        print(f"[✓] {os.path.basename(fpath)}: {len(records)} записей")
        all_records.extend(records)

    if not all_records:
        print("[!] Нет данных для экспорта.")
        sys.exit(1)

    print(f"\n[i] Итого: {len(all_records)} записей")
    print(f"[i] Экспорт в: {os.path.abspath(args.out)}")

    exporter = UNFExporter(
        records=all_records,
        output_dir=args.out,
        include_photos=not args.no_photos,
        split_by_group=args.split,
        price_field="price_rub" if args.price == "rub" else "price_ue",
    )
    saved = exporter.export()

    print("\n[✓] Готово! Созданные файлы:")
    for f in saved:
        print(f"    {f}")


if __name__ == "__main__":
    main()
