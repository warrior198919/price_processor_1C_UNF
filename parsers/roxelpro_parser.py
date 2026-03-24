"""
Парсер прайс-листа RoxelPro (.xlsx)
Структура: строка 0 – заголовки, далее данные.
Колонки: Номенклатура, Артикул, Описание, Количество, Цена розничная (руб),
         Ед. изм., Валюта, Бренд, Фотографии, Vendorcode, Full path,
         Штрихкод, Цена продажи руб, Количество остаток лок склад, Подгруппа
"""

import pandas as pd


class RoxelProParser:

    COLUMN_MAP = {
        "Номенклатура": ["Номенклатура", "номенклатура", "name"],
        "Артикул": ["Артикул", "артикул", "article"],
        "Описание": ["Описание", "описание", "description"],
        "Количество": ["Количество", "кол", "qty", "quantity"],
        "Цена розничная (руб)": ["Цена розничная (руб)", "Цена розничная", "цена", "price"],
        "Ед. изм.": ["Ед. изм.", "Единица", "unit"],
        "Валюта": ["Валюта", "currency"],
        "Бренд": ["Бренд", "brand"],
        "Фотографии": ["Фотографии", "фото", "photo", "image"],
        "Vendorcode": ["Vendorcode", "vendorcode", "vendor"],
        "Full path": ["Full path", "fullpath", "path"],
        "Штрихкод": ["Штрихкод", "barcode"],
        "Цена продажи руб": ["Цена продажи руб", "Цена продажи"],
        "Количество остаток лок склад": ["Количество остаток лок склад", "остаток"],
        "Подгруппа": ["Подгруппа", "subgroup", "группа"],
    }

    def __init__(self, path: str):
        self.path = path

    def parse(self) -> list[dict]:
        xl = pd.ExcelFile(self.path)
        records = []

        for sheet in xl.sheet_names:
            df = pd.read_excel(self.path, sheet_name=sheet, header=None)
            header_row = self._find_header_row(df)
            if header_row is None:
                continue

            df.columns = df.iloc[header_row]
            df = df.iloc[header_row + 1:].reset_index(drop=True)
            df.columns = [str(c).strip() for c in df.columns]

            col_map = self._map_columns(df.columns.tolist())

            for _, row in df.iterrows():
                nom = self._get(row, col_map, "Номенклатура")
                if not nom or str(nom) == "nan":
                    continue
                art = self._get(row, col_map, "Артикул")
                desc = self._get(row, col_map, "Описание") or nom
                qty = self._get(row, col_map, "Количество")
                price_rub = self._to_float(self._get(row, col_map, "Цена розничная (руб)"))
                unit = self._get(row, col_map, "Ед. изм.") or "шт."
                brand = self._get(row, col_map, "Бренд") or "RoxelPro"
                photo = self._get(row, col_map, "Фотографии") or ""
                vendorcode = self._get(row, col_map, "Vendorcode") or art
                full_path = self._get(row, col_map, "Full path") or f"RoxelPro\\"
                barcode = self._get(row, col_map, "Штрихкод") or ""
                subgroup = self._get(row, col_map, "Подгруппа") or ""

                records.append({
                    "Номенклатура": str(nom).strip(),
                    "Артикул": str(art).strip() if str(art) != "nan" else "",
                    "Описание": str(desc).strip(),
                    "Вариант": "",
                    "Количество": str(qty).strip() if str(qty) != "nan" else "",
                    "Цена розничная (у.е.)": 0.0,
                    "Цена розничная (руб)": price_rub,
                    "Ед. изм.": str(unit).strip(),
                    "Валюта": "RUB",
                    "Бренд": str(brand).strip(),
                    "Vendorcode": str(vendorcode).strip() if str(vendorcode) != "nan" else "",
                    "Full path": str(full_path).strip(),
                    "Подгруппа": str(subgroup).strip(),
                    "Фотографии": str(photo).strip() if str(photo) != "nan" else "",
                    "Штрихкод": str(barcode).strip() if str(barcode) != "nan" else "",
                })

        return records

    # ── helpers ─────────────────────────────────────────────────────────────

    def _find_header_row(self, df: pd.DataFrame):
        """Найти строку с заголовками (ищем 'Номенклатура' или 'Артикул')."""
        for i, row in df.iterrows():
            vals = [str(v).strip().lower() for v in row if str(v) != "nan"]
            if any(k in vals for k in ["номенклатура", "артикул", "наименование"]):
                return i
        return None

    def _map_columns(self, columns: list) -> dict:
        """Сопоставить реальные названия колонок с нашими ключами."""
        result = {}
        col_lower = {c.lower(): c for c in columns}
        for key, aliases in self.COLUMN_MAP.items():
            for alias in aliases:
                if alias.lower() in col_lower:
                    result[key] = col_lower[alias.lower()]
                    break
        return result

    def _get(self, row, col_map: dict, key: str):
        col = col_map.get(key)
        if col and col in row.index:
            v = row[col]
            return v if str(v) != "nan" else ""
        return ""

    @staticmethod
    def _to_float(v) -> float:
        try:
            return float(str(v).replace(",", ".").replace(" ", ""))
        except Exception:
            return 0.0
