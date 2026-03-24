"""
Парсер прайс-листа SATA (.xlsx)
Структура: лист 'Sata', строки с продуктами имеют col[0]=0,
           за каждым продуктом идут 2 строки: ярлыки дюз и артикулы.
"""

import pandas as pd


class SataParser:
    SHEET = "Sata"

    def __init__(self, path: str):
        self.path = path

    def parse(self) -> list[dict]:
        df = pd.read_excel(self.path, sheet_name=self.SHEET, header=None)
        records = []
        current_group = ""
        i = 6  # первые 6 строк – служебные

        while i < len(df):
            row = df.iloc[i]
            non_null = [(j, v) for j, v in enumerate(row) if str(v) != "nan"]
            cols = [x[0] for x in non_null]

            # ── Заголовок группы (только колонка 1, текст) ────────────────
            if cols == [1]:
                val = str(non_null[0][1]).strip()
                if not self._is_numeric(val):
                    current_group = val
                i += 1
                continue

            # ── Строка продукта: col[0] == 0 ──────────────────────────────
            if non_null and non_null[0][0] == 0 and non_null[0][1] == 0:
                name = str(row.iloc[1]).strip() if str(row.iloc[1]) != "nan" else ""
                price_ue = self._to_float(row.iloc[10])
                price_rub = self._to_float(row.iloc[11])

                # Следующие 2 строки: ярлыки дюз + артикулы
                labels, articles = self._read_variants(df, i + 1)

                if articles:
                    for label, art in zip(labels, articles):
                        nom = f"SATA {name} {label}".strip() if label else f"SATA {name}"
                        records.append(self._make_record(
                            nom, art, name, price_ue, price_rub, current_group, label
                        ))
                    i += 3
                    continue
                else:
                    records.append(self._make_record(
                        f"SATA {name}", "", name, price_ue, price_rub, current_group, ""
                    ))

            i += 1

        return records

    # ── helpers ───────────────────────────────────────────────────────────

    def _read_variants(self, df, start_idx: int):
        """Читает строку ярлыков и строку артикулов после продуктовой строки."""
        labels, articles = [], []

        if start_idx >= len(df):
            return labels, articles

        label_row = df.iloc[start_idx]
        label_vals = [(j, str(v).strip()) for j, v in enumerate(label_row) if str(v) != "nan"]

        # Это строка ярлыков если значения типа '1,1 I', '1,2 O'
        is_label = label_vals and any(
            ("," in v or "I" in v or "O" in v or "SR" in v)
            for _, v in label_vals
        )

        if is_label and start_idx + 1 < len(df):
            art_row = df.iloc[start_idx + 1]
            art_vals = [v for v in art_row if str(v) != "nan"]
            for i, (_, lbl) in enumerate(label_vals):
                labels.append(lbl)
            for v in art_vals:
                try:
                    articles.append(str(int(float(str(v)))))
                except Exception:
                    pass
        elif not is_label:
            # Строка сразу с артикулами (нет ярлыков)
            for _, v in label_vals:
                try:
                    articles.append(str(int(float(v))))
                    labels.append("")
                except Exception:
                    pass

        # Выравниваем длины
        max_len = max(len(labels), len(articles))
        labels += [""] * (max_len - len(labels))
        articles += [""] * (max_len - len(articles))

        return labels, articles

    def _make_record(self, nom, art, desc, price_ue, price_rub, group, variant):
        return {
            "Номенклатура": nom,
            "Артикул": art,
            "Описание": desc,
            "Вариант": variant,
            "Количество": "",
            "Цена розничная (у.е.)": price_ue,
            "Цена розничная (руб)": price_rub,
            "Ед. изм.": "шт.",
            "Валюта": "RUB",
            "Бренд": "SATA",
            "Vendorcode": art,
            "Full path": f"SATA\\{group}",
            "Подгруппа": group,
            "Фотографии": self._photo_url(art),
        }

    @staticmethod
    def _photo_url(art: str) -> str:
        if art:
            return f"https://lion-group.ru/images/SATA/{art}s.jpg"
        return ""

    @staticmethod
    def _is_numeric(val: str) -> bool:
        try:
            float(val.replace(",", "."))
            return True
        except Exception:
            return False

    @staticmethod
    def _to_float(v) -> float:
        try:
            return float(v)
        except Exception:
            return 0.0
