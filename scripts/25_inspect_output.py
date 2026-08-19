# -*- coding: utf-8 -*-
"""
25_inspect_output.py — обзор структуры всех книг в output/

Что делает: проходит по всем .xlsx в папке output/ и печатает по каждому
листу: название, число строк и колонок, имена колонок и первые 2 строки.

Зачем: перед сборкой дашборда нужно знать, какие показатели где лежат
и в каких единицах. Это «инвентаризация склада» перед открытием витрины.

Только чтение — ничего не создаёт и не меняет.
"""

import os
import pandas as pd

OUT = "output"  # папка с результатами (исходники не трогаем)

for fname in sorted(os.listdir(OUT)):
    if not fname.endswith(".xlsx"):
        continue  # docx и прочее пропускаем — это уже готовые отчёты

    path = os.path.join(OUT, fname)
    print("=" * 100)
    print("ФАЙЛ:", fname, f"({os.path.getsize(path) / 1024 / 1024:.1f} МБ)")

    try:
        xl = pd.ExcelFile(path)
    except Exception as e:
        print("  НЕ УДАЛОСЬ ОТКРЫТЬ:", e)
        continue

    for sheet in xl.sheet_names:
        try:
            # dtype=str — чтобы коды товаров с ведущими нулями не превращались в числа
            df = pd.read_excel(xl, sheet_name=sheet, dtype=str)
        except Exception as e:
            print(f"  ── Лист «{sheet}»: ошибка чтения: {e}")
            continue

        print(f"  ── Лист «{sheet}»: {len(df)} строк × {df.shape[1]} колонок")
        print("     Колонки:", ", ".join(str(c) for c in df.columns))
        for i in range(min(2, len(df))):
            cells = " | ".join(f"{c}={str(v)[:40]}" for c, v in df.iloc[i].items())
            print(f"     строка {i + 2}: {cells}")

    xl.close()

print("=" * 100)
print("ГОТОВО")
