# -*- coding: utf-8 -*-
"""
04_shrinkflation_pairs.py — Поиск пар «шринкфляции» по брендам
==============================================================
Шринкфляция — когда товар «тихо» уменьшается: та же марка, меньшая
фасовка, часто новый код и слегка изменённое название. В данных это
выглядит как пара: старый товар ушёл → новый товар того же бренда
пришёл вскоре после, с меньшей (или другой) фасовкой.

Алгоритм:
  1) извлекаем бренд из названия (первая капс-группа и следующие капс-слова);
  2) для «ушедших» товаров — дата последней продажи,
     для «пришедших» — дата первой продажи;
  3) ищем пары одного бренда с разницей дат ≤ 90 дней.

Вход:  data/Молоко.xlsx
Выход: печать отчёта в консоль
"""

import re
from typing import Optional

import pandas as pd

df = pd.read_excel(
    "data/Молоко.xlsx",
    dtype={"Код номенклатуры": str},
    parse_dates=["День"],
)
df["Год"] = df["День"].dt.year
df["Месяц"] = df["День"].dt.month

# Фасовка в граммах (как в 03)
VOL_PATTERN = re.compile(r"(\d+(?:[.,]\d+)?)\s*(г|мл|л|кг)", re.IGNORECASE)


def parse_volume(name: str) -> Optional[float]:
    m = VOL_PATTERN.search(name)
    if not m:
        return None
    val = float(m.group(1).replace(",", "."))
    unit = m.group(2).lower()
    if unit == "мл":
        return val
    if unit == "л":
        return val * 1000
    if unit == "г":
        return val
    if unit == "кг":
        return val * 1000
    return None


df["Фасовка_г"] = df["Номенклатура"].map(parse_volume)


def brand_of(name: str) -> str:
    """Бренд — первая группа слов КАПСОМ (до первой строчной/цифры)."""
    m = re.match(r"\s*([А-ЯЁA-Z0-9][А-ЯЁA-Z0-9 .\-]*)", name)
    return m.group(1).strip() if m else name.split()[0]


df["Бренд"] = df["Номенклатура"].map(brand_of)

# Активность по каждому товару: первая и последняя дата, фасовка, бренд
info = (
    df.groupby("Код номенклатуры")
    .agg(
        Бренд=("Бренд", "first"),
        Название=("Номенклатура", "first"),
        Первая=("День", "min"),
        Последняя=("День", "max"),
        Фасовка_г=("Фасовка_г", "first"),
    )
    .reset_index()
)

# «Ушедшие»: последняя продажа раньше 2026-01-01, а в янв-июле 2026 их нет
active_2026h1 = set(
    df[(df["Год"] == 2026) & (df["Месяц"] <= 7)]["Код номенклатуры"]
)
gone = info[~info["Код номенклатуры"].isin(active_2026h1)].copy()
gone = gone[gone["Последняя"] < pd.Timestamp("2026-01-01")]

# «Пришедшие»: их не было в янв-июле 2022
active_2022h1 = set(
    df[(df["Год"] == 2022) & (df["Месяц"] <= 7)]["Код номенклатуры"]
)
arrived = info[~info["Код номенклатуры"].isin(active_2022h1)].copy()

print("=" * 70)
print("ПАРЫ «УШЁЛ → ПРИШЁЛ» ОДНОГО БРЕНДА (разница ≤ 90 дней)")
print("=" * 70)

pairs = []
for _, g in gone.iterrows():
    for _, a in arrived.iterrows():
        if g["Бренд"] != a["Бренд"]:
            continue
        gap = (a["Первая"] - g["Последняя"]).days
        if 0 <= gap <= 90:
            pairs.append((g, a, gap))

print(f"Найдено пар: {len(pairs)}\n")
for g, a, gap in sorted(pairs, key=lambda p: p[2]):
    vol_g = g["Фасовка_г"]
    vol_a = a["Фасовка_г"]
    tag = ""
    if vol_a and vol_g and vol_a < vol_g:
        tag = f"  ← ФАСОВКА УМЕНЬШЕНА: {vol_g:.0f}г → {vol_a:.0f}г"
    elif vol_a and vol_g and vol_a > vol_g:
        tag = f"  (фасовка увеличена: {vol_g:.0f}г → {vol_a:.0f}г)"
    print(f"Зазор {gap} дн: {g['Последняя'].date()} {g['Название'][:55]}")
    print(f"           {a['Первая'].date()} {a['Название'][:55]}{tag}\n")

print("=" * 70)
print("Поиск пар завершён.")
print("=" * 70)
