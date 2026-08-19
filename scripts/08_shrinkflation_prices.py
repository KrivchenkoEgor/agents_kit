# -*- coding: utf-8 -*-
"""
08_shrinkflation_prices.py — Цены до/после шринкфляции
=======================================================
Вопрос: почему продажи рухнули после смены фасовки (ЛЕБЕДЕВСКАЯ −86%,
МОЛОЧНАЯ СКАЗКА −96%)? Гипотеза: скрытое подорожание — цена за кг
выросла, покупатель заметил, что молока стало меньше (или дороже).

Для каждой склеенной пары (из output/пары_шринкфляции.xlsx) считаем:
  - «до»  — последние 180 дней старого кода перед переходом;
  - «после» — первые 180 дней нового кода.
Показатели: цена за упаковку и цена за КГ (медиана по записям
и средневзвешенная), число продаж.
"""

import pandas as pd

# ---------------------------------------------------------------
# Чтение
# ---------------------------------------------------------------
df = pd.read_excel(
    "data/Молоко.xlsx",
    dtype={"Код номенклатуры": str},
    parse_dates=["День"],
)
pairs = pd.read_excel("output/пары_шринкфляции.xlsx", dtype=str)

df["Цена_за_упаковку"] = df["ТО, руб"] / df["ТО, в е.изм."]

print("=" * 78)
print("ЦЕНЫ ДО/ПОСЛЕ СМЕНЫ ФАСОВКИ (медиана за 180 дней)")
print("=" * 78)

for _, p in pairs.iterrows():
    old_code, new_code = p["Код старой упаковки"], p["Код новой упаковки"]
    old_vol, new_vol = float(p["Фасовка старая, г/мл"]), float(p["Фасовка новая, г/мл"])
    switch = pd.Timestamp(p["Первая продажа новой упаковки"])

    d_old = df[(df["Код номенклатуры"] == old_code) & (df["День"] < switch)]
    d_old = d_old[d_old["День"] >= switch - pd.Timedelta(days=180)]
    d_new = df[(df["Код номенклатуры"] == new_code) & (df["День"] >= switch)]
    d_new = d_new[d_new["День"] <= switch + pd.Timedelta(days=180)]

    if d_old.empty or d_new.empty:
        print(f"\nПара {old_code} → {new_code}: недостаточно данных (старое {len(d_old)}, новое {len(d_new)})")
        continue

    def prices(d: pd.DataFrame, vol_g: float) -> dict:
        per_pack = d["Цена_за_упаковку"].median()
        per_kg = per_pack / vol_g * 1000
        w_kg = d["ТО, руб"].sum() / (d["ТО, в е.изм."].sum() * vol_g / 1000)
        return {
            "упаковка": per_pack,
            "за_кг_медиана": per_kg,
            "за_кг_взвеш": w_kg,
            "штук": int(d["ТО, в е.изм."].sum()),
            "дней": int(d["День"].nunique()),
        }

    old_p = prices(d_old, old_vol)
    new_p = prices(d_new, new_vol)

    print(f"\n{old_code} → {new_code}:  {p['Название старой упаковки'][:40]}")
    print(f"  смена фасовки {old_vol:.0f}г → {new_vol:.0f}г  ({new_vol / old_vol - 1:+.0%})  "
          f"переход {switch.date()}")
    print(f"  ДО    ({old_p['дней']} дн, {old_p['штук']:,} шт): упаковка {old_p['упаковка']:7.2f} руб, "
          f"за кг {old_p['за_кг_медиана']:7.2f} руб (взвеш. {old_p['за_кг_взвеш']:7.2f})".replace(",", " "))
    print(f"  ПОСЛЕ ({new_p['дней']} дн, {new_p['штук']:,} шт): упаковка {new_p['упаковка']:7.2f} руб, "
          f"за кг {new_p['за_кг_медиана']:7.2f} руб (взвеш. {new_p['за_кг_взвеш']:7.2f})".replace(",", " "))
    print(f"  Цена за кг изменилась: {new_p['за_кг_медиана'] / old_p['за_кг_медиана'] - 1:+.1%} "
          f"(медиана) / {new_p['за_кг_взвеш'] / old_p['за_кг_взвеш'] - 1:+.1%} (взвешенная)")

# ---------------------------------------------------------------
# Дополнительно: цена за кг по всему магазину по годам (контекст)
# ---------------------------------------------------------------
import re

VOL_PATTERN = re.compile(r"(\d+(?:[.,]\d+)?)\s*(г|мл|л|кг)", re.IGNORECASE)


def parse_volume(name: str) -> float:
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
df["Объём_кг"] = df["ТО, в е.изм."] * df["Фасовка_г"] / 1000
df["Год"] = df["День"].dt.year
df["Месяц"] = df["День"].dt.month

print("\n" + "=" * 78)
print("КОНТЕКСТ: цена за кг молока по магазину в целом (янв-июль каждого года)")
print("=" * 78)
for y in range(2022, 2027):
    d = df[(df["Год"] == y) & (df["Месяц"] <= 7)]
    if d.empty:
        continue
    per_kg = d["ТО, руб"].sum() / d["Объём_кг"].sum()
    print(f"  {y}: ≈ {per_kg:6.2f} руб/кг  (штук {d['ТО, в е.изм.'].sum():,})".replace(",", " "))
