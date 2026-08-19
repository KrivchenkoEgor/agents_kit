# -*- coding: utf-8 -*-
"""
09_background_prices.py — Подорожание при шринкфляции ОТНОСИТЕЛЬНО фона
======================================================================
Пользователь прав: в данных только РОЗНИЧНАЯ цена (ТО руб / штук),
наценка сети и цена поставщика не разделяются. Поэтому сравниваем
подорожание пары с ФОНОМ — ростом цен по всему магазину за те же
180 дней до/после перехода. Избыток = подорожание пары − фон.

Вход:  data/Молоко.xlsx, output/пары_шринкфляции.xlsx
Выход: печать в консоль
"""

import re

import pandas as pd

# ---------------------------------------------------------------
# Чтение и подготовка
# ---------------------------------------------------------------
df = pd.read_excel(
    "data/Молоко.xlsx",
    dtype={"Код номенклатуры": str},
    parse_dates=["День"],
)
pairs = pd.read_excel("output/пары_шринкфляции.xlsx", dtype=str)

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
df["Цена_за_кг"] = df["ТО, руб"] / df["Объём_кг"]

print("=" * 84)
print("ПОДОРОЖАНИЕ ПАРЫ vs ФОН МАГАЗИНА (180 дней до/после перехода)")
print("=" * 84)
print(f"{'Пара':<38}{'пара Δкг':>10}{'фон Δкг':>10}{'избыток':>10}")

results = []
for _, p in pairs.iterrows():
    old_code, new_code = p["Код старой упаковки"], p["Код новой упаковки"]
    switch = pd.Timestamp(p["Первая продажа новой упаковки"])
    window = pd.Timedelta(days=180)

    # Окна для пары
    d_old = df[(df["Код номенклатуры"] == old_code) & (df["День"] < switch) & (df["День"] >= switch - window)]
    d_new = df[(df["Код номенклатуры"] == new_code) & (df["День"] >= switch) & (df["День"] <= switch + window)]

    # Фон: весь магазин БЕЗ кодов пары
    codes_pair = {old_code, new_code}
    f_old = df[(df["День"] < switch) & (df["День"] >= switch - window) & (~df["Код номенклатуры"].isin(codes_pair))]
    f_new = df[(df["День"] >= switch) & (df["День"] <= switch + window) & (~df["Код номенклатуры"].isin(codes_pair))]

    if d_old.empty or d_new.empty or f_old.empty or f_new.empty:
        continue

    pair_kg_old = d_old["ТО, руб"].sum() / d_old["Объём_кг"].sum()
    pair_kg_new = d_new["ТО, руб"].sum() / d_new["Объём_кг"].sum()
    f_kg_old = f_old["ТО, руб"].sum() / f_old["Объём_кг"].sum()
    f_kg_new = f_new["ТО, руб"].sum() / f_new["Объём_кг"].sum()

    pair_delta = pair_kg_new / pair_kg_old - 1
    f_delta = f_kg_new / f_kg_old - 1
    excess = pair_delta - f_delta

    label = (p["Название старой упаковки"][:20] + " → " + p["Название новой упаковки"][:15]).replace("Молоко ", "")
    results.append((label, pair_delta, f_delta, excess, switch))

    print(f"{label:<38}{pair_delta:>+9.1%}{f_delta:>+10.1%}{excess:>+10.1%}")

print("\nПояснение: «пара Δкг» — как изменилась цена за кг у самого товара при")
print("смене фасовки; «фон Δкг» — как за те же 180 дней подорожал весь магазин;")
print("«избыток» — разница: на сколько переход подорожал БОЛЬШЕ фона.")
print("Если избыток положительный — смена фасовки несла доп. удорожание сверх")
print("общего роста цен; если около нуля — подорожание = продолжение общего тренда.")
