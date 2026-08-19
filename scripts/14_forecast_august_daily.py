# -*- coding: utf-8 -*-
"""
14_forecast_august_daily.py — Прогноз на каждый день августа 2026 по каждому товару
====================================================================================
Правило пользователя: продажи 1–3 августа 2026 (уже есть в данных) — это ФАКТ.
Прогноз строится на остальные дни месяца (4–31 августа), и факт участвует
в расчёте: по нему калибруется модель.

Метод (для каждой «истории» товара — с учётом шринкфляции):
  1) база  — средние продажи в день за май–июль 2026 (3 полных месяца);
     для совсем новых товаров (продаются только с августа) — среднее по 3 дням факта;
  2) профиль недели — коэффициенты по дням недели, посчитанные по магазину
     в целом (в какие дни недели молоко покупают больше/меньше);
  3) калибровка — фактические штуки 1–3 августа делим на прогноз модели
     на эти же дни: множитель применяем к 4–31 августа;
  4) выручка = штуки × средняя цена товара (май–июль, при нехватке — август-факт).

Ограничения и честность:
  - калибровка по 3 дням — грубая; множитель ограничен диапазоном [0; 3],
    чтобы один удачный день не «вздул» месяц;
  - если товар не продавался 1–3 августа (но продавался в мае–июле) —
    калибровки нет, прогноз «как в мае–июле» (статус «без калибровки»).

Выход: output/Прогноз_август_2026.xlsx + печать в консоль.
"""

import numpy as np
import pandas as pd

# ---------------------------------------------------------------
# Чтение и склейка в истории (как в 10)
# ---------------------------------------------------------------
df = pd.read_excel(
    "data/Молоко.xlsx",
    dtype={"Код номенклатуры": str},
    parse_dates=["День"],
)
pairs = pd.read_excel("output/пары_шринкфляции.xlsx", dtype=str)

chain_map = {}
for _, p in pairs.iterrows():
    chain_map[p["Код старой упаковки"]] = p["Код старой упаковки"]
    chain_map[p["Код новой упаковки"]] = p["Код старой упаковки"]
for code in df["Код номенклатуры"].unique():
    chain_map.setdefault(code, code)

df["История"] = df["Код номенклатуры"].map(chain_map)

chain_names = {}
for _, p in pairs.iterrows():
    chain_names[p["Код старой упаковки"]] = (
        f"{p['Название старой упаковки'][:38]} → {p['Название новой упаковки'][:38]}"
    )
first_name = df.groupby("Код номенклатуры")["Номенклатура"].first().to_dict()


def history_label(h: str) -> str:
    return chain_names.get(h, first_name[h])


# Цена за штуку считается на всём датасете — до срезов по периодам
df["Цена_шт"] = df["ТО, руб"] / df["ТО, в е.изм."]

# ---------------------------------------------------------------
# Периоды: обучение (май–июль), факт (1–3 авг), прогноз (4–31 авг)
# ---------------------------------------------------------------
TRAIN_0, TRAIN_1 = pd.Timestamp("2026-05-01"), pd.Timestamp("2026-07-31")
FACT_0, FACT_1 = pd.Timestamp("2026-08-01"), pd.Timestamp("2026-08-03")

train = df[(df["День"] >= TRAIN_0) & (df["День"] <= TRAIN_1)]
fact = df[(df["День"] >= FACT_0) & (df["День"] <= FACT_1)]

# Факт с разбивкой по дням (каждый день — свои цифры)
fact_daily_units = fact.groupby(["История", "День"])["ТО, в е.изм."].sum()
fact_daily_rub = fact.groupby(["История", "День"])["ТО, руб"].sum()
fact_units = fact.groupby("История")["ТО, в е.изм."].sum()
fact_rub = fact.groupby("История")["ТО, руб"].sum()

# Контроль факта (сверка с исходником)
q_fact = int(fact_units.sum())
r_fact = float(fact_rub.sum())
print("=" * 78)
print("ФАКТ 1–3 АВГУСТА 2026 (из данных):")
print(f"  Штук: {q_fact:,}  |  Выручка: {r_fact:,.2f} руб".replace(",", " "))

# ---------------------------------------------------------------
# Профиль недели: средние продажи магазина по дням недели (май–июль)
# ---------------------------------------------------------------
daily = train.groupby("День")["ТО, в е.изм."].sum()
wd_mean = daily.groupby(daily.index.dayofweek).mean()
wf = wd_mean / wd_mean.mean()  # 1 = средний день; >1 — «горячий» день недели
WD_RU = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]
print("\nПрофиль недели (май–июль 2026, 1 = средний день):")
print("  " + "  ".join(f"{WD_RU[i]} {wf.iloc[i]:.2f}" for i in range(7)))

# ---------------------------------------------------------------
# База и цена по товарам
# ---------------------------------------------------------------
days_train = (TRAIN_1 - TRAIN_0).days + 1  # 92 дня
base_units = train.groupby("История")["ТО, в е.изм."].sum() / days_train

price_train = train.groupby("История")["Цена_шт"].median()
price_fact = fact.groupby("История")["Цена_шт"].median()
price = price_train.combine_first(price_fact)

# Товары в расчёте: продавались в мае–июле ИЛИ в 1–3 августа
histories = sorted(set(base_units.index) | set(fact_units.index))

# ---------------------------------------------------------------
# Калибровка: факт 1–3 авг / прогноз модели на 1–3 авг
# ---------------------------------------------------------------
days13 = pd.date_range(FACT_0, FACT_1)
model13 = pd.Series(0.0, index=histories)
for d in days13:
    model13 = model13 + base_units.reindex(histories).fillna(0.0) * wf.loc[d.dayofweek]

# Без fillna(0.0): если товар не продавался 1–3 авг, деление даёт NaN →
# fillna(1.0) = «как в мае–июле» (статус «без калибровки»); fillna(0.0) обнуляло бы прогноз
ratio = (fact_units.reindex(histories) / model13.replace(0, np.nan)).fillna(1.0)
ratio = ratio.clip(0.0, 3.0)  # защита от «вздутия» одним удачным днём

# ---------------------------------------------------------------
# Прогноз по дням
# ---------------------------------------------------------------
days_aug = pd.date_range("2026-08-01", "2026-08-31")
status_of = {}
for h in histories:
    if h in fact_units.index and h in base_units.index:
        status_of[h] = "прогноз (калибровка фактом)"
    elif h in fact_units.index:
        status_of[h] = "прогноз (новинка, база — 3 дня факта)"
    else:
        status_of[h] = "прогноз (без калибровки)"

rows = []
for d in days_aug:
    is_fact = d <= FACT_1
    for h in histories:
        # база: обычная (май–июль) или для новинки — среднее по 3 факт-дням
        if h in base_units.index:
            b = base_units[h]
        elif h in fact_units.index:
            b = fact_units[h] / 3
        else:
            b = 0.0
        p = price.get(h, np.nan)
        if is_fact:
            u = float(fact_daily_units.get((h, d), 0.0))
            r = float(fact_daily_rub.get((h, d), 0.0))
            status = "факт"
        else:
            u = b * wf.loc[d.dayofweek] * ratio[h]
            r = u * p
            status = status_of[h]
        rows.append(
            {
                "Дата": d.date(),
                "День недели": WD_RU[d.dayofweek],
                "Тип": status,
                "Код истории": h,
                "Название товара (история)": history_label(h),
                "Прогноз, штук": round(u, 1),
                "Прогноз выручки, руб": round(r, 2),
                "База (шт/день, май–июл)": round(b, 2) if b else None,
                "Калибровка (множитель)": round(ratio[h], 2) if not is_fact else None,
                "Цена, руб/шт": round(p, 2) if pd.notna(p) else None,
                "Статус прогноза": status,
            }
        )

long_tbl = pd.DataFrame(rows)

# ---------------------------------------------------------------
# Сводки
# ---------------------------------------------------------------
# Товар × день (штуки): факт 1–3 + прогноз 4–31
pivot = long_tbl.pivot_table(
    index=["Код истории", "Название товара (история)"],
    columns=long_tbl["Дата"].apply(lambda d: d.day),
    values="Прогноз, штук",
    aggfunc="sum",
).fillna(0.0)
pivot.columns = [f"{c:02d}" for c in pivot.columns]
pivot["Итого за месяц"] = pivot.sum(axis=1).round(1)
pivot = pivot.sort_values("Итого за месяц", ascending=False).reset_index()
pivot["Статус прогноза"] = pivot["Код истории"].map(status_of)
pivot = pivot[
    ["Код истории", "Название товара (история)", "Статус прогноза"] + [f"{c:02d}" for c in range(1, 32)] + ["Итого за месяц"]
]

# Итоги по дням (весь магазин, одна строка на день)
daily_tot = (
    long_tbl.groupby(["Дата", "День недели"])
    .agg(Штук=("Прогноз, штук", "sum"), Выручка_руб=("Прогноз выручки, руб", "sum"))
    .reset_index()
)
daily_tot["Тип"] = daily_tot["Дата"].apply(lambda d: "факт" if d <= FACT_1.date() else "прогноз")
cnt_days = long_tbl[long_tbl["Прогноз, штук"] > 0].groupby("Дата")["Код истории"].nunique()
# map по значениям дат надёжнее reindex с Series (см. LESSONS.md)
daily_tot["Товаров"] = daily_tot["Дата"].map(cnt_days).fillna(0).astype(int)
daily_tot = daily_tot.rename(columns={"Штук": "Штук", "Выручка_руб": "Выручка, руб"})

# Сверка с помесячным прогнозом (скрипт 13)
aug_fc_13_r = 914728.0
aug_fc_13_q = 7959.0
aug_this_r = long_tbl["Прогноз выручки, руб"].sum()
aug_this_q = long_tbl["Прогноз, штук"].sum()
check = pd.DataFrame(
    {
        "Показатель": ["Выручка за август, руб", "Штук за август"],
        "Прогноз скрипта 13 (месяц)": [aug_fc_13_r, aug_fc_13_q],
        "Этот расчёт (факт + прогноз)": [round(aug_this_r, 2), round(aug_this_q, 1)],
        "Расхождение, %": [
            round(aug_this_r / aug_fc_13_r - 1, 3),
            round(aug_this_q / aug_fc_13_q - 1, 3),
        ],
    }
)

# ---------------------------------------------------------------
# Печать
# ---------------------------------------------------------------
print("\n" + "=" * 78)
print("ИТОГИ АВГУСТА 2026 (факт 1–3 + прогноз 4–31)")
print("=" * 78)
print(f"  Факт 1–3 августа:    {q_fact:>8,} шт  {r_fact:>12,.0f} руб".replace(",", " "))
fc_u = long_tbl[long_tbl["Тип"] != "факт"]["Прогноз, штук"].sum()
fc_r = long_tbl[long_tbl["Тип"] != "факт"]["Прогноз выручки, руб"].sum()
print(f"  Прогноз 4–31 августа: {fc_u:>8,.0f} шт  {fc_r:>12,.0f} руб".replace(",", " "))
print(f"  Всего за август:      {aug_this_q:>8,.0f} шт  {aug_this_r:>12,.0f} руб".replace(",", " "))
print(f"  Товаров в расчёте: {len(histories)}")
print("  По статусам прогноза:")
for st, n in pd.Series(status_of).value_counts().items():
    print(f"    {st}: {n}")

print("\nСверка со скриптом 13 (помесячный прогноз):")
print(check.to_string(index=False))

print("\nТоп-10 товаров по прогнозу на 4–31 августа (штук):")
sub = long_tbl[long_tbl["Тип"] != "факт"].groupby("Название товара (история)").agg(
    Штук=("Прогноз, штук", "sum"), Выручка=("Прогноз выручки, руб", "sum")
).sort_values("Штук", ascending=False)
for name, r in sub.head(10).iterrows():
    print(f"  {r['Штук']:>7,.0f} шт  {r['Выручка']:>10,.0f} руб  {name[:60]}".replace(",", " "))

print("\nПо дням (весь магазин):")
print(daily_tot.to_string(index=False))

# ---------------------------------------------------------------
# Сохранение книги
# ---------------------------------------------------------------
OUT = "output/Прогноз_август_2026.xlsx"
with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
    long_tbl.to_excel(writer, sheet_name="Август_по_дням", index=False)
    pivot.to_excel(writer, sheet_name="Товар_на_день_штук", index=False)
    daily_tot.to_excel(writer, sheet_name="Итоги_по_дням", index=False)
    check.to_excel(writer, sheet_name="Сверка_со_скриптом_13", index=False)

print(f"\nСохранено: {OUT}")
print("  Лист «Август_по_дням»:", len(long_tbl), "строк (товар × день)")
print("  Лист «Товар_на_день_штук»:", len(pivot), "товаров")
print("  Лист «Итоги_по_дням»:", len(daily_tot), "строк (дней)")
print("  Лист «Сверка_со_скриптом_13»: 2 строки")
