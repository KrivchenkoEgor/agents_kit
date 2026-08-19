# -*- coding: utf-8 -*-
"""
07_decompose_chains.py — Декомпозиция и отчёт С УЧЁТОМ шринкфляции
=================================================================
Правило пользователя: товар, сменивший фасовку на меньшую
(шринкфляция), считается ПРОДОЛЖЕНИЕМ продаж того же товара.
Пары из output/пары_шринкфляции.xlsx (скрипт 06) склеиваются в одну
«историю товара», и все расчёты ведутся по историям.

Сравнение — только одинаковые периоды: январь–июль 2022 vs 2026.

Выход: консольный отчёт + output/Молоко_анализ_с_учетом_шринкфляции.xlsx
"""

import pandas as pd

# ---------------------------------------------------------------
# Чтение данных и пар
# ---------------------------------------------------------------
df = pd.read_excel(
    "data/Молоко.xlsx",
    dtype={"Код номенклатуры": str},
    parse_dates=["День"],
)
pairs = pd.read_excel("output/пары_шринкфляции.xlsx", dtype=str)

# Сопоставление код → история (id = старый код пары, либо сам код)
chain_map = {}
for _, p in pairs.iterrows():
    old, new = p["Код старой упаковки"], p["Код новой упаковки"]
    chain_map[old] = old
    chain_map[new] = old
for code in df["Код номенклатуры"].unique():
    chain_map.setdefault(code, code)

df["История"] = df["Код номенклатуры"].map(chain_map)

# Названия истории: старое и новое (для склеенных)
chain_names = {}
for _, p in pairs.iterrows():
    old, new = p["Код старой упаковки"], p["Код новой упаковки"]
    chain_names[old] = (
        f"{p['Название старой упаковки']}  →  {p['Название новой упаковки']}  "
        f"(фасовка {p['Фасовка старая, г/мл']}г → {p['Фасовка новая, г/мл']}г)"
    )
first_name = df.groupby("Код номенклатуры")["Номенклатура"].first().to_dict()


def history_label(h: str) -> str:
    return chain_names.get(h, first_name[h])


df["Год"] = df["День"].dt.year
df["Месяц"] = df["День"].dt.month


def jan_jul(y: int) -> pd.DataFrame:
    return df[(df["Год"] == y) & (df["Месяц"] <= 7)]


a = jan_jul(2022)   # было
b = jan_jul(2026)   # стало

hist_a = set(a["История"])
hist_b = set(b["История"])

gone = hist_a - hist_b          # ушли насовсем (нет продолжения)
stayed = hist_a & hist_b        # живут в обоих периодах
arrived = hist_b - hist_a       # пришли к 2026

# Отделим «продолжившиеся через шринкфляцию» (склеенные пары)
chained = set(chain_map[code] for code in df["Код номенклатуры"] if code in chain_map and chain_map[code] != code)
shrink_stayed = stayed & chained
plain_stayed = stayed - chained

print("=" * 70)
print("1. ИТОГ ПО ПЕРИОДАМ (январь–июль), ПО ИСТОРИЯМ")
print("=" * 70)


def totals(d: pd.DataFrame) -> dict:
    return {
        "штук": int(d["ТО, в е.изм."].sum()),
        "выручка": float(d["ТО, руб"].sum()),
    }


ta, tb = totals(a), totals(b)
print(f"Историй в 2022: {len(hist_a)}, в 2026: {len(hist_b)}")
print(f"Штук: 2022 = {ta['штук']:,} → 2026 = {tb['штук']:,} "
      f"({tb['штук'] / ta['штук'] - 1:+.1%})".replace(",", " "))
print(f"Выручка: 2022 = {ta['выручка']:,.0f} → 2026 = {tb['выручка']:,.0f} "
      f"({tb['выручка'] / ta['выручка'] - 1:+.1%})".replace(",", " "))

print("\n" + "=" * 70)
print("2. РАЗЛОЖЕНИЕ ПО ГРУППАМ ИСТОРИЙ (штуки, янв-июль)")
print("=" * 70)

qa_gone = a[a["История"].isin(gone)]["ТО, в е.изм."].sum()
qa_shrink = a[a["История"].isin(shrink_stayed)]["ТО, в е.изм."].sum()
qb_shrink = b[b["История"].isin(shrink_stayed)]["ТО, в е.изм."].sum()
qa_plain = a[a["История"].isin(plain_stayed)]["ТО, в е.изм."].sum()
qb_plain = b[b["История"].isin(plain_stayed)]["ТО, в е.изм."].sum()
qb_arrived = b[b["История"].isin(arrived)]["ТО, в е.изм."].sum()

print(f"Ушли НАСОВСЕМ ({len(gone)} историй):      в 2022 было {qa_gone:>8,} шт  → потеря".replace(",", " "))
print(f"Шринкфляция ({len(shrink_stayed)} историй):      в 2022 {qa_shrink:>8,} → в 2026 {qb_shrink:>8,} шт".replace(",", " "))
print(f"Остались без смены ({len(plain_stayed)}):  в 2022 {qa_plain:>8,} → в 2026 {qb_plain:>8,} шт".replace(",", " "))
print(f"Пришли ({len(arrived)} историй):              в 2026 дают {qb_arrived:>8,} шт".replace(",", " "))

print("\nБаланс штук:")
print(f"  Было (2022):            {ta['штук']:>9,}".replace(",", " "))
print(f"  Ушло насовсем:          -{qa_gone:>9,}".replace(",", " "))
print(f"  Шринкфляция потеряла:   -{qa_shrink - qb_shrink:>9,}".replace(",", " "))
print(f"  Оставшиеся потеряли:    -{qa_plain - qb_plain:>9,}".replace(",", " "))
print(f"  Пришло с новыми:        +{qb_arrived:>9,}".replace(",", " "))
print(f"  Стало (2026):           {tb['штук']:>9,}".replace(",", " "))

print("\n" + "=" * 70)
print("3. ИСТОРИИ С ШРИНКФЛЯЦИЕЙ: БЫЛО → СТАЛО")
print("=" * 70)
for h in sorted(shrink_stayed):
    q_old = a[a["История"] == h]["ТО, в е.изм."].sum()
    q_new = b[b["История"] == h]["ТО, в е.изм."].sum()
    chg = (q_new / q_old - 1) if q_old else float("nan")
    print(f"  {q_old:>7,} → {q_new:>7,} шт ({chg:+.0%})  {history_label(h)[:90]}".replace(",", " "))

# ---------------------------------------------------------------
# Книга отчёта с учётом шринкфляции
# ---------------------------------------------------------------
print("\n" + "=" * 70)
print("4. СОЗДАНИЕ КНИГИ ОТЧЁТА")
print("=" * 70)

# Лист 1: свод по месяцам (как в 05)
monthly = (
    df.groupby(["Год", "Месяц"])
    .agg(
        Штук=("ТО, в е.изм.", "sum"),
        Выручка_руб=("ТО, руб", "sum"),
        Активных_историй=("История", "nunique"),
        Дней_с_данными=("День", "nunique"),
    )
    .reset_index()
)
total_days = df.groupby(["Год", "Месяц"])["День"].nunique()
monthly["Месяц_полный"] = monthly.apply(
    lambda r: total_days.get((r["Год"], r["Месяц"]), 0) >= 28, axis=1
)

# Лист 2: истории × месяцы
by_hist_month = (
    df.groupby(["История", "Год", "Месяц"])
    .agg(Штук=("ТО, в е.изм.", "sum"), Выручка_руб=("ТО, руб", "sum"))
    .reset_index()
    .sort_values(["История", "Год", "Месяц"])
)
by_hist_month["Название"] = by_hist_month["История"].map(history_label)

# Лист 3: истории одной строкой — продажи по годам + статус
pivot_q = df.pivot_table(index="История", columns="Год", values="ТО, в е.изм.", aggfunc="sum", fill_value=0)
pivot_rub = df.pivot_table(index="История", columns="Год", values="ТО, руб", aggfunc="sum", fill_value=0)

rows = []
for h in sorted(set(chain_map.values())):
    if h in gone:
        status = "ушла насовсем"
    elif h in shrink_stayed:
        status = "осталась (шринкфляция)"
    elif h in plain_stayed:
        status = "осталась"
    else:
        status = "пришла"
    r = {
        "История": h,
        "Название": history_label(h),
        "Статус": status,
        "Всего_шт": int(pivot_q.loc[h].sum()),
        "Всего_руб": round(float(pivot_rub.loc[h].sum()), 2),
    }
    for y in pivot_q.columns:
        r[f"{y}_шт"] = int(pivot_q.loc[h, y])
        r[f"{y}_руб"] = round(float(pivot_rub.loc[h, y]), 2)
    rows.append(r)

hist_by_year = pd.DataFrame(rows)
year_cols = []
for y in pivot_q.columns:
    year_cols += [f"{y}_шт", f"{y}_руб"]
cols = ["История", "Название", "Статус"] + year_cols + ["Всего_шт", "Всего_руб"]
hist_by_year = hist_by_year[cols]

# Лист 4: пары (из 06) + лист 5: справочник кодов с историей
catalog = (
    df.groupby("Код номенклатуры")
    .agg(
        Название=("Номенклатура", "first"),
        История=("История", "first"),
        Первый_день=("День", "min"),
        Последний_день=("День", "max"),
    )
    .reset_index()
)

# Контрольная сверка
assert abs(by_hist_month["Штук"].sum() - df["ТО, в е.изм."].sum()) < 1
assert abs(by_hist_month["Выручка_руб"].sum() - df["ТО, руб"].sum()) < 0.01

# Понятные заголовки колонок: сущность + единица измерения
R = {
    "История": "Код истории",
    "Название": "Название товара (история)",
    "Штук": "Продано, штук",
    "Выручка_руб": "Выручка, руб",
    "Активных_историй": "Товаров в продаже (историй)",
    "Дней_с_данными": "Дней с продажами",
    "Месяц_полный": "Месяц полный",
    "Статус": "Статус",
    "Всего_шт": "Продано всего, штук",
    "Всего_руб": "Выручка всего, руб",
    "Код номенклатуры": "Код товара",
    "Первый_день": "Первая продажа",
    "Последний_день": "Последняя продажа",
}
for y in pivot_q.columns:
    R[f"{y}_шт"] = f"Продано {y}, штук"
    R[f"{y}_руб"] = f"Выручка {y}, руб"

OUT = "output/Молоко_анализ_с_учетом_шринкфляции.xlsx"
with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
    monthly.rename(columns=R).to_excel(writer, sheet_name="Свод_по_месяцам", index=False)
    by_hist_month.rename(columns=R).to_excel(writer, sheet_name="Истории_помесячно", index=False)
    hist_by_year.rename(columns=R).to_excel(writer, sheet_name="Истории_по_годам", index=False)
    pairs.to_excel(writer, sheet_name="Пары_шринкфляции", index=False)
    catalog.rename(columns=R).to_excel(writer, sheet_name="Справочник_кодов", index=False)

print(f"Сохранено: {OUT}")
print("  Лист «Свод_по_месяцам»:", len(monthly), "месяцев")
print("  Лист «Истории_помесячно»:", len(by_hist_month), "строк (история × месяц)")
print("  Лист «Истории_по_годам»:", len(hist_by_year), "историй")
print("  Лист «Пары_шринкфляции»:", len(pairs), "пар")
print("  Лист «Справочник_кодов»:", len(catalog), "кодов")
print("  Контрольная сверка: штуки и выручка совпадают с исходником.")
