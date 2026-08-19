# -*- coding: utf-8 -*-
"""
26_build_dashboard_data.py
==========================
Сбор данных для HTML-дашборда руководителя в единый JSON.

Источники (ТОЛЬКО чтение):
  - готовые отчёты из output/ (молоко и колбаса)
  - исходник data/колбасаНФ.xlsx (колбаса: месяцы, годы, группы_месяцы, итоги_всего)

Результат:
  - output/dashboard/dashboard_data.json  (ensure_ascii=False, indent=1)

Запуск: python3 scripts/26_build_dashboard_data.py  (из каталога проекта)
Python 3.8, pandas 2.0.3 (df.append удалён — используем pd.concat / groupby).
"""

import os
import json
import math
import re

import pandas as pd

# ---------------------------------------------------------------------------
# Пути
# ---------------------------------------------------------------------------
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "output")
DATA = os.path.join(BASE, "data")
DASH_DIR = os.path.join(OUT, "dashboard")
DASH_JSON = os.path.join(DASH_DIR, "dashboard_data.json")

# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------
def to_num(series):
    """Числовая колонка из строкового датафрейма -> float."""
    return pd.to_numeric(series, errors="coerce")


def r2(x):
    """Рубли: округление до 2 знаков."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    return round(float(x), 2)


def r1(x):
    """Кг / проценты: округление до 1 знака."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    return round(float(x), 1)


def r0(x):
    """Штуки: округление до целых."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    return int(round(float(x)))


def clean_code(s):
    """Код товара -> строка, zfill(5)."""
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return ""
    return str(s).strip().zfill(5)


def short_type(t):
    """Тип продажи -> короткий вид: 'штучный' / 'весовой'."""
    t = str(t).strip().lower()
    if "весов" in t:
        return "весовой"
    return "штучный"


def sold_value(x, is_weight):
    """
    Значение «продано» по типу товара:
    штучный -> целое число, весовой -> кг с 1 знаком.
    """
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return 0
    if is_weight:
        return r1(x)
    return r0(x)


FASOVKA_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*(кг|г|мл)", re.IGNORECASE)


def parse_fasovka(name):
    """
    Фасовка из названия товара -> граммы (int) или None.
    Ищем число + единицу (г/кг/мл), напр. «470г», «1кг», «350г».
    Если единица «кг» — умножаем на 1000. Если не нашли — None.
    """
    if name is None or (isinstance(name, float) and math.isnan(name)):
        return None
    m = FASOVKA_RE.search(str(name))
    if not m:
        return None
    num = float(m.group(1).replace(",", "."))
    unit = m.group(2).lower()
    if unit == "кг":
        return int(round(num * 1000))
    return int(round(num))


def continuation_name(s):
    """
    Из колонки «Продолжение в 2026 (код и название)» вытащить название:
    обрезаем всё до первой кавычки-ёлочки «, убираем хвостовую ».
    Для пустых значений -> пустая строка.
    """
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return ""
    s = str(s).strip()
    if "«" in s:
        s = s.split("«", 1)[1]
    if s.endswith("»"):
        s = s[:-1]
    return s.strip()


def check(name, value, etalon, tol=0.005, exact=False):
    """
    Печать контрольной строки: «КОНТРОЛЬ <имя>: <значение> | эталон <эталон> | OK/МИСМАТЧ».
    exact=True  -> строгое равенство
    иначе       -> допуск tol (по умолчанию 0.5%)
    """
    if exact:
        ok = (value == etalon)
    else:
        if etalon == 0:
            ok = (value == 0)
        else:
            ok = abs(float(value) - float(etalon)) / abs(float(etalon)) <= tol
    print("КОНТРОЛЬ {}: {} | эталон {} | {}".format(name, value, etalon, "OK" if ok else "МИСМАТЧ"))
    return ok


# ---------------------------------------------------------------------------
# 1. МОЛОКО
# ---------------------------------------------------------------------------
moloko = {}

# --- 1.1 месяцы ------------------------------------------------------------
m = pd.read_excel(os.path.join(OUT, "Молоко_анализ_по_месяцам.xlsx"),
                  sheet_name="Свод_по_месяцам", dtype=str)
m_months = []
for _, row in m.iterrows():
    m_months.append({
        "месяц": "{:04d}-{:02d}".format(int(row["Год"]), int(row["Месяц"])),
        "штук": r0(to_num(pd.Series([row["Продано, штук"]]))[0]),
        "руб": r2(to_num(pd.Series([row["Выручка, руб"]]))[0]),
        "кг": r1(to_num(pd.Series([row["Продано, кг"]]))[0]),
        "товаров": int(row["Товаров в продаже"]),
        "полный": str(row["Месяц полный"]).strip().lower() == "true",
    })
moloko["месяцы"] = m_months

# --- 1.2 abc и топ10 -------------------------------------------------------
abc = pd.read_excel(os.path.join(OUT, "ABC_анализ.xlsx"),
                    sheet_name="ABC_по_выручке", dtype=str)
abc_list = []
for _, row in abc.iterrows():
    abc_list.append({
        "код": clean_code(row["Код истории"]),
        "название": str(row["Название товара (история)"]).strip(),
        "выручка12": r2(to_num(pd.Series([row["Выручка за 12 мес, руб"]]))[0]),
        "доля": r1(to_num(pd.Series([row["Доля выручки, %"]]))[0] * 100.0),
        "группа": str(row["Группа ABC"]).strip(),
        "штук12": r0(to_num(pd.Series([row["Продано за 12 мес, штук"]]))[0]),
        "сравнение2022": str(row["Группа сейчас (сравнение с 2022)"]).strip(),
    })
moloko["abc"] = abc_list
moloko["топ10"] = [
    {"код": d["код"], "название": d["название"],
     "выручка12": d["выручка12"], "штук12": d["штук12"]}
    for d in abc_list[:10]
]

# --- 1.3 новинки -----------------------------------------------------------
nv = pd.read_excel(os.path.join(OUT, "Новинки_судьба.xlsx"),
                   sheet_name="Новинки_по_годам", dtype=str)
novinki = []
for _, row in nv.iterrows():
    novinki.append({
        "год": int(row["Год появления"]),
        "новинок": int(row["Новинок, шт"]),
        "живы": int(row["Продаются сейчас, шт"]),
        "звёздыA": int(row["Стали звёздами (A), шт"]),
        "выручка12": r2(to_num(pd.Series([row["Выручка за 12 мес, руб"]]))[0]),
        "доля_живых": r1(to_num(pd.Series([row["Доля живых, %"]]))[0]),
    })
moloko["новинки"] = novinki

# --- 1.4 прогноз -----------------------------------------------------------
pr = pd.read_excel(os.path.join(OUT, "Прогноз_2026.xlsx"),
                   sheet_name="Прогноз_авг_дек", dtype=str)
prognoz = []
for _, row in pr.iterrows():
    prognoz.append({
        "месяц": str(row["Месяц"]).strip(),
        "руб": r0(to_num(pd.Series([row["Прогноз выручки, руб"]]))[0]),
        "низ": r0(to_num(pd.Series([row["Нижняя граница (95%), руб"]]))[0]),
        "верх": r0(to_num(pd.Series([row["Верхняя граница (95%), руб"]]))[0]),
        "штук": r0(to_num(pd.Series([row["Прогноз штук"]]))[0]),
        "низШт": r0(to_num(pd.Series([row["Нижняя граница штук"]]))[0]),
        "верхШт": r0(to_num(pd.Series([row["Верхняя граница штук"]]))[0]),
    })
moloko["прогноз"] = prognoz

# --- 1.5 сверка прогноза (4–9 августа) -------------------------------------
sv = pd.read_excel(os.path.join(OUT, "Отчет_расхождения_прогноз_факт.xlsx"),
                   sheet_name="Итоги_по_дням", dtype=str)
sv_ps = to_num(sv["Прогноз, штук"]).sum()
sv_fs = to_num(sv["Факт, штук"]).sum()
sv_pr = to_num(sv["Прогноз, руб"]).sum()
sv_fr = to_num(sv["Факт, руб"]).sum()
dates = [str(d)[:10] for d in sv["Дата"]]
moloko["сверка_прогноза"] = {
    "период": "{} … {}".format(dates[0], dates[-1]),
    "прогноз_штук": round(sv_ps, 1),
    "факт_штук": r0(sv_fs),
    "откл_штук_проц": r1((sv_fs - sv_ps) / sv_ps * 100.0),
    "прогноз_руб": r2(sv_pr),
    "факт_руб": r2(sv_fr),
    "откл_руб_проц": r1((sv_fr - sv_pr) / sv_pr * 100.0),
}

# --- 1.6 шринкфляция -------------------------------------------------------
sh = pd.read_excel(os.path.join(OUT, "пары_шринкфляции.xlsx"),
                   sheet_name="Пары_шринкфляции", dtype=str)
shrink = []
for _, row in sh.iterrows():
    shrink.append({
        "код_старый": clean_code(row["Код старой упаковки"]),
        "название": str(row["Название старой упаковки"]).strip(),
        "фасовка_старая": int(to_num(pd.Series([row["Фасовка старая, г/мл"]]))[0]),
        "фасовка_новая": int(to_num(pd.Series([row["Фасовка новая, г/мл"]]))[0]),
        "зазор": int(to_num(pd.Series([row["Зазор, дней"]]))[0]),
    })
moloko["шринкфляция"] = shrink

# --- 1.7 группа C: статусы -------------------------------------------------
gc = pd.read_excel(os.path.join(OUT, "Группа_C_анализ.xlsx"),
                   sheet_name="Итоги_по_статусам", dtype=str)
gc_list = []
for _, row in gc.iterrows():
    gc_list.append({
        "статус": str(row["Статус"]).strip(),
        "товаров": int(row["Товаров, шт"]),
        "выручка12": r2(to_num(pd.Series([row["Выручка за 12 мес, руб"]]))[0]),
        "доля_товаров": r1(to_num(pd.Series([row["Доля товаров группы C, %"]]))[0]),
        "доля_выручки": r1(to_num(pd.Series([row["Доля выручки группы C, %"]]))[0]),
    })
moloko["группаC_статусы"] = gc_list

# --- 1.8 истории (ушла насовсем / осталась) --------------------------------
ist = pd.read_excel(os.path.join(OUT, "Молоко_анализ_с_учетом_шринкфляции.xlsx"),
                    sheet_name="Истории_по_годам", dtype=str)
ist_sub = ist[ist["Статус"].isin(["ушла насовсем", "осталась"])]
istorii = []
for _, row in ist_sub.iterrows():
    istorii.append({
        "код": clean_code(row["Код истории"]),
        "название": str(row["Название товара (история)"]).strip(),
        "статус": str(row["Статус"]).strip(),
        "выручка2026": r2(to_num(pd.Series([row["Выручка 2026, руб"]]))[0]),
    })
moloko["истории"] = istorii

# ---------------------------------------------------------------------------
# 2. КОЛБАСА — из исходника data/колбасаНФ.xlsx
# ---------------------------------------------------------------------------
kolb = {}

src = pd.read_excel(os.path.join(DATA, "колбасаНФ.xlsx"), dtype=str)
# колонка кода называется «Код » (с пробелом на конце)
src["Код "] = src["Код "].str.strip().str.zfill(5)
src["День"] = pd.to_datetime(src["День"], errors="coerce")
src["ТО, руб"] = to_num(src["ТО, руб"])
src["Наценка"] = to_num(src["Наценка"])
src["ТО, в е.изм."] = to_num(src["ТО, в е.изм."])

# Тип товара по правилу пользователя:
# весовой, если есть ХОТЯ БЫ ОДНА дробная продажа; иначе штучный.
frac_by_code = src.groupby("Код ")["ТО, в е.изм."].apply(
    lambda s: (s != s.round()).any()
)
type_map = {c: ("весовой" if frac_by_code[c] else "штучный") for c in frac_by_code.index}
src["тип"] = src["Код "].map(type_map)

src["год"] = src["День"].dt.year
src["месяц"] = src["День"].dt.strftime("%Y-%m")

# --- 2.1 месяцы ------------------------------------------------------------
m_agg = src.groupby("месяц").agg(
    руб=("ТО, руб", "sum"),
    прибыль=("Наценка", "sum"),
)
kg_agg = src.loc[src["тип"] == "весовой"].groupby("месяц")["ТО, в е.изм."].sum()
sht_agg = src.loc[src["тип"] == "штучный"].groupby("месяц")["ТО, в е.изм."].sum()

k_months = []
for month in sorted(m_agg.index):
    k_months.append({
        "месяц": month,
        "руб": r2(m_agg.loc[month, "руб"]),
        "прибыль": r2(m_agg.loc[month, "прибыль"]),
        "штук": r0(sht_agg.get(month, 0)),
        "кг": r1(kg_agg.get(month, 0)),
        "полный": month != "2026-08",
    })
kolb["месяцы"] = k_months

# --- 2.2 годы --------------------------------------------------------------
y_agg = src.groupby("год").agg(
    руб=("ТО, руб", "sum"),
    прибыль=("Наценка", "sum"),
)
y_kg = src.loc[src["тип"] == "весовой"].groupby("год")["ТО, в е.изм."].sum()
y_sht = src.loc[src["тип"] == "штучный"].groupby("год")["ТО, в е.изм."].sum()

k_years = []
for year in sorted(y_agg.index):
    k_years.append({
        "год": int(year),
        "руб": r2(y_agg.loc[year, "руб"]),
        "прибыль": r2(y_agg.loc[year, "прибыль"]),
        "штук": r0(y_sht.get(year, 0)),
        "кг": r1(y_kg.get(year, 0)),
        "полный": int(year) != 2026,
    })
kolb["годы"] = k_years

# --- 2.3 группы_месяцы -----------------------------------------------------
gm = src.groupby(["месяц", "группа"])["ТО, руб"].sum().reset_index()
kolb["группы_месяцы"] = [
    {"месяц": row["месяц"], "группа": str(row["группа"]).strip(), "руб": r2(row["ТО, руб"])}
    for _, row in gm.iterrows()
]

# --- 2.4 итоги_всего -------------------------------------------------------
kolb["итоги_всего"] = {
    "строк": int(len(src)),
    "товаров": int(src["Код "].nunique()),
    "руб": r2(src["ТО, руб"].sum()),
    "прибыль": r2(src["Наценка"].sum()),
    "штук": r0(src.loc[src["тип"] == "штучный", "ТО, в е.изм."].sum()),
    "кг": r1(src.loc[src["тип"] == "весовой", "ТО, в е.изм."].sum()),
}

# --- 2.5 топ5_по_годам -----------------------------------------------------
# Для каждого года (2023..2026) и каждой группы — 5 товаров с наибольшей
# выручкой «ТО, руб» за этот год, отсортированные по убыванию (место 1..5).
# «тип» переиспользуем из type_map (правило пользователя: есть хоть одна
# дробная продажа -> весовой). «продано» — в собственных единицах товара
# (весовой: кг с 1 знаком; штучный: целые шт). «продано_кг» — для весовых
# равно «продано», для штучных = «продано» × фасовка/1000 (фасовка из
# названия регуляркой; если не извлеклась — null).
# Важно: как и в отчётах script 21 (эталон сверки), учитываем только
# положительные продажи (возвраты с отрицательной суммой исключаем).
YEARS_TOP5 = [2023, 2024, 2025, 2026]

pos_src = src[(src["ТО, в е.изм."] > 0) & (src["ТО, руб"] > 0)]

top5_by_year = {}
for year in YEARS_TOP5:
    sub = pos_src[pos_src["год"] == year]
    year_rows = []
    for grp in sorted(sub["группа"].unique()):
        g = sub[sub["группа"] == grp]
        agg = g.groupby("Код ").agg(
            выручка=("ТО, руб", "sum"),
            продано=("ТО, в е.изм.", "sum"),
            название=("Номенклатура", "first"),
            тип=("тип", "first"),
        ).sort_values("выручка", ascending=False).head(5)
        for place, (code, r) in enumerate(agg.iterrows(), 1):
            is_w = r["тип"] == "весовой"
            sold = sold_value(r["продано"], is_w)
            if is_w:
                sold_kg = r1(r["продано"])
            else:
                fas = parse_fasovka(r["название"])
                if fas is None:
                    sold_kg = None
                else:
                    sold_kg = r1(sold * fas / 1000.0)
            year_rows.append({
                "группа": str(grp).strip(),
                "место": place,
                "код": code,
                "название": str(r["название"]).strip(),
                "тип": r["тип"],
                "выручка": r2(r["выручка"]),
                "продано": sold,
                "продано_кг": sold_kg,
            })
    top5_by_year[str(year)] = {
        "полный": year != 2026,
        "строки": year_rows,
    }
kolb["топ5_по_годам"] = top5_by_year

# ---------------------------------------------------------------------------
# 3. КОЛБАСА — из готовых отчётов
# ---------------------------------------------------------------------------
# --- 3.1 группы_итог (Сводка_по_группам из отчёта топ-5) -------------------
g5 = pd.read_excel(os.path.join(OUT, "Отчет_топ5_колбаса_по_группам.xlsx"),
                   sheet_name="Сводка_по_группам", dtype=str)
g5 = g5[g5["Группа"] != "ИТОГО по магазину"]
gruppy_itog = []
for _, row in g5.iterrows():
    gruppy_itog.append({
        "группа": str(row["Группа"]).strip(),
        "товаров": int(row["Товаров"]),
        "выручка": r2(to_num(pd.Series([row["Выручка за всё время, руб"]]))[0]),
        "кг": r1(to_num(pd.Series([row["Продано весовых, кг"]]))[0]),
        "шт": r0(to_num(pd.Series([row["Продано штучных, шт"]]))[0]),
        "доля": r1(to_num(pd.Series([row["Доля в выручке магазина, %"]]))[0]),
        "лидер_кол": str(row["Лидер по количеству"]).strip(),
        "лидер_выр": str(row["Лидер по выручке"]).strip(),
    })
kolb["группы_итог"] = gruppy_itog

# --- 3.2 прибыль_топ / прибыль_группы / прибыль_теряем ---------------------
pf = pd.read_excel(os.path.join(OUT, "Отчет_топ_прибыли_колбаса.xlsx"),
                   sheet_name="Топ_по_валовой_прибыли", dtype=str)
pf = pf.head(20)  # первые 20 строк; лишние Unnamed: 11/12 игнорируем
pribyl_top = []
for _, row in pf.iterrows():
    is_w = short_type(row["Тип продажи"]) == "весовой"
    pribyl_top.append({
        "место": int(row["Место"]),
        "группа": str(row["Группа"]).strip(),
        "код": clean_code(row["Код"]),
        "название": str(row["Название товара"]).strip(),
        "тип": short_type(row["Тип продажи"]),
        "продано": sold_value(to_num(pd.Series([row["Продано за всё время (шт/кг)"]]))[0], is_w),
        "продано_кг": r1(to_num(pd.Series([row["Продано ≈ кг"]]))[0]),
        "выручка": r2(to_num(pd.Series([row["Выручка, руб"]]))[0]),
        "прибыль": r2(to_num(pd.Series([row["Валовая прибыль, руб"]]))[0]),
        "наценка": r1(to_num(pd.Series([row["Наценка, % от выручки"]]))[0]),
        "в2026": str(row["В продаже 2026"]).strip(),
    })
kolb["прибыль_топ"] = pribyl_top

pg = pd.read_excel(os.path.join(OUT, "Отчет_топ_прибыли_колбаса.xlsx"),
                   sheet_name="Сводка_по_группам", dtype=str)
pribyl_gruppy = []
for _, row in pg.iterrows():
    pribyl_gruppy.append({
        "группа": str(row["Группа"]).strip(),
        "товаров": int(row["Товаров за период"]),
        "выручка": r2(to_num(pd.Series([row["Выручка, руб"]]))[0]),
        "прибыль": r2(to_num(pd.Series([row["Валовая прибыль, руб"]]))[0]),
        "доля_прибыли": r1(to_num(pd.Series([row["Доля в прибыли, %"]]))[0]),
        "наценка": r1(to_num(pd.Series([row["Наценка, % от выручки"]]))[0]),
    })
kolb["прибыль_группы"] = pribyl_gruppy

pt = pd.read_excel(os.path.join(OUT, "Отчет_топ_прибыли_колбаса.xlsx"),
                   sheet_name="Где_теряем", dtype=str)
pt = pt.head(5)
pribyl_teryaem = []
for _, row in pt.iterrows():
    pribyl_teryaem.append({
        "место": int(row["Место"]),
        "группа": str(row["Группа"]).strip(),
        "код": clean_code(row["Код"]),
        "название": str(row["Название товара"]).strip(),
        "прибыль": r2(to_num(pd.Series([row["Валовая прибыль, руб"]]))[0]),
        "наценка": r1(to_num(pd.Series([row["Наценка, % от выручки"]]))[0]),
        "в2026": str(row["В продаже 2026"]).strip(),
    })
kolb["прибыль_теряем"] = pribyl_teryaem

# --- 3.3 abc (Отчет_ABC_колбаса_группы.xlsx) -------------------------------
abc_v = pd.read_excel(os.path.join(OUT, "Отчет_ABC_колбаса_группы.xlsx"),
                      sheet_name="ABC_по_выручке_руб", dtype=str)
abc_o = pd.read_excel(os.path.join(OUT, "Отчет_ABC_колбаса_группы.xlsx"),
                      sheet_name="ABC_по_объёму_шт_кг", dtype=str)
abc_n = pd.read_excel(os.path.join(OUT, "Отчет_ABC_колбаса_группы.xlsx"),
                      sheet_name="ABC_по_наценке_руб", dtype=str)
abc_all = pd.read_excel(os.path.join(OUT, "Отчет_ABC_колбаса_группы.xlsx"),
                        sheet_name="ABC_общий_по_магазину", dtype=str)

KL = "Класс (A=первые 80%, B=до 95%, C=хвост)"
abc_v["Код"] = abc_v["Код"].str.strip().str.zfill(5)
abc_o["Код"] = abc_o["Код"].str.strip().str.zfill(5)
abc_n["Код"] = abc_n["Код"].str.strip().str.zfill(5)
abc_all["Код"] = abc_all["Код"].str.strip().str.zfill(5)

setA_v = set(abc_v.loc[abc_v[KL] == "A", "Код"])
setA_o = set(abc_o.loc[abc_o[KL] == "A", "Код"])
setA_n = set(abc_n.loc[abc_n[KL] == "A", "Код"])
setC_v = set(abc_v.loc[abc_v[KL] == "C", "Код"])
setC_n = set(abc_n.loc[abc_n[KL] == "C", "Код"])

stars = setA_v & setA_o & setA_n
outsiders = setC_v & setC_n

poolA = abc_all[abc_all["Класс по выручке (общий)"] == "A"]
vyr_all = to_num(abc_all["Товарооборот за год, руб"])
vyr_poolA = to_num(poolA["Товарооборот за год, руб"])

# по группам: число товаров и распределение A/B/C по выручке
po_gruppam = []
for grp, sub in abc_v.groupby("Группа"):
    po_gruppam.append({
        "группа": str(grp).strip(),
        "товаров": int(len(sub)),
        "a": int((sub[KL] == "A").sum()),
        "b": int((sub[KL] == "B").sum()),
        "c": int((sub[KL] == "C").sum()),
    })

kolb["abc"] = {
    "период": "2025-08-11 … 2026-08-10",
    "товаров": int(len(abc_all)),
    "весовых": int((abc_all["Тип продажи"] == "весовой, кг").sum()),
    "штучных": int((abc_all["Тип продажи"] == "штучный, шт").sum()),
    "выручка_год": r2(vyr_all.sum()),
    "пул_A": {
        "товаров": int(len(poolA)),
        "доля_выручки": r1(vyr_poolA.sum() / vyr_all.sum() * 100.0),
    },
    "звёзды": int(len(stars)),
    "аутсайдеры": int(len(outsiders)),
    "по_группам": po_gruppam,
}

# --- 3.4 потерянные_2023 / потерянные_2024 ---------------------------------
def parse_lost(path, year_col, sold_col, rev_col, sold26_col):
    df = pd.read_excel(path, sheet_name="Потерянные_топы", dtype=str)
    out = []
    for _, row in df.iterrows():
        status = str(row["Статус"]).strip()
        if "шринкфляция" in status:
            status_short = "шринкфляция"
        else:
            status_short = "исчез"
        is_w = short_type(row["Тип продажи"]) == "весовой"
        sold26 = to_num(pd.Series([row[sold26_col]]))[0]
        if sold26 is None or math.isnan(sold26):
            sold26 = 0
        out.append({
            "группа": str(row["Группа"]).strip(),
            "код": clean_code(row["Код"]),
            "название": str(row["Название товара"]).strip(),
            "продано" + year_col: sold_value(to_num(pd.Series([row[sold_col]]))[0], is_w),
            "выручка" + year_col: r2(to_num(pd.Series([row[rev_col]]))[0]),
            "статус": status_short,
            "продолжение": continuation_name(row["Продолжение в 2026 (код и название)"]),
            "продано2026": sold_value(sold26, is_w),
        })
    return out

kolb["потерянные_2023"] = parse_lost(
    os.path.join(OUT, "Отчет_потерянные_топы_2023_2026_v2.xlsx"),
    "2023", "Продано в 2023 (шт/кг)", "Выручка в 2023, руб",
    "Продано в 2026 (продолжение, шт/кг)")
kolb["потерянные_2024"] = parse_lost(
    os.path.join(OUT, "Отчет_потерянные_топы_2024_2026_v2.xlsx"),
    "2024", "Продано в 2024 (шт/кг)", "Выручка в 2024, руб",
    "Продано в 2026 (продолжение, шт/кг)")

# --- 3.5 топ5_выручка ------------------------------------------------------
t5 = pd.read_excel(os.path.join(OUT, "Отчет_топ5_колбаса_по_группам.xlsx"),
                   sheet_name="Топ-5_по_выручке", dtype=str)
top5 = []
for _, row in t5.iterrows():
    is_w = short_type(row["Тип продажи"]) == "весовой"
    top5.append({
        "группа": str(row["Группа"]).strip(),
        "место": int(row["Место в группе"]),
        "код": clean_code(row["Код"]),
        "название": str(row["Название товара"]).strip(),
        "тип": str(row["Тип продажи"]).strip(),
        "выручка": r2(to_num(pd.Series([row["Выручка за всё время, руб"]]))[0]),
        "продано": sold_value(to_num(pd.Series([row["Продано за всё время, шт/кг"]]))[0], is_w),
        "продано_кг": r1(to_num(pd.Series([row["Продано ≈ кг"]]))[0]),
    })
kolb["топ5_выручка"] = top5

# ---------------------------------------------------------------------------
# 4. СБОРКА JSON
# ---------------------------------------------------------------------------
dashboard = {
    "meta": {
        "собрано": "2026-08-19",
        "магазин": "один магазин",
        "периоды": {
            "молоко": "2022-01-01 … 2026-08-03",
            "колбаса": "2023-01-01 … 2026-08-10",
        },
    },
    "молоко": moloko,
    "колбаса": kolb,
}

os.makedirs(DASH_DIR, exist_ok=True)
with open(DASH_JSON, "w", encoding="utf-8") as f:
    json.dump(dashboard, f, ensure_ascii=False, indent=1)

# ---------------------------------------------------------------------------
# 5. КОНТРОЛЬНЫЕ СУММЫ
# ---------------------------------------------------------------------------
print("=" * 70)
print("КОНТРОЛЬНЫЕ СУММЫ")
print("=" * 70)

# --- Молоко ---
m_sht = sum(d["штук"] for d in m_months)
m_rub = round(sum(d["руб"] for d in m_months), 2)
check("молоко_месяцев_штук", m_sht, 857675, exact=True)
check("молоко_месяцев_выручка", m_rub, 71427373.81, exact=True)
check("молоко_месяцев_кол-во", len(m_months), 56, exact=True)
check("молоко_последний_месяц_полный", m_months[-1]["полный"], False, exact=True)

top1 = abc_list[0]
check("молоко_топ1_код", top1["код"], "93235", exact=True)
check("молоко_топ1_выручка", top1["выручка12"], 1255686.26, exact=True)

aug = prognoz[0]
check("молоко_прогноз_авг_руб", aug["руб"], 914728, exact=True)
check("молоко_прогноз_авг_штук", aug["штук"], 7959, exact=True)

svk = moloko["сверка_прогноза"]
check("сверка_факт_штук", svk["факт_штук"], 1689, exact=True)
check("сверка_откл_штук_проц", svk["откл_штук_проц"], 0.9, tol=0.05)

# --- Колбаса (исходник) ---
it = kolb["итоги_всего"]
check("колбаса_руб_всего", it["руб"], 1065900000, tol=0.005)
check("колбаса_прибыль_всего", it["прибыль"], 237200000, tol=0.005)
check("колбаса_кг_всего", it["кг"], 410761, tol=0.005)
check("колбаса_шт_всего", it["штук"], 5036968, tol=0.005)
check("колбаса_строк", it["строк"], 246300, tol=0.005)

y_by_year = {d["год"]: d for d in k_years}
check("колбаса_2023_руб", y_by_year[2023]["руб"], 320800000, tol=0.005)
check("колбаса_2024_руб", y_by_year[2024]["руб"], 321300000, tol=0.005)
check("колбаса_2025_руб", y_by_year[2025]["руб"], 280400000, tol=0.005)
check("колбаса_2026_руб", y_by_year[2026]["руб"], 143400000, tol=0.005)

# --- Колбаса (отчёты) ---
g_sum = round(sum(d["выручка"] for d in gruppy_itog), 2)
check("группы_итог_сумма_выручки", g_sum, 1065900000, tol=0.005)
check("группы_итог_вареная", gruppy_itog[0]["выручка"], 297158637.08, exact=True)

p_sum = round(sum(d["прибыль"] for d in pribyl_gruppy), 2)
check("прибыль_сумма_по_группам", p_sum, 237200000, tol=0.005)
check("прибыль_топ1", pribyl_top[0]["прибыль"], 12308814.86, exact=True)

abc_k = kolb["abc"]
check("abc_пул_A_товаров", abc_k["пул_A"]["товаров"], 89, exact=True)
check("abc_пул_A_доля", abc_k["пул_A"]["доля_выручки"], 79.8, tol=0.01)
check("abc_выручка_год", abc_k["выручка_год"], 248300000, tol=0.005)
check("abc_звёзды", abc_k["звёзды"], 70, exact=True)
check("abc_аутсайдеры", abc_k["аутсайдеры"], 112, exact=True)

lost23 = kolb["потерянные_2023"]
lost24 = kolb["потерянные_2024"]
check("потерянные_2023_строк", len(lost23), 18, exact=True)
check("потерянные_2023_шринкфляция",
      sum(1 for d in lost23 if d["статус"] == "шринкфляция"), 10, exact=True)
check("потерянные_2023_исчез",
      sum(1 for d in lost23 if d["статус"] == "исчез"), 8, exact=True)
check("потерянные_2024_строк", len(lost24), 10, exact=True)
check("потерянные_2024_шринкфляция",
      sum(1 for d in lost24 if d["статус"] == "шринкфляция"), 3, exact=True)
check("потерянные_2024_исчез",
      sum(1 for d in lost24 if d["статус"] == "исчез"), 7, exact=True)

check("топ5_выручка_строк", len(top5), 35, exact=True)

# --- топ5_по_годам: сверки -------------------------------------------------
t5y = kolb["топ5_по_годам"]

# a) в каждом году ровно 35 строк (7 групп × 5)
for year in YEARS_TOP5:
    check("топ5_по_годам_{}_строк".format(year),
          len(t5y[str(year)]["строки"]), 35, exact=True)

# b) сверка 2023 и 2024 с готовыми проверенными отчётами
for year, ref_file, ref_sheet, ref_col in [
    (2023, "Отчет_потерянные_топы_2023_2026_v2.xlsx",
     "Топ-10_по_выручке_2023", "Выручка в 2023, руб"),
    (2024, "Отчет_потерянные_топы_2024_2026_v2.xlsx",
     "Топ-10_по_выручке_2024", "Выручка в 2024, руб"),
]:
    ref = pd.read_excel(os.path.join(OUT, ref_file), sheet_name=ref_sheet, dtype=str)
    my_rows = t5y[str(year)]["строки"]
    my_map = {(r["группа"], r["место"]): (r["код"], r["выручка"]) for r in my_rows}
    mism = 0
    for _, r in ref.iterrows():
        grp = str(r["Группа"]).strip()
        place = (r.name % 10) + 1
        if place > 5:
            continue
        ref_code = clean_code(r["Код"])
        ref_rev = r2(to_num(pd.Series([r[ref_col]]))[0])
        my = my_map.get((grp, place))
        if my is None:
            mism += 1
            continue
        if my[0] != ref_code or abs(my[1] - ref_rev) > 0.01:
            mism += 1
            print("  МИСМАТЧ {} {} место {}: мой {} {} vs эталон {} {}".format(
                year, grp, place, my[0], my[1], ref_code, ref_rev))
    check("топ5_по_годам_сверка_{}".format(year), mism, 0, exact=True)

# c) внутри года и группы: выручка убывает; каждая строка ≤ годовой выручке группы
for year in YEARS_TOP5:
    rows = t5y[str(year)]["строки"]
    grp_total = pos_src.loc[pos_src["год"] == year].groupby("группа")["ТО, руб"].sum()
    for grp in sorted(set(r["группа"] for r in rows)):
        grp_rows = [r for r in rows if r["группа"] == grp]
        mono = all(grp_rows[i]["выручка"] >= grp_rows[i + 1]["выручка"]
                   for i in range(len(grp_rows) - 1))
        check("топ5_по_годам_{}_{}_убывание".format(year, grp), mono, True, exact=True)
        total = grp_total.get(grp, 0)
        in_bounds = all(r["выручка"] <= total + 0.01 for r in grp_rows)
        check("топ5_по_годам_{}_{}_в_пределах_группы".format(year, grp),
              in_bounds, True, exact=True)

print("=" * 70)
print("JSON: {}".format(DASH_JSON))
print("Размер: {} байт".format(os.path.getsize(DASH_JSON)))
print("Готово.")