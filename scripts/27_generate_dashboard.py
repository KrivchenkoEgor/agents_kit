#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
27_generate_dashboard.py — генератор интерактивного HTML-дашборда магазина.

Что делает:
  1. Читает output/dashboard/dashboard_data.json (только чтение, исходник не трогаем).
  2. Встраивает данные прямо в HTML как `const DATA = {...};` — без fetch/XHR,
     поэтому файл открывается двойным кликом через file:// и работает без интернета.
  3. Подключает Chart.js: локальный файл assets/chart.umd.min.js, если он скачан,
     иначе — CDN-ссылку.
  4. Пишет один самодостаточный файл output/dashboard/Дашборд_магазина.html.

Запуск:  python3 scripts/27_generate_dashboard.py
Требования: Python 3.8+, только стандартная библиотека (json, os).
"""

import json
import os

# Пути считаем от расположения скрипта, чтобы запуск работал из любой папки.
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_PATH = os.path.join(BASE_DIR, "output", "dashboard", "dashboard_data.json")
OUT_PATH = os.path.join(BASE_DIR, "output", "dashboard", "Дашборд_магазина.html")
ASSETS_PATH = os.path.join(BASE_DIR, "output", "dashboard", "assets", "chart.umd.min.js")

# Локальный Chart.js предпочтительнее: файл откроется без интернета.
CHART_SRC = "assets/chart.umd.min.js" if os.path.exists(ASSETS_PATH) else (
    "https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"
)

# ---------------------------------------------------------------------------
# Шаблон страницы. Плейсхолдеры %%DATA%% и %%CHART_SRC%% заменяются ниже.
# CSS и JS написаны «как есть» (без f-строк), чтобы фигурные скобки не мешали.
# ---------------------------------------------------------------------------
TEMPLATE = """<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Дашборд магазина — продажи молока и колбасы</title>
<style>
:root{
  --blue:#2563eb; --blue-soft:#60a5fa; --blue-pale:#bfdbfe; --blue-hover:#dbeafe;
  --green:#16a34a; --green-soft:#86efac;
  --orange:#f59e0b; --red:#dc2626;
  --text:#111827; --muted:#6b7280; --bg:#f3f4f6; --card:#ffffff; --border:#e5e7eb;
  --radius:12px;
  --shadow:0 1px 3px rgba(0,0,0,.07),0 4px 14px rgba(0,0,0,.05);
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Roboto,'Segoe UI',Arial,Helvetica,sans-serif;background:var(--bg);color:var(--text);line-height:1.45}
.wrap{max-width:1280px;margin:0 auto;padding:20px}
.card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);padding:18px 20px}
header.card{padding:22px 26px;margin-bottom:16px}
h1{font-size:23px;font-weight:700;letter-spacing:-.3px}
.sub{color:var(--muted);font-size:13.5px;margin-top:6px}
.tabs{display:flex;gap:8px;margin:0 0 18px;flex-wrap:wrap;align-items:center}
.tab-btn{padding:9px 22px;border:1px solid var(--border);background:#fff;border-radius:8px;cursor:pointer;
  font-size:15px;font-weight:600;color:var(--muted);font-family:inherit;transition:background .15s,color .15s}
.tab-btn:hover{background:#f9fafb}
.tab-btn.active{background:var(--blue);color:#fff;border-color:var(--blue)}
.tabs-spacer{flex:1}
/* Переключатели категорий — «пилюли» */
.pill{display:inline-flex;align-items:center;gap:7px;padding:9px 16px;border-radius:999px;cursor:pointer;
  font-family:inherit;font-size:14px;font-weight:600;border:1.5px solid var(--border);background:#fff;
  color:var(--muted);transition:background .15s,color .15s,border-color .15s}
.pill .pill-check{display:none;width:14px;height:14px;flex:none}
.pill.on .pill-check{display:block}
.pill.pill-milk.on{background:var(--blue);border-color:var(--blue);color:#fff}
.pill.pill-saus.on{background:var(--green);border-color:var(--green);color:#fff}
.pill.off{background:#f3f4f6;border-color:var(--border);color:#9ca3af}
.pill.off .pill-label{text-decoration:line-through}
.pill:hover{box-shadow:var(--shadow)}
/* Переключатель года для таблицы «Топ-5 по выручке в группах» */
.top5-head{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:3px}
.top5-head h2{margin-bottom:0}
.year-switch{display:inline-flex;gap:2px;background:#f3f4f6;border:1px solid var(--border);
  border-radius:8px;padding:2px;margin-left:auto}
.year-btn{border:0;background:transparent;padding:5px 12px;border-radius:6px;cursor:pointer;
  font-family:inherit;font-size:12.5px;font-weight:600;color:var(--muted);
  transition:background .15s,color .15s}
.year-btn:hover{background:#e5e7eb}
.year-btn.active{background:var(--green);color:#fff}
/* Сортировка по «Продано» в таблице «Топ-5 по выручке в группах» */
.th-sort{cursor:pointer;user-select:none;transition:background .15s}
.th-sort:hover{background:var(--blue-hover)}
.sort-ind{color:var(--green);font-size:11px;margin-left:4px;font-weight:700}
.top5-note{font-size:12px;color:var(--muted);margin-bottom:10px}
.tab-panel{display:none}
#tab-overview{display:block}
/* Сообщение «категория отключена» / «включите хотя бы одну» */
.cat-off-msg{background:#fef3c7;border:1px solid #f59e0b;color:#92400e;border-radius:var(--radius);
  padding:24px 20px;text-align:center;font-size:16px;font-weight:700;margin-bottom:18px}
.kpi-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:18px}
.kpi-value{font-size:29px;font-weight:700;letter-spacing:-.5px;font-variant-numeric:tabular-nums}
.kpi-unit{font-size:15px;font-weight:600;color:var(--muted);margin-left:4px}
.kpi-label{font-size:13px;color:#374151;font-weight:600;margin-top:5px}
.kpi-sub{font-size:12px;color:var(--muted);margin-top:2px}
.kpi-green .kpi-value{color:var(--green)}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
.card h2{font-size:15.5px;font-weight:700;margin-bottom:3px}
.card .note{font-size:12px;color:var(--muted);margin-bottom:10px}
.chart-box{position:relative;height:320px}
.table-wrap{overflow:auto;max-height:430px;border:1px solid var(--border);border-radius:8px}
table{width:100%;border-collapse:collapse;font-size:13px}
th{position:sticky;top:0;background:#f9fafb;text-align:left;padding:8px 10px;border-bottom:2px solid var(--border);
  color:#374151;font-weight:600;white-space:nowrap;z-index:1}
td{padding:7px 10px;border-bottom:1px solid var(--border);vertical-align:top}
tr:hover td{background:#f9fafb}
.num{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}
.tbl-title{font-size:13.5px;font-weight:700;margin:12px 0 6px;color:#374151}
.tbl-title:first-child{margin-top:0}
.badge{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11.5px;font-weight:600;white-space:nowrap}
.bad-red{background:#fee2e2;color:#b91c1c}
.bad-orange{background:#fef3c7;color:#b45309}
.bad-green{background:#dcfce7;color:#15803d}
.bad-gray{background:#f3f4f6;color:#6b7280}
footer{margin-top:22px;color:var(--muted);font-size:12.5px;text-align:center;padding:14px 0 6px}
@media(max-width:900px){
  .grid{grid-template-columns:1fr}
  .kpi-grid{grid-template-columns:repeat(2,1fr)}
}
@media(max-width:600px){.kpi-grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="wrap">

  <header class="card">
    <h1>Дашборд магазина — продажи молока и колбасы</h1>
    <div class="sub" id="header-sub"></div>
  </header>

  <nav class="tabs">
    <button class="tab-btn active" data-tab="overview">Обзор</button>
    <button class="tab-btn" data-tab="milk">Молоко</button>
    <button class="tab-btn" data-tab="saus">Колбаса</button>
    <span class="tabs-spacer"></span>
    <button class="pill pill-milk on" data-cat="молоко" aria-pressed="true" title="Показать/скрыть категорию «Молоко»">
      <svg class="pill-check" viewBox="0 0 14 14" fill="none" aria-hidden="true"><path d="M2.5 7.5L5.5 10.5L11.5 3.5" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>
      <span class="pill-label">Молоко</span>
    </button>
    <button class="pill pill-saus on" data-cat="колбаса" aria-pressed="true" title="Показать/скрыть категорию «Колбаса»">
      <svg class="pill-check" viewBox="0 0 14 14" fill="none" aria-hidden="true"><path d="M2.5 7.5L5.5 10.5L11.5 3.5" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>
      <span class="pill-label">Колбаса</span>
    </button>
  </nav>

  <!-- ======================= ВКЛАДКА «ОБЗОР» ======================= -->
  <section id="tab-overview" class="tab-panel">
    <div class="cat-off-msg" id="both-off-msg" style="display:none">Включите хотя бы одну категорию</div>
    <div class="kpi-grid" id="kpi-grid"></div>
    <div class="grid">
      <div class="card" id="card-a">
        <h2>Динамика выручки по месяцам</h2>
        <div class="note">Молоко — с января 2022, колбаса — с января 2023</div>
        <div class="chart-box"><canvas id="ch-a"></canvas></div>
      </div>
      <div class="card" id="card-b">
        <h2>Выручка по годам</h2>
        <div class="note">* 2026 — неполный год, данные по 10 августа</div>
        <div class="chart-box"><canvas id="ch-b"></canvas></div>
      </div>
      <div class="card" id="card-e1">
        <h2>ABC-структура — молоко</h2>
        <div class="note">Доли выручки за 12 мес по классам A/B/C</div>
        <div class="chart-box"><canvas id="ch-e1"></canvas></div>
      </div>
      <div class="card" id="card-e2">
        <h2>ABC-структура — колбаса</h2>
        <div class="note">Пул A — ядро ассортимента и выручки</div>
        <div class="chart-box"><canvas id="ch-e2"></canvas></div>
      </div>
      <div class="card" id="card-g">
        <h2>Прогноз молока, авг–дек 2026</h2>
        <div class="note">Полоса — 95% доверительный интервал прогноза</div>
        <div class="chart-box"><canvas id="ch-g"></canvas></div>
      </div>
      <div class="card" id="card-h">
        <h2>Топ-10 товаров по выручке — молоко</h2>
        <div class="note">Выручка за 12 мес</div>
        <div class="chart-box"><canvas id="ch-h"></canvas></div>
      </div>
    </div>
  </section>

  <!-- ======================= ВКЛАДКА «МОЛОКО» ======================= -->
  <section id="tab-milk" class="tab-panel">
    <div class="cat-off-msg" id="cat-off-msg-milk" style="display:none">Категория «Молоко» отключена — включите её переключателем в шапке</div>
    <div id="milk-content">
    <div class="grid">
      <div class="card">
        <h2>Динамика выручки — молоко</h2>
        <div class="note">По месяцам с января 2022</div>
        <div class="chart-box"><canvas id="ch-a-milk"></canvas></div>
      </div>
      <div class="card">
        <h2>ABC-структура — молоко</h2>
        <div class="note">Доли выручки за 12 мес по классам</div>
        <div class="chart-box"><canvas id="ch-e-milk"></canvas></div>
      </div>
      <div class="card">
        <h2>Прогноз молока, авг–дек 2026</h2>
        <div class="note">Полоса — 95% доверительный интервал</div>
        <div class="chart-box"><canvas id="ch-g-milk"></canvas></div>
      </div>
      <div class="card">
        <h2>Топ-10 товаров по выручке</h2>
        <div class="note">Выручка за 12 мес</div>
        <div class="chart-box"><canvas id="ch-h-milk"></canvas></div>
      </div>
    </div>
    <div class="grid">
      <div class="card">
        <h2>Новинки молока по годам</h2>
        <div class="note">Судьба новинок: сколько осталось в продаже и вышли в класс A</div>
        <div class="table-wrap" id="tbl-new"></div>
      </div>
      <div class="card">
        <h2>Шринкфляция молока</h2>
        <div class="note">Уменьшение фасовки при прежней цене</div>
        <div class="table-wrap" id="tbl-shrink"></div>
      </div>
      <div class="card">
        <h2>Статусы товаров группы C</h2>
        <div class="note">Что происходит с аутсайдерами ассортимента молока</div>
        <div class="table-wrap" id="tbl-gc"></div>
      </div>
    </div>
    </div>
  </section>

  <!-- ======================= ВКЛАДКА «КОЛБАСА» ======================= -->
  <section id="tab-saus" class="tab-panel">
    <div class="cat-off-msg" id="cat-off-msg-saus" style="display:none">Категория «Колбаса» отключена — включите её переключателем в шапке</div>
    <div id="saus-content">
    <div class="grid">
      <div class="card">
        <h2>Структура продаж колбасы по группам</h2>
        <div class="note">Выручка за всё время по группам товаров</div>
        <div class="chart-box"><canvas id="ch-c"></canvas></div>
      </div>
      <div class="card">
        <h2>Прибыль по группам — колбаса</h2>
        <div class="note">Валовая прибыль и наценка по группам</div>
        <div class="chart-box"><canvas id="ch-d"></canvas></div>
      </div>
      <div class="card">
        <h2>Потерянные топы — колбаса</h2>
        <div class="note">Позиции из топ-2023/2024 без продаж по коду в 2026</div>
        <div class="chart-box"><canvas id="ch-f"></canvas></div>
      </div>
      <div class="card">
        <h2>ABC-структура — колбаса</h2>
        <div class="note">Пул A — ядро выручки</div>
        <div class="chart-box"><canvas id="ch-e-saus"></canvas></div>
      </div>
    </div>
    <div class="grid">
      <div class="card">
        <h2>Потерянные топы 2023/2024</h2>
        <div class="note">Исчезнувшие позиции и их продолжения</div>
        <div class="table-wrap" id="tbl-lost"></div>
      </div>
      <div class="card">
        <div class="top5-head">
          <h2>Топ-5 по выручке в группах</h2>
          <div class="year-switch" role="group" aria-label="Выбор года">
            <button class="year-btn active" data-year="all">Всё время</button>
            <button class="year-btn" data-year="2023">2023</button>
            <button class="year-btn" data-year="2024">2024</button>
            <button class="year-btn" data-year="2025">2025</button>
            <button class="year-btn" data-year="2026">2026</button>
          </div>
        </div>
        <div class="note top5-note" id="top5-note">за всё время: 2023-01-01 – 2026-08-10</div>
        <div class="table-wrap" id="tbl-top5"></div>
      </div>
      <div class="card">
        <h2>Топ по валовой прибыли</h2>
        <div class="note">20 самых прибыльных позиций</div>
        <div class="table-wrap" id="tbl-profit"></div>
      </div>
      <div class="card">
        <h2>Где теряем</h2>
        <div class="note">Убыточные позиции — кандидаты на вывод из ассортимента</div>
        <div class="table-wrap" id="tbl-lose"></div>
      </div>
    </div>
    </div>
  </section>

  <footer>
    Данные: отчёты анализа продаж, собрано <span id="footer-date"></span>.<br>
    Единицы: молоко — штуки упаковок; колбаса — весовые в кг, штучные в шт.
  </footer>
</div>

<script src="%%CHART_SRC%%"></script>
<script>
const DATA = %%DATA%%;

(function () {
  'use strict';
  var M = DATA['молоко'];
  var S = DATA['колбаса'];

  /* ---------- глобальное состояние: включённые категории ---------- */
  var cats = { 'молоко': true, 'колбаса': true };
  var currentTab = 'overview';
  var top5Year = 'all'; // выбранный год для таблицы «Топ-5 по выручке в группах»
  var top5Sort = null; // сортировка «Продано»: {key:'продано', dir:'desc'|'asc'}; null — исходный порядок
  var charts = {}; // реестр созданных графиков по id canvas

  function destroyChart(id) {
    if (charts[id]) { charts[id].destroy(); delete charts[id]; }
  }

  /* ---------- форматирование чисел (русский разделитель) ---------- */
  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function fmt(n) { // целое с пробелами: 1 234 567
    if (n == null || isNaN(n)) return '—';
    return Math.round(Number(n)).toLocaleString('ru-RU');
  }
  function fmtM(n) { // большие суммы в млн с 1 знаком: 1 065,9 млн
    if (n == null || isNaN(n)) return '—';
    n = Number(n);
    if (Math.abs(n) >= 1e6) {
      return (n / 1e6).toLocaleString('ru-RU', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' млн';
    }
    return Math.round(n).toLocaleString('ru-RU');
  }
  function fmtP(n) { // проценты с 1 знаком
    if (n == null || isNaN(n)) return '—';
    return Number(n).toLocaleString('ru-RU', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + '%';
  }

  /* ---------- общие настройки Chart.js ---------- */
  Chart.defaults.font.family = "Roboto, 'Segoe UI', Arial, sans-serif";
  Chart.defaults.font.size = 12;
  Chart.defaults.color = '#374151';

  var moneyTicks = { callback: function (v) { return fmtM(v); } };
  var intTicks = { callback: function (v) { return (Math.round(v) === v) ? fmt(v) : ''; } };

  /* Плагин: текст в центре доната */
  var centerText = {
    id: 'centerText',
    afterDraw: function (chart, args, opts) {
      if (!opts || !opts.text) return;
      var ctx = chart.ctx;
      var meta = chart.getDatasetMeta(0);
      if (!meta.data.length) return;
      var x = meta.data[0].x, y = meta.data[0].y;
      var size = opts.size || 16;
      ctx.save();
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.font = '600 ' + size + 'px Roboto, Arial, sans-serif';
      ctx.fillStyle = opts.color || '#111827';
      var lines = String(opts.text).split('\\n');
      for (var i = 0; i < lines.length; i++) {
        ctx.fillText(lines[i], x, y + (i - (lines.length - 1) / 2) * size * 1.25);
      }
      ctx.restore();
    }
  };
  Chart.register(centerText);

  /* ---------- график a: динамика выручки по месяцам ---------- */
  function buildMonthly(id, showMilk, showSaus) {
    var ctx = document.getElementById(id);
    if (!ctx) return;
    destroyChart(id);
    var labels = M['месяцы'].map(function (x) { return x['месяц']; });
    var sausByMonth = {};
    S['месяцы'].forEach(function (x) { sausByMonth[x['месяц']] = x['руб']; });
    var datasets = [];
    if (showMilk) {
      datasets.push({
        label: 'Молоко',
        data: M['месяцы'].map(function (x) { return x['руб']; }),
        borderColor: '#2563eb', backgroundColor: '#2563eb',
        tension: 0.3, pointRadius: 0, borderWidth: 2, spanGaps: false
      });
    }
    if (showSaus) {
      datasets.push({
        label: 'Колбаса',
        data: M['месяцы'].map(function (x) {
          return (sausByMonth[x['месяц']] !== undefined) ? sausByMonth[x['месяц']] : null;
        }),
        borderColor: '#16a34a', backgroundColor: '#16a34a',
        tension: 0.3, pointRadius: 0, borderWidth: 2, spanGaps: false
      });
    }
    charts[id] = new Chart(ctx, {
      type: 'line',
      data: { labels: labels, datasets: datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { position: 'top' },
          tooltip: {
            callbacks: {
              label: function (c) {
                var v = c.parsed.y;
                if (v == null) return null;
                return ' ' + c.dataset.label + ': ' + fmtM(v) + ' руб';
              }
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: {
              maxRotation: 60, autoSkip: false,
              callback: function (val, idx) { return (idx % 3 === 0) ? this.getLabelForValue(val) : ''; }
            }
          },
          y: { title: { display: true, text: 'Выручка, руб' }, ticks: moneyTicks }
        }
      }
    });
  }

  /* ---------- график b: выручка по годам ---------- */
  function buildYears(id, showMilk, showSaus) {
    var ctx = document.getElementById(id);
    if (!ctx) return;
    destroyChart(id);
    var yearLabels = ['2022', '2023', '2024', '2025', '2026*'];
    var milk = {};
    M['месяцы'].forEach(function (x) {
      var y = x['месяц'].slice(0, 4);
      milk[y] = (milk[y] || 0) + x['руб'];
    });
    var saus = {};
    S['годы'].forEach(function (x) { saus[String(x['год'])] = x['руб']; });
    var datasets = [];
    if (showMilk) {
      datasets.push({ label: 'Молоко', data: yearLabels.map(function (y) { return milk[y] || null; }), backgroundColor: '#2563eb', borderRadius: 4 });
    }
    if (showSaus) {
      datasets.push({ label: 'Колбаса', data: yearLabels.map(function (y) { return saus[y] || null; }), backgroundColor: '#16a34a', borderRadius: 4 });
    }
    charts[id] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: yearLabels,
        datasets: datasets
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top' },
          tooltip: {
            callbacks: {
              label: function (c) {
                var v = c.parsed.y;
                if (v == null) return null;
                return ' ' + c.dataset.label + ': ' + fmtM(v) + ' руб';
              }
            }
          }
        },
        scales: {
          x: { grid: { display: false } },
          y: { title: { display: true, text: 'Выручка, руб' }, ticks: moneyTicks }
        }
      }
    });
  }

  /* ---------- график c: структура продаж колбасы по группам ---------- */
  function buildGroupsRev(id) {
    var ctx = document.getElementById(id);
    if (!ctx) return;
    destroyChart(id);
    var g = S['группы_итог'];
    charts[id] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: g.map(function (x) { return x['группа']; }),
        datasets: [{ label: 'Выручка', data: g.map(function (x) { return x['выручка']; }), backgroundColor: '#16a34a', borderRadius: 4 }]
      },
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: function (c) {
                var d = g[c.dataIndex];
                return ' Выручка: ' + fmtM(c.parsed.x) + ' руб · доля ' + fmtP(d['доля']);
              }
            }
          }
        },
        scales: {
          x: { title: { display: true, text: 'Выручка за всё время, руб' }, ticks: moneyTicks },
          y: { grid: { display: false } }
        }
      }
    });
  }

  /* ---------- график d: прибыль по группам, колбаса ---------- */
  function buildGroupsProfit(id) {
    var ctx = document.getElementById(id);
    if (!ctx) return;
    destroyChart(id);
    var g = S['прибыль_группы'];
    charts[id] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: g.map(function (x) { return x['группа']; }),
        datasets: [{
          label: 'Прибыль',
          data: g.map(function (x) { return x['прибыль']; }),
          backgroundColor: g.map(function (x) { return x['прибыль'] >= 0 ? '#16a34a' : '#dc2626'; }),
          borderRadius: 4
        }]
      },
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: function (c) {
                var d = g[c.dataIndex];
                return ' Прибыль: ' + fmtM(c.parsed.x) + ' руб · наценка ' + fmtP(d['наценка']) +
                       ' · доля прибыли ' + fmtP(d['доля_прибыли']);
              }
            }
          }
        },
        scales: {
          x: { title: { display: true, text: 'Прибыль за всё время, руб' }, ticks: moneyTicks },
          y: { grid: { display: false } }
        }
      }
    });
  }

  /* ---------- график e1: ABC-структура молока (донат) ---------- */
  function buildDonutMilk(id) {
    var ctx = document.getElementById(id);
    if (!ctx) return;
    destroyChart(id);
    var sums = { A: 0, B: 0, C: 0 };
    var cnt = { A: 0, B: 0, C: 0 };
    M['abc'].forEach(function (r) {
      sums[r['группа']] = (sums[r['группа']] || 0) + r['выручка12'];
      cnt[r['группа']] = (cnt[r['группа']] || 0) + 1;
    });
    var keys = ['A', 'B', 'C'];
    var labels = keys.map(function (k) { return 'Класс ' + k + ' (' + cnt[k] + ' тов.)'; });
    var total = keys.reduce(function (a, k) { return a + sums[k]; }, 0);
    charts[id] = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: labels,
        datasets: [{
          data: keys.map(function (k) { return sums[k]; }),
          backgroundColor: ['#2563eb', '#60a5fa', '#bfdbfe'],
          borderColor: '#ffffff', borderWidth: 2
        }]
      },
      options: {
        cutout: '62%', responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom' },
          centerText: { text: M['abc'].length + ' товаров' },
          tooltip: {
            callbacks: {
              label: function (c) {
                return ' ' + c.label + ': ' + fmtM(c.parsed) + ' руб · ' + fmtP(c.parsed / total * 100);
              }
            }
          }
        }
      }
    });
  }

  /* ---------- график e2: ABC-структура колбасы (донат) ---------- */
  function buildDonutSaus(id) {
    var ctx = document.getElementById(id);
    if (!ctx) return;
    destroyChart(id);
    var pool = S['abc']['пул_A'];
    var total = S['abc']['товаров'];
    var rest = total - pool['товаров'];
    charts[id] = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: ['Группа A', 'Остальные'],
        datasets: [{
          data: [pool['товаров'], rest],
          backgroundColor: ['#16a34a', '#86efac'],
          borderColor: '#ffffff', borderWidth: 2
        }]
      },
      options: {
        cutout: '62%', responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom' },
          centerText: { text: fmtP(pool['доля_выручки']) + '\\nвыручки', size: 15 },
          tooltip: {
            callbacks: {
              label: function (c) {
                return ' ' + c.label + ': ' + fmt(c.parsed) + ' товаров · ' + fmtP(c.parsed / total * 100);
              }
            }
          }
        }
      }
    });
  }

  /* ---------- график f: потерянные топы, колбаса ---------- */
  function buildLost(id) {
    var ctx = document.getElementById(id);
    if (!ctx) return;
    destroyChart(id);
    var L23 = S['потерянные_2023'];
    var L24 = S['потерянные_2024'];
    var grpSet = {};
    L23.concat(L24).forEach(function (x) { grpSet[x['группа']] = true; });
    var groups = Object.keys(grpSet).sort(function (a, b) {
      var ta = L23.filter(function (x) { return x['группа'] === a; }).length +
               L24.filter(function (x) { return x['группа'] === a; }).length;
      var tb = L23.filter(function (x) { return x['группа'] === b; }).length +
               L24.filter(function (x) { return x['группа'] === b; }).length;
      return tb - ta;
    });
    function cnt(list, g, st) {
      return list.filter(function (x) { return x['группа'] === g && x['статус'] === st; }).length;
    }
    charts[id] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: groups,
        datasets: [
          { label: '2023 · исчез', data: groups.map(function (g) { return cnt(L23, g, 'исчез'); }), backgroundColor: '#dc2626' },
          { label: '2023 · шринкфляция', data: groups.map(function (g) { return cnt(L23, g, 'шринкфляция'); }), backgroundColor: '#f59e0b' },
          { label: '2024 · исчез', data: groups.map(function (g) { return cnt(L24, g, 'исчез'); }), backgroundColor: 'rgba(220,38,38,0.55)' },
          { label: '2024 · шринкфляция', data: groups.map(function (g) { return cnt(L24, g, 'шринкфляция'); }), backgroundColor: 'rgba(245,158,11,0.55)' }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top' },
          tooltip: {
            callbacks: {
              label: function (c) { return ' ' + c.dataset.label + ': ' + fmt(c.parsed.y) + ' поз.'; }
            }
          }
        },
        scales: {
          x: { stacked: true, grid: { display: false } },
          y: { stacked: true, ticks: intTicks, title: { display: true, text: 'Потерянных позиций' } }
        }
      }
    });
  }

  /* ---------- график g: прогноз молока с полосой 95% ---------- */
  function buildForecast(id) {
    var ctx = document.getElementById(id);
    if (!ctx) return;
    destroyChart(id);
    var P = M['прогноз'];
    var labels = P.map(function (p) { return p['месяц']; });
    charts[id] = new Chart(ctx, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [
          { label: 'Нижняя граница', data: P.map(function (p) { return p['низ']; }), borderWidth: 0, pointRadius: 0, fill: false },
          { label: '95% интервал', data: P.map(function (p) { return p['верх']; }), borderWidth: 0, pointRadius: 0, fill: { target: -1, above: 'rgba(37,99,235,0.15)' } },
          { label: 'Прогноз', data: P.map(function (p) { return p['руб']; }), borderColor: '#2563eb', backgroundColor: '#2563eb', pointRadius: 3, tension: 0.25, borderWidth: 2 }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'top',
            labels: { filter: function (item) { return ['Прогноз', '95% интервал'].indexOf(item.text) >= 0; } }
          },
          tooltip: {
            callbacks: {
              label: function (c) {
                var i = c.dataIndex;
                if (c.dataset.label === 'Прогноз') return ' Прогноз: ' + fmtM(P[i]['руб']) + ' руб (' + fmt(P[i]['штук']) + ' шт)';
                if (c.dataset.label === 'Нижняя граница') return ' Нижняя граница: ' + fmtM(P[i]['низ']) + ' руб';
                if (c.dataset.label === '95% интервал') return ' 95% интервал: ' + fmtM(P[i]['низ']) + ' – ' + fmtM(P[i]['верх']) + ' руб';
                return null;
              }
            }
          }
        },
        scales: {
          x: { grid: { display: false } },
          y: { title: { display: true, text: 'Выручка, руб' }, ticks: moneyTicks }
        }
      }
    });
  }

  /* ---------- график h: топ-10 молока по выручке ---------- */
  function buildTop10(id) {
    var ctx = document.getElementById(id);
    if (!ctx) return;
    destroyChart(id);
    var t = M['топ10'];
    charts[id] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: t.map(function (x) { return x['название']; }),
        datasets: [{ label: 'Выручка 12 мес', data: t.map(function (x) { return x['выручка12']; }), backgroundColor: '#2563eb', borderRadius: 4 }]
      },
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: function (c) {
                var d = t[c.dataIndex];
                return ' Выручка: ' + fmtM(c.parsed.x) + ' руб · штук: ' + fmt(d['штук12']);
              }
            }
          }
        },
        scales: {
          x: { title: { display: true, text: 'Выручка за 12 мес, руб' }, ticks: moneyTicks },
          y: {
            grid: { display: false },
            ticks: {
              autoSkip: false, font: { size: 10.5 },
              callback: function (val) {
                var lbl = this.getLabelForValue(val);
                return lbl.length > 58 ? lbl.slice(0, 55) + '…' : lbl;
              }
            }
          }
        }
      }
    });
  }

  /* ---------- KPI-карточки (пересчитываются по включённым категориям) ---------- */
  function buildKPI() {
    var el = document.getElementById('kpi-grid');
    if (!el) return;
    var milkOn = cats['молоко'], sausOn = cats['колбаса'];
    var isIn12 = function (m) {
      var y = +m['месяц'].slice(0, 4), mo = +m['месяц'].slice(5);
      return (y === 2025 && mo >= 8) || (y === 2026 && mo <= 7);
    };
    var milk12 = milkOn ? M['месяцы'].filter(isIn12).reduce(function (s, m) { return s + m['руб']; }, 0) : 0;
    var milk12шт = milkOn ? M['месяцы'].filter(isIn12).reduce(function (s, m) { return s + m['штук']; }, 0) : 0;
    var saus12 = sausOn ? S['abc']['выручка_год'] : 0;
    var allProfit = S['итоги_всего']['прибыль'];
    var margin = S['итоги_всего']['прибыль'] / S['итоги_всего']['руб'] * 100;
    var accMilk = milkOn ? M['истории'].filter(function (h) { return h['статус'] === 'осталась'; }).length : 0;
    var accSaus = sausOn ? S['итоги_всего']['товаров'] : 0;
    var dev = M['сверка_прогноза']['откл_штук_проц'];

    var cards = [];
    var revSub = [];
    if (milkOn) revSub.push('молоко ' + fmtM(milk12));
    if (sausOn) revSub.push('колбаса ' + fmtM(saus12));
    cards.push({ v: fmtM(milk12 + saus12), u: '', t: 'Выручка за 12 мес',
      s: revSub.join(' · ') + ' руб' });

    if (sausOn) {
      cards.push({ v: fmtM(allProfit), u: '', t: 'Валовая прибыль',
        s: 'колбаса, за всё время' });
      cards.push({ v: fmtP(margin), u: '', t: 'Средняя наценка',
        s: 'колбаса: прибыль / выручка за всё время' });
    }

    if (milkOn) {
      var natSub = ['молоко, 12 мес'];
      if (sausOn) natSub.push('колбаса: ' + fmt(S['итоги_всего']['штук']) + ' шт + ' + fmt(S['итоги_всего']['кг']) + ' кг за всё время');
      cards.push({ v: fmt(milk12шт), u: 'шт', t: 'Продано натуральных единиц',
        s: natSub.join(' · ') });
    }

    var accSub = [];
    if (sausOn) accSub.push('колбаса ' + fmt(accSaus));
    if (milkOn) accSub.push('молоко ' + fmt(accMilk));
    cards.push({ v: fmt(accMilk + accSaus), u: '', t: 'Товаров в ассортименте',
      s: accSub.join(' · ') });

    if (milkOn) {
      cards.push({ v: fmtP(dev), u: '', t: 'Точность прогноза',
        s: 'отклонение прогноза августа от факта, молоко', accent: 'green' });
    }

    el.innerHTML = cards.map(function (c) {
      return '<div class="card kpi' + (c.accent ? ' kpi-' + c.accent : '') + '">' +
        '<div class="kpi-value">' + c.v + (c.u ? '<span class="kpi-unit">' + c.u + '</span>' : '') + '</div>' +
        '<div class="kpi-label">' + c.t + '</div>' +
        '<div class="kpi-sub">' + c.s + '</div></div>';
    }).join('');
  }

  /* ---------- таблицы ---------- */
  function makeTable(headers, rows) {
    var head = '<thead><tr>' + headers.map(function (h) {
      var cls = (h.cls || '') + (h.sortable ? ' th-sort' : '');
      var ind = h.sortInd ? '<span class="sort-ind">' + h.sortInd + '</span>' : '';
      var title = h.sortable ? ' title="Сортировать по количеству продаж"' : '';
      return '<th class="' + cls + '"' + title + '>' + esc(h.t) + ind + '</th>';
    }).join('') + '</tr></thead>';
    var body = '<tbody>' + rows.map(function (r) {
      return '<tr>' + r.map(function (c, i) {
        return '<td class="' + (headers[i].cls || '') + '">' + c + '</td>';
      }).join('') + '</tr>';
    }).join('') + '</tbody>';
    return '<table>' + head + body + '</table>';
  }
  var H = [
    { t: 'Группа' }, { t: 'Название' }, { t: 'Выручка в базовом году, руб', cls: 'num' },
    { t: 'Статус' }, { t: 'Продолжение' }
  ];
  function lostRows(L, yearKey) {
    return L.map(function (x) {
      return [
        esc(x['группа']), esc(x['название']),
        '<span class="num">' + fmt(x['выручка' + yearKey]) + '</span>',
        x['статус'] === 'исчез'
          ? '<span class="badge bad-red">исчез</span>'
          : '<span class="badge bad-orange">шринкфляция</span>',
        x['продолжение'] ? esc(x['продолжение']) : '—'
      ];
    });
  }
  /* ---------- таблица «Топ-5 по выручке в группах» с выбором года ---------- */
  function top5Rows(rows) {
    return rows.map(function (x) {
      var t = x['тип'] || '';
      var sold = (t.indexOf('кг') >= 0 || t.indexOf('весовой') >= 0)
        ? fmt(x['продано_кг']) + ' кг'
        : fmt(x['продано']) + ' шт';
      return [esc(x['группа']), esc(x['место']), esc(x['название']), esc(x['тип']),
        '<span class="num">' + sold + '</span>', '<span class="num">' + fmt(x['выручка']) + '</span>'];
    });
  }
  /* Ключ сортировки «Продано»: единая мера — продано_кг (≈ кг), фолбэк на продано */
  function top5SortVal(x) {
    var v = x['продано_кг'];
    if (v == null || isNaN(v)) v = x['продано'];
    return (v == null || isNaN(v)) ? null : Number(v);
  }
  function top5Sorted(rows) {
    if (!top5Sort) return rows;
    var dir = top5Sort.dir === 'asc' ? 1 : -1;
    var sorted = rows.slice();
    sorted.sort(function (a, b) {
      var va = top5SortVal(a), vb = top5SortVal(b);
      if (va == null && vb == null) return 0;
      if (va == null) return 1;
      if (vb == null) return -1;
      return (va - vb) * dir;
    });
    return sorted;
  }
  function buildTop5Table() {
    var el = document.getElementById('tbl-top5');
    if (!el) return;
    var rows;
    if (top5Year === 'all') {
      rows = S['топ5_выручка'];
    } else {
      var byYear = S['топ5_по_годам'][top5Year];
      rows = byYear ? byYear['строки'] : [];
    }
    rows = top5Sorted(rows);
    var sortInd = top5Sort ? (top5Sort.dir === 'asc' ? '▼' : '▲') : '';
    el.innerHTML = makeTable([
      { t: 'Группа' }, { t: 'Место' }, { t: 'Название' }, { t: 'Тип' },
      { t: 'Продано (шт/кг)', cls: 'num', sortable: true, sortInd: sortInd },
      { t: 'Выручка, руб', cls: 'num' }
    ], top5Rows(rows));
    var sortTh = el.querySelector('.th-sort');
    if (sortTh) {
      sortTh.addEventListener('click', function () {
        if (!top5Sort || top5Sort.key !== 'продано') {
          top5Sort = { key: 'продано', dir: 'desc' };
        } else {
          top5Sort.dir = top5Sort.dir === 'desc' ? 'asc' : 'desc';
        }
        buildTop5Table();
      });
    }
    var note = document.getElementById('top5-note');
    if (note) {
      if (top5Year === 'all') note.textContent = 'за всё время: 2023-01-01 – 2026-08-10';
      else if (top5Year === '2026') note.textContent = '* 2026 — данные по 10 августа';
      else note.textContent = top5Year + ' год';
    }
    var btns = document.querySelectorAll('.year-btn');
    for (var i = 0; i < btns.length; i++) {
      btns[i].classList.toggle('active', btns[i].getAttribute('data-year') === top5Year);
    }
  }
  function buildTablesSaus() {
    var el;
    el = document.getElementById('tbl-lost');
    if (el) {
      var L23 = S['потерянные_2023'], L24 = S['потерянные_2024'];
      el.innerHTML =
        '<h3 class="tbl-title">Потерянные топы 2023 — ' + L23.length + ' поз.</h3>' +
        makeTable(H, lostRows(L23, '2023')) +
        '<h3 class="tbl-title">Потерянные топы 2024 — ' + L24.length + ' поз.</h3>' +
        makeTable(H, lostRows(L24, '2024'));
    }
    buildTop5Table();
    el = document.getElementById('tbl-profit');
    if (el) {
      var rowsP = S['прибыль_топ'].map(function (x) {
        return [esc(x['место']), esc(x['название']),
          '<span class="num">' + fmt(x['прибыль']) + '</span>',
          '<span class="num">' + fmtP(x['наценка']) + '</span>',
          x['в2026'] === 'да' ? '<span class="badge bad-green">да</span>' : '<span class="badge bad-gray">нет</span>'];
      });
      el.innerHTML = makeTable([
        { t: 'Место' }, { t: 'Название' }, { t: 'Прибыль, руб', cls: 'num' },
        { t: 'Наценка %', cls: 'num' }, { t: 'В продаже 2026' }
      ], rowsP);
    }
    el = document.getElementById('tbl-lose');
    if (el) {
      var rowsL = S['прибыль_теряем'].map(function (x) {
        return [esc(x['место']), esc(x['группа']), esc(x['название']),
          '<span class="num" style="color:var(--red)">' + fmt(x['прибыль']) + '</span>',
          '<span class="num">' + fmtP(x['наценка']) + '</span>',
          x['в2026'] === 'да' ? '<span class="badge bad-green">да</span>' : '<span class="badge bad-gray">нет</span>'];
      });
      el.innerHTML = makeTable([
        { t: 'Место' }, { t: 'Группа' }, { t: 'Название' }, { t: 'Прибыль, руб', cls: 'num' },
        { t: 'Наценка %', cls: 'num' }, { t: 'В продаже 2026' }
      ], rowsL);
    }
  }
  function buildTablesMilk() {
    var el;
    el = document.getElementById('tbl-new');
    if (el) {
      var rowsN = M['новинки'].map(function (x) {
        return [esc(x['год']), '<span class="num">' + fmt(x['новинок']) + '</span>',
          '<span class="num">' + fmt(x['живы']) + '</span>',
          '<span class="num">' + fmt(x['звёздыA']) + '</span>',
          '<span class="num">' + fmtP(x['доля_живых']) + '</span>'];
      });
      el.innerHTML = makeTable([
        { t: 'Год' }, { t: 'Новинок', cls: 'num' }, { t: 'Живы', cls: 'num' },
        { t: 'Звёзды A', cls: 'num' }, { t: 'Доля живых %', cls: 'num' }
      ], rowsN);
    }
    el = document.getElementById('tbl-shrink');
    if (el) {
      var rowsS = M['шринкфляция'].map(function (x) {
        return [esc(x['название']),
          '<span class="num">' + fmt(x['фасовка_старая']) + ' мл</span>',
          '<span class="num">' + fmt(x['фасовка_новая']) + ' мл</span>',
          '<span class="num" style="color:var(--red)">' + fmtP(x['зазор']) + '</span>'];
      });
      el.innerHTML = makeTable([
        { t: 'Название' }, { t: 'Старая фасовка', cls: 'num' },
        { t: 'Новая фасовка', cls: 'num' }, { t: 'Зазор %', cls: 'num' }
      ], rowsS);
    }
    el = document.getElementById('tbl-gc');
    if (el) {
      var rowsG = M['группаC_статусы'].map(function (x) {
        return [esc(x['статус']), '<span class="num">' + fmt(x['товаров']) + '</span>',
          '<span class="num">' + fmtM(x['выручка12']) + '</span>',
          '<span class="num">' + fmtP(x['доля_товаров']) + '</span>',
          '<span class="num">' + fmtP(x['доля_выручки']) + '</span>'];
      });
      el.innerHTML = makeTable([
        { t: 'Статус' }, { t: 'Товаров', cls: 'num' }, { t: 'Выручка 12 мес, руб', cls: 'num' },
        { t: 'Доля товаров %', cls: 'num' }, { t: 'Доля выручки %', cls: 'num' }
      ], rowsG);
    }
  }

  /* ---------- рендер вкладок по состоянию cats ---------- */
  function safe(fn) {
    try { fn(); } catch (e) { console.error('Ошибка построения графика:', e); }
  }
  function setVisible(id, visible) {
    var el = document.getElementById(id);
    if (el) el.style.display = visible ? '' : 'none';
  }
  function renderOverview() {
    var milkOn = cats['молоко'], sausOn = cats['колбаса'];
    var bothOff = !milkOn && !sausOn;
    setVisible('both-off-msg', bothOff);
    setVisible('kpi-grid', !bothOff);
    setVisible('card-a', !bothOff);
    setVisible('card-b', !bothOff);
    setVisible('card-e1', milkOn);
    setVisible('card-e2', sausOn);
    setVisible('card-g', milkOn);
    setVisible('card-h', milkOn);
    if (bothOff) {
      ['ch-a', 'ch-b', 'ch-e1', 'ch-e2', 'ch-g', 'ch-h'].forEach(destroyChart);
      return;
    }
    safe(function () { buildKPI(); });
    safe(function () { buildMonthly('ch-a', milkOn, sausOn); });
    safe(function () { buildYears('ch-b', milkOn, sausOn); });
    if (milkOn) {
      safe(function () { buildDonutMilk('ch-e1'); });
      safe(function () { buildForecast('ch-g'); });
      safe(function () { buildTop10('ch-h'); });
    } else {
      destroyChart('ch-e1'); destroyChart('ch-g'); destroyChart('ch-h');
    }
    if (sausOn) {
      safe(function () { buildDonutSaus('ch-e2'); });
    } else {
      destroyChart('ch-e2');
    }
  }
  function renderMilk() {
    var milkOn = cats['молоко'];
    setVisible('cat-off-msg-milk', !milkOn);
    setVisible('milk-content', milkOn);
    if (!milkOn) {
      ['ch-a-milk', 'ch-e-milk', 'ch-g-milk', 'ch-h-milk'].forEach(destroyChart);
      return;
    }
    safe(function () { buildMonthly('ch-a-milk', true, false); });
    safe(function () { buildDonutMilk('ch-e-milk'); });
    safe(function () { buildForecast('ch-g-milk'); });
    safe(function () { buildTop10('ch-h-milk'); });
    safe(function () { buildTablesMilk(); });
  }
  function renderSaus() {
    var sausOn = cats['колбаса'];
    setVisible('cat-off-msg-saus', !sausOn);
    setVisible('saus-content', sausOn);
    if (!sausOn) {
      ['ch-c', 'ch-d', 'ch-f', 'ch-e-saus'].forEach(destroyChart);
      return;
    }
    safe(function () { buildGroupsRev('ch-c'); });
    safe(function () { buildGroupsProfit('ch-d'); });
    safe(function () { buildLost('ch-f'); });
    safe(function () { buildDonutSaus('ch-e-saus'); });
    safe(function () { buildTablesSaus(); });
  }
  function renderTab(name) {
    if (name === 'overview') renderOverview();
    else if (name === 'milk') renderMilk();
    else if (name === 'saus') renderSaus();
  }
  function renderCurrentTab() {
    renderTab(currentTab);
  }
  function switchTab(name) {
    currentTab = name;
    var btns = document.querySelectorAll('.tab-btn');
    for (var i = 0; i < btns.length; i++) {
      btns[i].classList.toggle('active', btns[i].getAttribute('data-tab') === name);
    }
    var panels = document.querySelectorAll('.tab-panel');
    for (var j = 0; j < panels.length; j++) {
      panels[j].style.display = (panels[j].id === 'tab-' + name) ? 'block' : 'none';
    }
    renderTab(name);
  }
  var tabBtns = document.querySelectorAll('.tab-btn');
  for (var k = 0; k < tabBtns.length; k++) {
    tabBtns[k].addEventListener('click', function () { switchTab(this.getAttribute('data-tab')); });
  }

  /* ---------- переключатели категорий (пилюли) ---------- */
  function syncPills() {
    var pills = document.querySelectorAll('.pill');
    for (var i = 0; i < pills.length; i++) {
      var on = cats[pills[i].getAttribute('data-cat')];
      pills[i].classList.toggle('on', on);
      pills[i].classList.toggle('off', !on);
      pills[i].setAttribute('aria-pressed', on ? 'true' : 'false');
    }
  }
  var pillBtns = document.querySelectorAll('.pill');
  for (var p = 0; p < pillBtns.length; p++) {
    pillBtns[p].addEventListener('click', function () {
      var cat = this.getAttribute('data-cat');
      cats[cat] = !cats[cat];
      syncPills();
      renderCurrentTab();
    });
  }

  /* ---------- переключатель года таблицы «Топ-5 по выручке в группах» ---------- */
  var yearBtns = document.querySelectorAll('.year-btn');
  for (var y = 0; y < yearBtns.length; y++) {
    yearBtns[y].addEventListener('click', function () {
      top5Year = this.getAttribute('data-year');
      buildTop5Table();
    });
  }

  /* ---------- шапка и подвал ---------- */
  var hs = document.getElementById('header-sub');
  if (hs) {
    hs.textContent = 'Периоды данных: молоко ' + DATA['meta']['периоды']['молоко'] +
      ' · колбаса ' + DATA['meta']['периоды']['колбаса'] +
      ' · собрано ' + DATA['meta']['собрано'];
  }
  var fd = document.getElementById('footer-date');
  if (fd) fd.textContent = DATA['meta']['собрано'];

  /* ---------- старт ---------- */
  syncPills();
  switchTab('overview');
})();
</script>
</body>
</html>
"""


def main():
    """Читает JSON, собирает HTML и пишет файл дашборда."""
    with open(JSON_PATH, encoding="utf-8") as f:
        data = json.load(f)

    # json.dumps(ensure_ascii=False) оставляет кириллицу и числа как есть.
    # Экранируем "</" как "<\/" — это валидный JSON-escape, зато строка
    # никогда не сможет «закрыть» тег <script> раньше времени.
    data_json = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")

    html = TEMPLATE.replace("%%DATA%%", data_json).replace("%%CHART_SRC%%", CHART_SRC)

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    size = os.path.getsize(OUT_PATH)
    charts = html.count("<canvas")
    print("Файл: %s" % OUT_PATH)
    print("Размер: %d байт (%.1f КБ)" % (size, size / 1024.0))
    print("Графиков (canvas): %d" % charts)
    print("Chart.js: %s" % ("локальный assets/chart.umd.min.js" if CHART_SRC.startswith("assets") else "CDN"))
    print("ГОТОВО")


if __name__ == "__main__":
    main()