#!/usr/bin/env python3
import json
import re
import os
from collections import Counter, defaultdict
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_EXCEL = True
except ImportError:
    HAS_EXCEL = False

JSON_PATH = "ChatExport_2026-05-23/result.json"

print("Читаем JSON...")
with open(JSON_PATH, "r", encoding="utf-8") as f:
    data = json.load(f)

def get_text(m):
    t = m.get("text", "")
    if isinstance(t, list):
        return "".join(x["text"] if isinstance(x, dict) else x for x in t)
    return t

def extract_field(text, *labels):
    for label in labels:
        m = re.search(rf"{re.escape(label)}: ?([^\n]+)", text)
        if m:
            val = m.group(1).strip()
            val = re.sub(r"https?://\S+", "", val).strip()
            return val
    return ""

orders = []

for msg in data["messages"]:
    if msg.get("type") != "message":
        continue
    text = get_text(msg)
    if "Платежная система" not in text and "Сумма оплаты" not in text and "Сумма платежа" not in text:
        continue

    # Дата
    try:
        dt = datetime.fromisoformat(msg["date"])
    except Exception:
        continue

    # Номер заказа
    raw_parts = msg.get("text", [])
    order_num = ""
    if isinstance(raw_parts, list):
        for i, part in enumerate(raw_parts):
            if isinstance(part, dict) and part.get("type") == "bold" and "Заказ" in part.get("text", ""):
                # next part might be the phone/number
                if i + 1 < len(raw_parts):
                    nxt = raw_parts[i + 1]
                    if isinstance(nxt, dict):
                        order_num = nxt.get("text", "")
                    elif isinstance(nxt, str):
                        order_num = nxt.strip()
                break
    if not order_num:
        m2 = re.search(r"Заказ #\s*(\d+)", text)
        if m2:
            order_num = m2.group(1)

    # Сумма
    sum_match = re.search(r"(?:Сумма оплаты|Сумма платежа): (\d+) RSD", text)
    amount = int(sum_match.group(1)) if sum_match else 0

    # Способ оплаты
    payment = extract_field(text, "Платежная система")

    # Имя
    name = extract_field(text, "Имя_и_Фамилия", "Имя", "Name", "ФИО")

    # Телефон
    phone = extract_field(text, "Телефон", "Phone")

    # Telegram
    tg = extract_field(text, "Telegram")
    # также можно взять из ссылки
    if not tg:
        tg_link = re.search(r"https://t\.me/(\S+)", text)
        if tg_link:
            tg = "@" + tg_link.group(1)

    # Город
    city = extract_field(text, "Город", "City")

    # Адрес
    addr = extract_field(text, "Адрес")

    # Доставка
    delivery = extract_field(text, "Доставка")

    # Товары (строки вида "1. Товар, г: ...")
    items = []
    for item in re.findall(r"(?m)^\d+\. (.+)$", text):
        item = item.strip()
        if item:
            items.append(item)

    # Категория по URL
    url_m = re.search(r"https://vkusno\.rs/([^\s#\"<]+)", text)
    category = url_m.group(1) if url_m else ""

    orders.append({
        "order_num": order_num,
        "dt": dt.isoformat(),
        "date": dt.strftime("%d.%m.%Y"),
        "time": dt.strftime("%H:%M"),
        "year": dt.year,
        "month": dt.strftime("%Y-%m"),
        "month_label": dt.strftime("%b %Y"),
        "amount": amount,
        "payment": payment,
        "name": name,
        "phone": phone,
        "telegram": tg,
        "city": city,
        "address": addr,
        "delivery": delivery,
        "items": items,
        "items_str": "; ".join(items),
        "category": category,
    })

print(f"Найдено заказов: {len(orders)}")

# --- Аналитика ---
total = len(orders)
total_sum = sum(o["amount"] for o in orders)
avg_sum = total_sum / total if total else 0
max_order = max(orders, key=lambda o: o["amount"])
positive = [o for o in orders if o["amount"] > 0]
min_order = min(positive, key=lambda o: o["amount"]) if positive else orders[0]

cities = Counter(o["city"] for o in orders if o["city"])

months = defaultdict(lambda: {"count": 0, "sum": 0, "label": ""})
years_stat = defaultdict(lambda: {"count": 0, "sum": 0})
for o in orders:
    m = o["month"]
    months[m]["count"] += 1
    months[m]["sum"] += o["amount"]
    months[m]["label"] = o["month_label"]
    years_stat[o["year"]]["count"] += 1
    years_stat[o["year"]]["sum"] += o["amount"]

months_sorted = sorted(months.items())

all_items = []
for o in orders:
    for item in o["items"]:
        nm = re.match(r"^([^,]+),", item)
        all_items.append(nm.group(1).strip() if nm else item.split(":")[0].strip())

top_products = Counter(all_items).most_common(20)

analytics = {
    "total": total,
    "total_sum": total_sum,
    "avg_sum": round(avg_sum),
    "max_order": max_order,
    "min_order": min_order,
    "cities": cities.most_common(),
    "months": months_sorted,
    "years": sorted(years_stat.items()),
    "top_products": top_products,
}

# ============================================================
# EXCEL
# ============================================================
if HAS_EXCEL:
    print("Генерируем Excel...")
    wb = openpyxl.Workbook()

    # Стили
    HEADER_FILL = PatternFill("solid", fgColor="667EEA")
    HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
    EVEN_FILL = PatternFill("solid", fgColor="F8F9FF")
    AMOUNT_FONT = Font(bold=True, color="2D6A4F")
    CENTER = Alignment(horizontal="center", vertical="top", wrap_text=False)
    TOP = Alignment(vertical="top", wrap_text=True)
    THIN = Side(style="thin", color="DDDDDD")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    def style_header(ws, headers):
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = CENTER
            cell.border = BORDER

    # --- Лист 1: Все заказы ---
    ws1 = wb.active
    ws1.title = "Заказы"
    ws1.freeze_panes = "A2"
    headers = ["№ заказа", "Дата", "Время", "Сумма (RSD)", "Имя", "Город",
               "Телефон", "Telegram", "Адрес", "Доставка", "Оплата", "Товары", "Категория"]
    style_header(ws1, headers)

    for i, o in enumerate(sorted(orders, key=lambda x: x["dt"], reverse=True), 2):
        row = [
            o["order_num"], o["date"], o["time"], o["amount"],
            o["name"], o["city"], o["phone"], o["telegram"],
            o["address"], o["delivery"], o["payment"],
            o["items_str"], o["category"],
        ]
        for col, val in enumerate(row, 1):
            cell = ws1.cell(row=i, column=col, value=val)
            cell.border = BORDER
            if i % 2 == 0:
                cell.fill = EVEN_FILL
            if col == 4:
                cell.font = AMOUNT_FONT
                cell.alignment = CENTER
            elif col in (1, 2, 3):
                cell.alignment = CENTER
            else:
                cell.alignment = TOP

    col_widths = [16, 12, 8, 14, 22, 14, 18, 18, 28, 22, 14, 60, 14]
    for i, w in enumerate(col_widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.row_dimensions[1].height = 22

    # --- Лист 2: По месяцам ---
    ws2 = wb.create_sheet("По месяцам")
    h2 = ["Месяц", "Заказов", "Выручка (RSD)", "Средний чек (RSD)"]
    style_header(ws2, h2)
    for i, (key, m) in enumerate(months_sorted, 2):
        avg = round(m["sum"] / m["count"]) if m["count"] else 0
        for col, val in enumerate([m["label"], m["count"], m["sum"], avg], 1):
            cell = ws2.cell(row=i, column=col, value=val)
            cell.border = BORDER
            cell.alignment = CENTER
            if i % 2 == 0:
                cell.fill = EVEN_FILL
    for i, w in enumerate([16, 12, 18, 18], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # --- Лист 3: По годам ---
    ws3 = wb.create_sheet("По годам")
    h3 = ["Год", "Заказов", "Выручка (RSD)", "Средний чек (RSD)"]
    style_header(ws3, h3)
    for i, (year, y) in enumerate(sorted(years_stat.items()), 2):
        avg = round(y["sum"] / y["count"]) if y["count"] else 0
        for col, val in enumerate([year, y["count"], y["sum"], avg], 1):
            cell = ws3.cell(row=i, column=col, value=val)
            cell.border = BORDER
            cell.alignment = CENTER
            if i % 2 == 0:
                cell.fill = EVEN_FILL
    for i, w in enumerate([10, 12, 18, 18], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    # --- Лист 4: Города ---
    ws4 = wb.create_sheet("По городам")
    h4 = ["Город", "Заказов"]
    style_header(ws4, h4)
    for i, (city, count) in enumerate(cities.most_common(), 2):
        for col, val in enumerate([city, count], 1):
            cell = ws4.cell(row=i, column=col, value=val)
            cell.border = BORDER
            cell.alignment = CENTER
            if i % 2 == 0:
                cell.fill = EVEN_FILL
    for i, w in enumerate([20, 12], 1):
        ws4.column_dimensions[get_column_letter(i)].width = w

    # --- Лист 5: Топ товаров ---
    ws5 = wb.create_sheet("Топ товаров")
    h5 = ["Товар", "Количество заказов"]
    style_header(ws5, h5)
    for i, (prod, count) in enumerate(Counter(all_items).most_common(50), 2):
        for col, val in enumerate([prod, count], 1):
            cell = ws5.cell(row=i, column=col, value=val)
            cell.border = BORDER
            if col == 1:
                cell.alignment = TOP
            else:
                cell.alignment = CENTER
            if i % 2 == 0:
                cell.fill = EVEN_FILL
    ws5.column_dimensions["A"].width = 40
    ws5.column_dimensions["B"].width = 20

    wb.save("orders.xlsx")
    print("Excel сохранён: orders.xlsx")
else:
    print("openpyxl не установлен, Excel пропускаем")

# ============================================================
# HTML DASHBOARD
# ============================================================
print("Генерируем дашборд...")
orders_json = json.dumps(orders, ensure_ascii=False)
analytics_json = json.dumps(analytics, ensure_ascii=False, default=str)

html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>VKUSNO.RS — Аналитика заказов</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f0f2f5;color:#1a1a2e}}
  .header{{background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;padding:24px 32px}}
  .header h1{{font-size:24px;font-weight:700}}
  .header p{{opacity:.85;margin-top:4px;font-size:14px}}
  .container{{max-width:1440px;margin:0 auto;padding:24px 16px}}

  .stats-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:24px}}
  .stat-card{{background:white;border-radius:12px;padding:20px;box-shadow:0 2px 8px rgba(0,0,0,.07)}}
  .stat-card .label{{font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.5px}}
  .stat-card .value{{font-size:26px;font-weight:700;margin-top:6px;color:#667eea}}
  .stat-card .sub{{font-size:12px;color:#666;margin-top:4px}}

  .section{{background:white;border-radius:12px;padding:20px;box-shadow:0 2px 8px rgba(0,0,0,.07);margin-bottom:20px}}
  .section h2{{font-size:15px;font-weight:600;margin-bottom:14px;color:#333}}

  .charts-row{{display:grid;grid-template-columns:2fr 1fr;gap:20px;margin-bottom:20px}}
  @media(max-width:900px){{.charts-row{{grid-template-columns:1fr}}}}

  .year-tabs{{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap}}
  .ytab{{padding:6px 16px;border-radius:20px;border:1px solid #ddd;font-size:13px;cursor:pointer;background:white}}
  .ytab.active{{background:#667eea;color:white;border-color:#667eea}}

  .bar-chart{{display:flex;flex-direction:column;gap:8px}}
  .bar-row{{display:flex;align-items:center;gap:10px;font-size:12px}}
  .bar-label{{width:72px;text-align:right;color:#666;white-space:nowrap;flex-shrink:0}}
  .bar-wrap{{flex:1;background:#f0f2f5;border-radius:4px;height:22px}}
  .bar-fill{{height:100%;border-radius:4px;background:linear-gradient(90deg,#667eea,#764ba2);display:flex;align-items:center;padding-left:8px;color:white;font-size:11px;white-space:nowrap;overflow:hidden;min-width:4px}}
  .bar-val{{margin-left:8px;color:#333;font-weight:600;font-size:11px;white-space:nowrap}}

  .city-list{{display:flex;flex-wrap:wrap;gap:8px}}
  .city-tag{{background:#f0f2f5;border-radius:20px;padding:5px 12px;font-size:13px}}
  .city-tag strong{{color:#667eea}}

  .controls{{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;align-items:center}}
  .search-box{{flex:1;min-width:200px;padding:8px 14px;border:1px solid #ddd;border-radius:8px;font-size:13px;outline:none}}
  .search-box:focus{{border-color:#667eea}}
  select{{padding:8px 12px;border:1px solid #ddd;border-radius:8px;font-size:13px;outline:none;background:white;cursor:pointer}}
  select:focus{{border-color:#667eea}}
  .results-count{{font-size:12px;color:#888;margin-left:auto}}

  .table-wrap{{overflow-x:auto}}
  table{{width:100%;border-collapse:collapse;font-size:12px}}
  thead th{{background:#f8f9ff;padding:9px 10px;text-align:left;font-weight:600;color:#555;border-bottom:2px solid #eee;white-space:nowrap;cursor:pointer;user-select:none}}
  thead th:hover{{background:#eef0ff}}
  thead th.sorted-asc::after{{content:" ▲";font-size:9px}}
  thead th.sorted-desc::after{{content:" ▼";font-size:9px}}
  tbody tr{{border-bottom:1px solid #f5f5f5}}
  tbody tr:hover{{background:#fafbff}}
  td{{padding:9px 10px;vertical-align:top}}
  .td-num{{font-weight:600;color:#667eea;white-space:nowrap}}
  .td-date{{color:#888;white-space:nowrap}}
  .td-amount{{font-weight:700;color:#2d6a4f;white-space:nowrap}}
  .td-name{{white-space:nowrap}}
  .item-list{{list-style:none;padding:0}}
  .item-list li{{padding:1px 0;color:#555}}
  .expand-btn{{font-size:11px;color:#667eea;cursor:pointer;border:none;background:none;padding:2px 0}}
  .td-phone{{white-space:nowrap;font-size:11px}}
  .td-tg{{font-size:11px;color:#667eea}}
  .no-results{{text-align:center;padding:40px;color:#aaa;font-size:15px}}

  .top-products{{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:8px}}
  .product-item{{display:flex;justify-content:space-between;align-items:center;padding:7px 12px;background:#f8f9ff;border-radius:8px;font-size:12px}}
  .product-item .prod-count{{font-weight:700;color:#667eea}}

  .year-stat-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:20px}}
  .year-stat{{background:white;border-radius:10px;padding:16px;box-shadow:0 2px 8px rgba(0,0,0,.07);border-left:4px solid #667eea}}
  .year-stat .y-label{{font-size:20px;font-weight:700;color:#667eea}}
  .year-stat .y-val{{font-size:13px;color:#555;margin-top:4px}}
</style>
</head>
<body>
<div class="header">
  <h1>VKUSNO.RS — Аналитика заказов</h1>
  <p>Telegram-экспорт · авг 2023 — май 2026</p>
</div>
<div class="container">

  <div class="stats-grid" id="statsGrid"></div>

  <div class="year-stat-grid" id="yearGrid"></div>

  <div class="charts-row">
    <div class="section">
      <h2>Выручка по месяцам</h2>
      <div class="year-tabs" id="yearTabs"></div>
      <div class="bar-chart" id="monthChart"></div>
    </div>
    <div class="section">
      <h2>Заказы по городам</h2>
      <div class="city-list" id="cityList"></div>
    </div>
  </div>

  <div class="section">
    <h2>Топ-20 товаров</h2>
    <div class="top-products" id="topProducts"></div>
  </div>

  <div class="section">
    <h2>Все заказы</h2>
    <div class="controls">
      <input class="search-box" id="searchBox" placeholder="Поиск по имени, телефону, адресу, товару..." type="text"/>
      <select id="yearFilter"><option value="">Все годы</option></select>
      <select id="cityFilter"><option value="">Все города</option></select>
      <select id="monthFilter"><option value="">Все месяцы</option></select>
      <span class="results-count" id="resultsCount"></span>
    </div>
    <div class="table-wrap">
      <table id="ordersTable">
        <thead>
          <tr>
            <th data-col="order_num">#</th>
            <th data-col="dt">Дата</th>
            <th data-col="amount">Сумма</th>
            <th data-col="name">Клиент</th>
            <th data-col="city">Город</th>
            <th>Телефон / TG</th>
            <th>Товары</th>
          </tr>
        </thead>
        <tbody id="tableBody"></tbody>
      </table>
      <div class="no-results" id="noResults" style="display:none">Ничего не найдено</div>
    </div>
  </div>
</div>
<script>
const ORDERS = {orders_json};
const ANALYTICS = {analytics_json};

// Stats
const statsGrid = document.getElementById('statsGrid');
[
  {{label:'Всего заказов', value:ANALYTICS.total.toLocaleString('ru'), sub:'авг 2023 — май 2026'}},
  {{label:'Общая выручка', value:ANALYTICS.total_sum.toLocaleString('ru')+' RSD', sub:''}},
  {{label:'Средний чек', value:ANALYTICS.avg_sum.toLocaleString('ru')+' RSD', sub:''}},
  {{label:'Макс. заказ', value:ANALYTICS.max_order.amount.toLocaleString('ru')+' RSD', sub:ANALYTICS.max_order.name}},
  {{label:'Мин. заказ', value:ANALYTICS.min_order.amount.toLocaleString('ru')+' RSD', sub:ANALYTICS.min_order.name}},
  {{label:'Городов', value:ANALYTICS.cities.length, sub:''}},
].forEach(s=>{{
  statsGrid.innerHTML+=`<div class="stat-card"><div class="label">${{s.label}}</div><div class="value">${{s.value}}</div><div class="sub">${{s.sub}}</div></div>`;
}});

// Year stats
const yearGrid = document.getElementById('yearGrid');
ANALYTICS.years.forEach(([year, y])=>{{
  const avg = y.count ? Math.round(y.sum/y.count) : 0;
  yearGrid.innerHTML+=`<div class="year-stat"><div class="y-label">${{year}}</div><div class="y-val">${{y.count}} заказов</div><div class="y-val">${{y.sum.toLocaleString('ru')}} RSD</div><div class="y-val">~${{avg.toLocaleString('ru')}} RSD / заказ</div></div>`;
}});

// Month chart with year filter
const allYears = [...new Set(ANALYTICS.months.map(([k])=>k.slice(0,4)))].sort();
let activeYear = 'all';
const yearTabs = document.getElementById('yearTabs');
const monthChart = document.getElementById('monthChart');

function renderMonthChart() {{
  const filtered = activeYear === 'all' ? ANALYTICS.months : ANALYTICS.months.filter(([k])=>k.startsWith(activeYear));
  const maxSum = Math.max(...filtered.map(([,m])=>m.sum), 1);
  monthChart.innerHTML = filtered.map(([key, m])=>{{
    const pct = Math.max(4, (m.sum/maxSum)*100);
    return `<div class="bar-row">
      <div class="bar-label">${{m.label}}</div>
      <div class="bar-wrap"><div class="bar-fill" style="width:${{pct}}%">${{m.count}} зак.</div></div>
      <div class="bar-val">${{m.sum.toLocaleString('ru')}} RSD</div>
    </div>`;
  }}).join('');
}}

[['all','Все'],...allYears.map(y=>[y,y])].forEach(([val, label])=>{{
  const btn = document.createElement('button');
  btn.className = 'ytab' + (val==='all' ? ' active' : '');
  btn.textContent = label;
  btn.onclick = ()=>{{
    activeYear = val;
    document.querySelectorAll('.ytab').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    renderMonthChart();
  }};
  yearTabs.appendChild(btn);
}});
renderMonthChart();

// Cities
const cityList = document.getElementById('cityList');
ANALYTICS.cities.forEach(([city, count])=>{{
  cityList.innerHTML+=`<div class="city-tag">${{city}} <strong>${{count}}</strong></div>`;
}});

// Top products
const topProducts = document.getElementById('topProducts');
ANALYTICS.top_products.forEach(([prod, count])=>{{
  topProducts.innerHTML+=`<div class="product-item"><span>${{prod}}</span><span class="prod-count">${{count}}</span></div>`;
}});

// Filters
const yearFilter = document.getElementById('yearFilter');
allYears.forEach(y=>{{ yearFilter.innerHTML+=`<option value="${{y}}">${{y}}</option>`; }});

const cityFilter = document.getElementById('cityFilter');
ANALYTICS.cities.forEach(([city])=>{{ cityFilter.innerHTML+=`<option value="${{city}}">${{city}}</option>`; }});

const monthFilter = document.getElementById('monthFilter');
ANALYTICS.months.forEach(([key, m])=>{{ monthFilter.innerHTML+=`<option value="${{key}}">${{m.label}}</option>`; }});

// Table
let sortCol='dt', sortDir=-1, filtered=[...ORDERS];

function renderTable() {{
  const tbody = document.getElementById('tableBody');
  tbody.innerHTML = '';
  document.getElementById('noResults').style.display = filtered.length ? 'none' : '';
  document.getElementById('resultsCount').textContent = filtered.length < ORDERS.length ? `Показано: ${{filtered.length}} из ${{ORDERS.length}}` : `Всего: ${{ORDERS.length}}`;
  filtered.forEach(o=>{{
    const max3 = o.items.slice(0,3).map(i=>`<li>${{i}}</li>`).join('');
    const extra = o.items.length > 3 ? `<li><button class="expand-btn" onclick="expandItems(this)" data-order="${{o.order_num}}">+ ещё ${{o.items.length-3}}</button></li>` : '';
    tbody.innerHTML+=`<tr>
      <td class="td-num">${{o.order_num}}</td>
      <td class="td-date">${{o.date}}<br><small style="color:#bbb">${{o.time}}</small></td>
      <td class="td-amount">${{o.amount.toLocaleString('ru')}} RSD</td>
      <td class="td-name">${{o.name||'—'}}</td>
      <td>${{o.city}}</td>
      <td class="td-phone">${{o.phone}}<br><span class="td-tg">${{o.telegram}}</span></td>
      <td><ul class="item-list">${{max3}}${{extra}}</ul></td>
    </tr>`;
  }});
}}

function expandItems(btn) {{
  const o = ORDERS.find(x=>x.order_num===btn.dataset.order);
  if (o) btn.closest('ul').innerHTML = o.items.map(i=>`<li>${{i}}</li>`).join('');
}}

function applyFilters() {{
  const q = document.getElementById('searchBox').value.toLowerCase();
  const city = cityFilter.value;
  const month = monthFilter.value;
  const year = yearFilter.value;
  filtered = ORDERS.filter(o=>{{
    if (year && o.year != year) return false;
    if (city && o.city !== city) return false;
    if (month && o.month !== month) return false;
    if (q) {{
      const hay = [o.name,o.phone,o.telegram,o.address,...o.items].join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }}
    return true;
  }});
  filtered.sort((a,b)=>{{
    let av=a[sortCol], bv=b[sortCol];
    if (sortCol==='amount'){{av=+av;bv=+bv;}}
    return av<bv ? sortDir : av>bv ? -sortDir : 0;
  }});
  renderTable();
}}

document.getElementById('searchBox').addEventListener('input', applyFilters);
[cityFilter, monthFilter, yearFilter].forEach(el=>el.addEventListener('change', applyFilters));

document.querySelectorAll('thead th[data-col]').forEach(th=>{{
  th.addEventListener('click',()=>{{
    const col=th.dataset.col;
    if (sortCol===col) sortDir*=-1; else {{sortCol=col;sortDir=-1;}}
    document.querySelectorAll('thead th').forEach(h=>h.classList.remove('sorted-asc','sorted-desc'));
    th.classList.add(sortDir===-1?'sorted-desc':'sorted-asc');
    applyFilters();
  }});
}});

document.querySelector('thead th[data-col="dt"]').classList.add('sorted-desc');
applyFilters();
</script>
</body>
</html>
"""

with open("orders_dashboard.html", "w", encoding="utf-8") as f:
    f.write(html)

print("Готово! orders_dashboard.html + orders.xlsx")
