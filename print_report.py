"""
print_report.py — формирует HTML-отчёт и открывает диалог печати в браузере
"""
import os, tempfile, webbrowser
from datetime import date

DOC_TYPE_LABELS = {
    "purchase":     "Закупка",
    "sale":         "Продажа",
    "movement_in":  "Перемещение (приход)",
    "movement_out": "Перемещение (расход)",
}

CSS = """
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, sans-serif; font-size: 11pt; color: #1a1a2e; padding: 20mm; }
  h1 { font-size: 18pt; color: #3730a3; margin-bottom: 4px; }
  .meta { font-size: 9pt; color: #666; margin-bottom: 16px; }
  table { width: 100%; border-collapse: collapse; margin-top: 8px; }
  th { background: #4f46e5; color: #fff; padding: 7px 10px;
       text-align: left; font-size: 10pt; }
  td { padding: 5px 10px; border-bottom: 1px solid #e0e0f0; font-size: 10pt; }
  tr:nth-child(even) td { background: #f5f3ff; }
  tr:hover td { background: #ede9fe; }
  .total-row td { font-weight: bold; background: #ede9fe !important;
                  border-top: 2px solid #4f46e5; }
  .num { text-align: right; }
  .footer { margin-top: 20px; font-size: 9pt; color: #888;
            border-top: 1px solid #ddd; padding-top: 8px; }
  @media print {
    body { padding: 10mm; }
    .no-print { display: none; }
    tr { page-break-inside: avoid; }
  }
  .print-btn {
    position: fixed; top: 16px; right: 20px;
    background: #4f46e5; color: white; border: none;
    padding: 10px 22px; border-radius: 6px; cursor: pointer;
    font-size: 12pt; font-weight: bold; box-shadow: 0 2px 8px rgba(0,0,0,.2);
  }
  .print-btn:hover { background: #3730a3; }
</style>
"""


def _html_page(title, subtitle, headers, rows, total_col_idx=None):
    today = date.today().strftime("%d.%m.%Y")
    rows_html = ""
    total = 0
    for i, row in enumerate(rows):
        cells = ""
        for j, val in enumerate(row):
            cls = " class='num'" if isinstance(val, (int, float)) else ""
            cells += f"<td{cls}>{val}</td>"
            if total_col_idx is not None and j == total_col_idx:
                try:
                    total += float(val)
                except:
                    pass
        rows_html += f"<tr>{cells}</tr>\n"

    total_row = ""
    if total_col_idx is not None and rows:
        empty = "".join(
            f"<td></td>" for _ in range(total_col_idx)
        )
        total_row = f"""
        <tr class='total-row'>
          {empty}
          <td class='num'><b>ИТОГО: {total:,.0f}</b></td>
          {''.join(f"<td></td>" for _ in range(len(headers) - total_col_idx - 1))}
        </tr>"""

    ths = "".join(f"<th>{h}</th>" for h in headers)
    count = len(rows)

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <title>{title}</title>
  {CSS}
</head>
<body>
  <button class="print-btn no-print" onclick="window.print()">🖨 Печать</button>
  <h1>{title}</h1>
  <p class="meta">{subtitle} &nbsp;·&nbsp; Дата формирования: {today} &nbsp;·&nbsp; Строк: {count}</p>
  <table>
    <thead><tr>{ths}</tr></thead>
    <tbody>
      {rows_html}
      {total_row}
    </tbody>
  </table>
  <div class="footer">Складской учёт · {today}</div>
  <script>
    // Авто-печать если открыт из приложения
    if (window.location.search.includes('autoprint')) {{
      window.onload = () => setTimeout(() => window.print(), 400);
    }}
  </script>
</body>
</html>"""


def _open_html(html: str):
    """Сохраняем во временный файл и открываем в браузере."""
    fd, path = tempfile.mkstemp(suffix=".html", prefix="warehouse_report_")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(html)
    webbrowser.open(f"file://{path}?autoprint=1")
    return path


# ── Публичные функции ──────────────────────────────────────────────────────

def print_purchases(rows):
    table_rows = [(r["id"], r["date"], r["supplier"],
                   r["nomenclature"], r["unit"], r["warehouse"], r["quantity"])
                  for r in rows]
    html = _html_page(
        "Отчёт по закупкам", "Все поступления товара",
        ["ID", "Дата", "Поставщик", "Номенклатура", "Ед.", "Склад", "Кол-во"],
        table_rows, total_col_idx=6
    )
    return _open_html(html)


def print_sales(rows):
    table_rows = [(r["id"], r["date"], r["customer"],
                   r["nomenclature"], r["unit"], r["warehouse"], r["quantity"])
                  for r in rows]
    html = _html_page(
        "Отчёт по продажам", "Все реализации товара",
        ["ID", "Дата", "Покупатель", "Номенклатура", "Ед.", "Склад", "Кол-во"],
        table_rows, total_col_idx=6
    )
    return _open_html(html)


def print_movements(rows):
    table_rows = [(r["id"], r["date"], r["from_warehouse"], r["to_warehouse"],
                   r["nomenclature"], r["unit"], r["quantity"])
                  for r in rows]
    html = _html_page(
        "Отчёт по перемещениям", "Все перемещения между складами",
        ["ID", "Дата", "Со склада", "На склад", "Номенклатура", "Ед.", "Кол-во"],
        table_rows, total_col_idx=6
    )
    return _open_html(html)


def print_stock(rows):
    table_rows = [(r["nomenclature"], r["unit"], r["warehouse"], r["balance"])
                  for r in rows]
    html = _html_page(
        "Остатки на складах",
        f"Текущие остатки по всем складам на {date.today().strftime('%d.%m.%Y')}",
        ["Номенклатура", "Ед.", "Склад", "Остаток"],
        table_rows, total_col_idx=3
    )
    return _open_html(html)


def print_history(rows):
    table_rows = [
        (r["date"],
         DOC_TYPE_LABELS.get(r["document_type"], r["document_type"]),
         r["nomenclature"], r["unit"], r["warehouse"], r["quantity"])
        for r in rows
    ]
    html = _html_page(
        "История движений товара", "Журнал всех складских операций",
        ["Дата", "Тип документа", "Номенклатура", "Ед.", "Склад", "Кол-во"],
        table_rows, total_col_idx=5
    )
    return _open_html(html)
