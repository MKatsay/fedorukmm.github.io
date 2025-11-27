import pandas as pd
import re
import os

# ----------------------------
# 1. Наши цены
# ----------------------------
def parse_our_prices(file_path):
    df = pd.read_excel(file_path, header=None, dtype=str)
    df.fillna('', inplace=True)

    sections = []
    current_section = None

    for idx, row in df.iterrows():
        cell0 = str(row[0]).strip()

        # Определяем заголовок секции: не пустой, не код услуги (не начинается с A/B/C + цифры), и не "Код НМУ"
        if cell0 and cell0 != '' and not re.match(r'^[A-Z]\d', cell0) and 'Код НМУ' not in cell0 and 'Наименование услуги' not in cell0:
            # Проверим, что это действительно заголовок, а не просто текст в услуге
            # Например, если в строке 4 непустых ячейки — это, скорее, услуга
            non_empty = sum(1 for c in row if str(c).strip() != '')
            if non_empty == 1:  # только первый столбец заполнен → это заголовок
                current_section = cell0.upper()
                sections.append({'name': current_section, 'rows': []})
                continue

        # Пропускаем строки с заголовками колонок
        if 'Код НМУ' in str(row[0]) or 'Наименование услуги' in str(row[2]):
            continue

        # Пропускаем пустые строки
        if all(str(c).strip() == '' for c in row):
            continue

        # Пропускаем примечания внизу
        if 'Примечания:' in cell0 or '* Повторным приемом' in cell0:
            continue

        # Добавляем строку в текущую секцию
        if current_section is not None:
            code_nmu = str(row[0]).strip() if pd.notna(row[0]) else ''
            code_lu = str(row[1]).strip() if pd.notna(row[1]) else ''
            service = str(row[2]).strip().replace('\n', ' ') if pd.notna(row[2]) else ''
            price = str(row[3]).strip().replace('\n', ' ') if pd.notna(row[3]) else ''

            # Убеждаемся, что это строка услуги (есть хотя бы код и цена)
            if code_nmu and price and re.search(r'\d', price):
                sections[-1]['rows'].append({
                    'code_nmu': code_nmu,
                    'code_lu': code_lu,
                    'service': service,
                    'price': price
                })

    return sections



# ----------------------------
# Генерация HTML
# ----------------------------
def gen_our_html(sections):
    html = '''<div class="our-prices-header">
  <a href="https://fedorukmm.ru/" target="_blank">
    <img src="https://static.tildacdn.com/tild6164-3435-4336-a161-373866376663/logozaru.jpg" alt="МЦ Федорук М.М.">
  </a>
  <div><h2>Цены на услуги</h2><p>Уточняйте актуальность у администратора</p></div>
</div>
<div class="price-search">
  <input type="text" placeholder="Поиск..." onkeyup="filterTable(this, 'our')">
  <button onclick="clearSearch('our')">Очистить</button>
</div>'''
    for sec in sections:
        html += f'<div class="section-header"><h3>{sec["name"]}</h3></div>\n<table class="tpl-table">\n<thead><tr><th>Код НМУ</th><th>Код ЛУ</th><th>Услуга</th><th>Цена</th></tr></thead>\n<tbody>\n'
        for r in sec['rows']:
            html += f'  <tr><td>{r["code_nmu"]}</td><td>{r["code_lu"]}</td><td>{r["service"]}</td><td>{r["price"]}</td></tr>\n'
        html += '</tbody>\n</table>\n'
    return html


# ----------------------------
# Запуск
# ----------------------------
if __name__ == '__main__':
    our_file = "Прейскурант 26.11.2025 (для сайта).xlsx"

    # 1. Наши цены
    our = parse_our_prices(our_file)
    with open("our_prices.html", "w", encoding="utf-8") as f:
        f.write(gen_our_html(our))



    print("✅ Готово! Созданы 1 файл:")
    print("  • our_prices.html")
