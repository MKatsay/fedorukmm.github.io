from excel_to_html import excel_to_html

html = excel_to_html("КДЛ доп.xls")

with open("kdl_dop_copy.html", "w", encoding="utf-8") as f:
    f.write(html)
