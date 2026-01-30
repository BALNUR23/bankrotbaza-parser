import time
import os
import re
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup


class BankrotBazaParser:
    def __init__(self):
        chrome_options = Options()
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")

        script_dir = os.path.dirname(os.path.abspath(__file__))
        profile_path = os.path.join(script_dir, "selenium_profile")
        chrome_options.add_argument(f"--user-data-dir={profile_path}")

        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )
        self.data = []

    def get_links(self, pages_count=10):
        links = []
        base_url = "https://bankrotbaza.ru/c/квартиры"

        for p in range(1, pages_count + 1):
            url = f"{base_url}?page={p}"
            print(f"--- Сканирование списка лотов: страница {p} ---")

            self.driver.get(url)
            time.sleep(4)

            soup = BeautifulSoup(self.driver.page_source, "html.parser")
            for a in soup.find_all("a", href=True):
                href = a["href"]

                if "/lot/" in href and not any(
                    x in href for x in ["login", "register", "map", "favorit", "c/", "nedvizhimost"]
                ):
                    full_url = "https://bankrotbaza.ru" + href if href.startswith("/") else href
                    if full_url not in links:
                        links.append(full_url)

        return links

    def get_smart_val(self, soup, label):
        """Улучшенный поиск значения по метке (фиксит проблему Н/Д)"""
        target = soup.find(string=re.compile(label, re.IGNORECASE))
        if not target:
            return "Н/Д"

        parent = target.find_parent()
        if not parent:
            return "Н/Д"

        # 1) Соседний блок (часто верстка "лейбл | значение")
        next_sib = parent.find_next_sibling()
        if next_sib:
            txt = next_sib.get_text(strip=True)
            return txt if txt else "Н/Д"

        # 2) Внутри того же блока через разделитель
        text = parent.get_text(separator="|", strip=True)
        parts = [p.strip() for p in text.split("|") if p.strip()]
        for i, part in enumerate(parts):
            if re.search(label, part, re.IGNORECASE) and i + 1 < len(parts):
                return parts[i + 1]

        # 3) У родителя (если вложенная верстка)
        grand = parent.parent
        if grand:
            g_text = grand.get_text(separator="|", strip=True)
            g_parts = [p.strip() for p in g_text.split("|") if p.strip()]
            for i, part in enumerate(g_parts):
                if re.search(label, part, re.IGNORECASE) and i + 1 < len(g_parts):
                    return g_parts[i + 1]

        return "Н/Д"

    def parse_lot_page(self, url):
        self.driver.get(url)

        # Скроллим, чтобы подгрузка сработала
        self.driver.execute_script("window.scrollTo(0, 1000);")
        time.sleep(2.5)

        soup = BeautifulSoup(self.driver.page_source, "html.parser")
        title = soup.find("h1").get_text(strip=True) if soup.find("h1") else "Без названия"

        lot_num_match = re.search(r"№?\s?(\d+)", title)

        docs = []
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if ".pdf" in href.lower() or ".zip" in href.lower():
                if href.startswith("/"):
                    href = "https://bankrotbaza.ru" + href
                docs.append(href)

        item = {
            "Номер лота": lot_num_match.group(1) if lot_num_match else self.get_smart_val(soup, "Лот"),
            "Название": title,
            "Адрес объекта": self.get_smart_val(soup, "Адрес"),
            "Начальная цена": self.get_smart_val(soup, "Начальная цена"),
            "Шаг аукциона": self.get_smart_val(soup, "Шаг повышения"),
            "Размер задатка": self.get_smart_val(soup, "Задаток"),
            "Дата начала": self.get_smart_val(soup, "Прием заявок с"),
            "Дата окончания": self.get_smart_val(soup, "Прием заявок до"),
            "Должник": f"{self.get_smart_val(soup, 'Наименование')} (ИНН: {self.get_smart_val(soup, 'ИНН')})",
            "Документация": "\n".join(docs) if docs else "Н/Д",
            "Ссылка": url
        }
        return item

    def run(self, max_pages=10):
        try:
            links = self.get_links(max_pages)
            print(f"Всего лотов: {len(links)}")

            for i, link in enumerate(links, start=1):
                try:
                    data = self.parse_lot_page(link)
                    self.data.append(data)

                    preview = data["Название"][:35]
                    price = data["Начальная цена"]
                    print(f"[{i}/{len(links)}] Собрано: {preview}... | Цена: {price}")

                    if price == "Н/Д":
                        print("Предупреждение: Данные не найдены, делаю паузу...")
                        time.sleep(5)

                except Exception as e:
                    print(f"Ошибка в {link}: {e}")

            if self.data:
                self.save_to_excel()

        finally:
            self.driver.quit()

    def save_to_excel(self):
        import pandas as pd
        import os

        filename = "bankrot_data.xlsx"
        df = pd.DataFrame(self.data)

        order = [
            "Номер лота", "Название", "Адрес объекта",
            "Начальная цена", "Шаг аукциона", "Размер задатка",
            "Дата начала", "Дата окончания",
            "Должник", "Документация", "Ссылка"
        ]
        order = [c for c in order if c in df.columns]
        df = df[order]

        with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Лоты")
            wb = writer.book
            ws = writer.sheets["Лоты"]

            # ---------- Форматы ----------
            header_fmt = wb.add_format({
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "bg_color": "#EDF2F7",  # светло-серо-голубой
                "font_color": "#000000"
            })

            cell_fmt = wb.add_format({
                "border": 1,
                "valign": "vcenter"
            })

            zebra_fmt = wb.add_format({
                "border": 1,
                "valign": "vcenter",
                "bg_color": "#F7FAFC"  # очень лёгкая зебра
            })

            wrap_fmt = wb.add_format({
                "border": 1,
                "valign": "top",
                "text_wrap": True
            })

            link_fmt = wb.add_format({
                "border": 1,
                "valign": "vcenter",
                "font_color": "#0563C1",
                "underline": 1
            })

            # ---------- Шапка ----------
            ws.set_row(0, 26)
            for col, name in enumerate(df.columns):
                ws.write(0, col, name, header_fmt)

            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(df), len(df.columns) - 1)
            ws.hide_gridlines(2)

            # ---------- Ширины ----------
            width_map = {
                "Номер лота": 10,
                "Название": 48,
                "Адрес объекта": 48,
                "Начальная цена": 16,
                "Шаг аукциона": 16,
                "Размер задатка": 18,
                "Дата начала": 14,
                "Дата окончания": 14,
                "Должник": 42,
                "Документация": 22,
                "Ссылка": 14
            }

            for i, col in enumerate(df.columns):
                ws.set_column(i, i, width_map.get(col, 18))

            # ---------- Высота строк ----------
            for r in range(1, len(df) + 1):
                ws.set_row(r, 22)

            # ---------- Запись данных ----------
            for r in range(len(df)):
                row_fmt = zebra_fmt if r % 2 else cell_fmt
                for c in range(len(df.columns)):
                    val = df.iat[r, c]
                    excel_row = r + 1
                    col_name = df.columns[c]

                    if col_name == "Ссылка" and isinstance(val, str) and val.startswith("http"):
                        ws.write_url(excel_row, c, val, link_fmt, string="Открыть")
                    elif col_name in ["Название", "Адрес объекта", "Документация"]:
                        ws.write(excel_row, c, val, wrap_fmt)
                    else:
                        ws.write(excel_row, c, val, row_fmt)

        print(f"\nФАЙЛ СОЗДАН: {os.path.abspath(filename)}")


if __name__ == "__main__":
    parser = BankrotBazaParser()
    parser.run(max_pages=1)  # 10 страниц
