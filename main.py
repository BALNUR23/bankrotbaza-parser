import os
import re
import time
from datetime import datetime

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

BASE = "https://bankrotbaza.ru"


class BankrotBazaParser:
    def __init__(self, headless: bool = False):
        chrome_options = Options()
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--start-maximized")
        if headless:
            chrome_options.add_argument("--headless=new")

        script_dir = os.path.dirname(os.path.abspath(__file__))
        profile_path = os.path.join(script_dir, "selenium_profile")
        chrome_options.add_argument(f"--user-data-dir={profile_path}")

        # ✅ стабильнее на macOS, чем webdriver_manager
        self.driver = webdriver.Chrome(options=chrome_options)
        self.data = []

    # ---------- utils ----------
    def _abs_url(self, href: str) -> str:
        if not href:
            return ""
        return (BASE + href) if href.startswith("/") else href

    def _soup(self) -> BeautifulSoup:
        return BeautifulSoup(self.driver.page_source, "html.parser")

    def _clean(self, s: str) -> str:
        return re.sub(r"\s+", " ", (s or "")).strip()

    def lot_id_from_url(self, url: str) -> str:
        # /lot/123456 -> 123456
        m = re.search(r"/lot/(\d+)", url)
        return m.group(1) if m else ""

    def get_val(self, soup: BeautifulSoup, label: str) -> str:
        """
        ✅ Возвращает "" (пусто), если нет значения.
        """
        target = soup.find(string=re.compile(label, re.IGNORECASE))
        if not target:
            return ""

        parent = target.find_parent()
        if not parent:
            return ""

        # 1) значение в соседнем блоке
        next_sib = parent.find_next_sibling()
        if next_sib:
            v = self._clean(next_sib.get_text(" ", strip=True))
            if v:
                return v

        # 2) "лейбл | значение" внутри блока
        text = self._clean(parent.get_text("|", strip=True))
        parts = [p.strip() for p in text.split("|") if p.strip()]
        for i, part in enumerate(parts):
            if re.search(label, part, re.IGNORECASE) and i + 1 < len(parts):
                return parts[i + 1].strip()

        # 3) выше по дереву
        grand = parent.parent
        if grand:
            gtext = self._clean(grand.get_text("|", strip=True))
            gparts = [p.strip() for p in gtext.split("|") if p.strip()]
            for i, part in enumerate(gparts):
                if re.search(label, part, re.IGNORECASE) and i + 1 < len(gparts):
                    return gparts[i + 1].strip()

        return ""

    # ---------- links ----------
    def get_links(self, pages_count: int = 1) -> list:
        links = []
        base_url = f"{BASE}/c/квартиры"

        for p in range(1, pages_count + 1):
            url = f"{base_url}?page={p}"
            print(f"--- Поиск лотов на странице {p} ---")

            self.driver.get(url)
            time.sleep(3)

            soup = self._soup()
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if "/lot/" not in href:
                    continue
                if any(x in href for x in ["login", "register", "map", "/c/", "nedvizhimost"]):
                    continue

                full_url = self._abs_url(href)
                if full_url and full_url not in links:
                    links.append(full_url)

        return links

    # ---------- parse ----------
    def parse_lot_page(self, url: str) -> dict:
        print(f"Обработка: {url}")
        self.driver.get(url)

        self.driver.execute_script("window.scrollTo(0, 1500);")
        time.sleep(2)

        soup = self._soup()

        title_el = soup.find("h1")
        title_text = self._clean(title_el.get_text(" ", strip=True)) if title_el else ""

        # ✅ 1) берём “номер” из URL (хоть как-то стабильно)
        lot_number = self.lot_id_from_url(url)

        # ✅ 2) если на странице реально есть “Номер лота/Лот” — перезаписываем
        # (иногда у них именно так называется)
        real_lot = (
            self.get_val(soup, "Номер лота") or
            self.get_val(soup, "Лот") or
            self.get_val(soup, "Лот №")
        )
        if real_lot:
            # если там текст типа "12345" — ок, если мусор — оставим url-id
            m = re.search(r"(\d{3,})", real_lot)
            if m:
                lot_number = m.group(1)

        # адрес/цены/даты
        address = self.get_val(soup, "Адрес")
        start_price = self.get_val(soup, "Начальная цена")
        step = self.get_val(soup, "Шаг повышения") or self.get_val(soup, "Шаг аукциона") or self.get_val(soup, "Шаг")
        deposit = self.get_val(soup, "Задаток") or self.get_val(soup, "Размер задатка")

        date_start = self.get_val(soup, "Прием заявок с") or self.get_val(soup, "Приём заявок с")
        date_end = self.get_val(soup, "Прием заявок до") or self.get_val(soup, "Приём заявок до")

        status = self.get_val(soup, "Статус") or "Торги объявлены"

        # должник
        debtor_name = self.get_val(soup, "Наименование") or self.get_val(soup, "Должник")
        debtor_inn = self.get_val(soup, "ИНН")
        if debtor_name and debtor_inn:
            debtor_info = f"{debtor_name} (ИНН: {debtor_inn})"
        elif debtor_name:
            debtor_info = debtor_name
        elif debtor_inn:
            debtor_info = f"ИНН: {debtor_inn}"
        else:
            debtor_info = ""

        # описание
        desc_block = soup.select_one(".lot-description, .lot-card__description")
        description = self._clean(desc_block.get_text(" ", strip=True)) if desc_block else title_text
        description_short = (description[:500] + "...") if len(description) > 500 else description

        # документы
        doc_links = []
        for a in soup.find_all("a", href=True):
            href = a["href"]
            low = href.lower()
            if any(ext in low for ext in [".pdf", ".zip", ".doc", ".docx"]):
                doc_links.append(self._abs_url(href))
        docs_str = "\n".join(sorted(set(doc_links))) if doc_links else ""

        # DEBUG: чтобы ты видела, что номер реально извлекается
        print(f"  -> lot_number = {lot_number or '(пусто)'}")

        return {
            "Номер лота": lot_number,
            "Название/Описание": title_text,
            "Адрес объекта": address,
            "Начальная цена": start_price,
            "Шаг аукциона": step,
            "Размер задатка": deposit,
            "Начало приема заявок": date_start,
            "Окончание приема заявок": date_end,
            "Статус аукциона": status,
            "Информация о должнике": debtor_info,
            "Ссылка на документацию": docs_str,
            "Полное описание": description_short,
            "Ссылка на лот": url
        }

    # ---------- run ----------
    def run(self, max_pages: int = 1):
        try:
            links = self.get_links(max_pages)
            print(f"Всего найдено лотов: {len(links)}")

            for i, link in enumerate(links, start=1):
                try:
                    res = self.parse_lot_page(link)
                    self.data.append(res)
                    print(f"[{i}/{len(links)}] ✅ OK")
                except Exception as e:
                    print(f"Ошибка в лоте {link}: {e}")

            if self.data:
                self.save_to_excel()
        finally:
            self.driver.quit()

    # ---------- excel ----------
    def save_to_excel(self):
        # ✅ новое имя, чтобы ты точно открыла новый файл
        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"full_bankrot_report_{ts}.xlsx"

        df = pd.DataFrame(self.data).fillna("").replace("Н/Д", "")

        order = [
            "Номер лота",
            "Название/Описание",
            "Адрес объекта",
            "Начальная цена",
            "Шаг аукциона",
            "Размер задатка",
            "Начало приема заявок",
            "Окончание приема заявок",
            "Статус аукциона",
            "Информация о должнике",
            "Ссылка на документацию",
            "Полное описание",
            "Ссылка на лот",
        ]
        df = df[[c for c in order if c in df.columns]]

        with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Данные торгов")
            wb = writer.book
            ws = writer.sheets["Данные торгов"]

            # ---- форматы ----
            header_fmt = wb.add_format({
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "bg_color": "#1F4E79",
                "font_color": "white"
            })
            cell_fmt = wb.add_format({"border": 1, "valign": "top"})
            zebra_fmt = wb.add_format({"border": 1, "valign": "top", "bg_color": "#E9F0FA"})
            wrap_fmt = wb.add_format({"border": 1, "valign": "top", "text_wrap": True})
            link_fmt = wb.add_format({"border": 1, "valign": "top", "font_color": "#0563C1", "underline": 1})

            # ---- шапка ----
            ws.set_row(0, 28)
            for col, name in enumerate(df.columns):
                ws.write(0, col, name, header_fmt)

            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(df), len(df.columns) - 1)
            ws.hide_gridlines(2)

            # ---- ширины ----
            widths = {
                "Номер лота": 12,
                "Название/Описание": 46,
                "Адрес объекта": 46,
                "Начальная цена": 18,
                "Шаг аукциона": 16,
                "Размер задатка": 16,
                "Начало приема заявок": 22,
                "Окончание приема заявок": 22,
                "Статус аукциона": 18,
                "Информация о должнике": 42,
                "Ссылка на документацию": 36,
                "Полное описание": 60,
                "Ссылка на лот": 14
            }
            for i, col in enumerate(df.columns):
                ws.set_column(i, i, widths.get(col, 20))

            # ---- данные (зебра + переносы + ссылки) ----
            for r in range(len(df)):
                excel_row = r + 1
                ws.set_row(excel_row, 55)
                row_fmt = zebra_fmt if (r % 2 == 1) else cell_fmt

                for c, col_name in enumerate(df.columns):
                    val = df.iat[r, c]

                    if col_name == "Ссылка на лот" and isinstance(val, str) and val.startswith("http"):
                        ws.write_url(excel_row, c, val, link_fmt, string="Открыть")
                    elif col_name in ["Название/Описание", "Адрес объекта", "Полное описание", "Ссылка на документацию", "Информация о должнике"]:
                        ws.write(excel_row, c, val, wrap_fmt)
                    else:
                        ws.write(excel_row, c, val, row_fmt)

        print(f"\n✅ EXCEL СОЗДАН: {os.path.abspath(filename)}")


if __name__ == "__main__":
    parser = BankrotBazaParser(headless=False)
    parser.run(max_pages=10)
