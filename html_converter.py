import re
from html import unescape
from bs4 import BeautifulSoup


class HTMLToTextConverter:
    """Конвертер HTML в читаемый текст с сохранением структуры"""

    @staticmethod
    def convert(html_content):
        """
        Конвертирует HTML в читаемый текст с сохранением структуры

        Args:
            html_content (str): HTML контент для конвертации

        Returns:
            str: Читаемый текст с сохраненной структурой
        """
        if not html_content:
            return ""

        # Декодируем HTML entities
        cleaned_html = unescape(html_content)

        # Используем BeautifulSoup для парсинга
        soup = BeautifulSoup(cleaned_html, "html.parser")

        # Удаляем ненужные теги
        for unwanted in soup(["script", "style", "meta", "link"]):
            unwanted.decompose()

        # Обрабатываем заголовки
        HTMLToTextConverter._process_headings(soup)

        # Обрабатываем списки
        HTMLToTextConverter._process_lists(soup)

        # Обрабатываем таблицы
        HTMLToTextConverter._process_tables(soup)

        # Обрабатываем ссылки
        HTMLToTextConverter._process_links(soup)

        # Обрабатываем параграфы и переносы
        HTMLToTextConverter._process_paragraphs(soup)

        # Получаем текст и финализируем форматирование
        text = soup.get_text()
        text = HTMLToTextConverter._finalize_formatting(text)

        return text.strip()

    @staticmethod
    def _process_headings(soup):
        """Обрабатывает заголовки"""
        heading_map = {
            "h1": "====== ",
            "h2": "===== ",
            "h3": "==== ",
            "h4": "=== ",
            "h5": "== ",
            "h6": "= ",
        }

        for tag, prefix in heading_map.items():
            for heading in soup.find_all(tag):
                heading_text = heading.get_text().strip()
                if heading_text:
                    heading.replace_with(
                        f"\n\n{prefix}{heading_text.upper()}\n{prefix.replace(' ', '=')}\n"
                    )

    @staticmethod
    def _process_lists(soup):
        """Обрабатывает упорядоченные и неупорядоченные списки"""
        # Неупорядоченные списки
        for ul in soup.find_all("ul"):
            list_items = []
            for li in ul.find_all("li"):
                item_text = li.get_text().strip()
                if item_text:
                    list_items.append(f"  • {item_text}")

            if list_items:
                ul.replace_with("\n" + "\n".join(list_items) + "\n")
            else:
                ul.decompose()

        # Упорядоченные списки
        for ol in soup.find_all("ol"):
            list_items = []
            for i, li in enumerate(ol.find_all("li"), 1):
                item_text = li.get_text().strip()
                if item_text:
                    list_items.append(f"  {i}. {item_text}")

            if list_items:
                ol.replace_with("\n" + "\n".join(list_items) + "\n")
            else:
                ol.decompose()

    @staticmethod
    def _process_tables(soup):
        """Обрабатывает таблицы в текстовом формате"""
        for table in soup.find_all("table"):
            rows = []
            for tr in table.find_all("tr"):
                row_cells = []
                for cell in tr.find_all(["td", "th"]):
                    cell_text = cell.get_text().strip()
                    row_cells.append(cell_text)

                if row_cells:
                    rows.append(" | ".join(row_cells))

            if rows:
                table_content = "\n".join([f"  {row}" for row in rows])
                table.replace_with(f"\n{table_content}\n")
            else:
                table.decompose()

    @staticmethod
    def _process_links(soup):
        """Обрабатывает ссылки, сохраняя их как текст с URL"""
        for link in soup.find_all("a", href=True):
            link_text = link.get_text().strip()
            link_url = link["href"].strip()

            if link_text and link_url:
                if link_text != link_url:
                    link.replace_with(f"{link_text} [{link_url}]")
                else:
                    link.replace_with(link_url)
            elif link_text:
                link.replace_with(link_text)
            elif link_url:
                link.replace_with(link_url)
            else:
                link.decompose()

    @staticmethod
    def _process_paragraphs(soup):
        """Обрабатывает параграфы и переносы строк"""
        for br in soup.find_all("br"):
            br.replace_with("\n")

        for p in soup.find_all("p"):
            p_text = p.get_text().strip()
            if p_text:
                p.replace_with(f"\n{p_text}\n")
            else:
                p.decompose()

        for div in soup.find_all("div"):
            div_text = div.get_text().strip()
            if div_text:
                # Проверяем, не является ли div контейнером без собственного текста
                if not any(
                    child.name
                    in ["p", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "table"]
                    for child in div.children
                ):
                    div.replace_with(f"\n{div_text}\n")

    @staticmethod
    def _finalize_formatting(text):
        """Финализирует форматирование текста"""
        # Заменяем множественные пустые строки
        text = re.sub(r"\n\s*\n\s*\n", "\n\n", text)

        # Заменяем множественные пробелы
        text = re.sub(r"[ ]{2,}", " ", text)

        # Очищаем пробелы в начале и конце строк
        lines = [line.strip() for line in text.split("\n")]
        text = "\n".join(lines)

        # Удаляем пустые строки в начале и конце
        text = text.strip()

        return text

    @staticmethod
    def extract_attachments_info(html_content):
        """
        Извлекает информацию о вложениях из HTML (если есть упоминания)

        Args:
            html_content (str): HTML контент

        Returns:
            list: Список упомянутых вложений
        """
        if not html_content:
            return []

        soup = BeautifulSoup(html_content, "html.parser")
        attachments = []

        # Ищем упоминания о вложениях (обычно в виде ссылок на файлы)
        for link in soup.find_all("a", href=True):
            href = link["href"]
            text = link.get_text().strip().lower()

            # Проверяем, похоже ли на ссылку на вложение
            if any(
                keyword in text
                for keyword in ["attachment", "вложение", "файл", "download", "скачать"]
            ):
                attachments.append({"name": link.get_text().strip(), "url": href})

        return attachments
