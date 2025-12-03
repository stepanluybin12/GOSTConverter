import sys
import os
import re

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError as e:
    print(f"Ошибка импорта: {e}")
    input("Нажмите Enter для выхода...")
    sys.exit(1)


class GOSTConverter:
    def __init__(self):
        self.headings = []
        self.image_counter = 1
        self.table_counter = 1
        self.current_list_info = None

    def detect_headings(self, doc):
        self.headings = []

        pattern = re.compile(
            r'^(\d+(?:\.\d+)*)'  # Номер (может быть многоуровневым: 1, 1.1, 1.1.1)
            r'\s+'  # Один или более пробелов (НО НЕ ТОЧКА!)
            r'([А-ЯA-Z].*)'  # Текст заголовка
            r'$',
            re.UNICODE
        )

        for i, paragraph in enumerate(doc.paragraphs):
            raw_text = paragraph.text.strip()
            if not raw_text:
                continue

            match = pattern.match(raw_text)
            if not match:
                continue

            number_part = match.group(1)
            title_text = match.group(2).strip()

            if not title_text[0].isupper():
                continue
            if len(title_text) < 3:
                continue
            if title_text.lower().startswith(('в ', 'на ', 'по ', 'из ', 'от ', 'к ', 'с ')):
                continue

            level = number_part.count('.') + 1
            if level > 3:
                level = 3

            clean_title = f"{number_part} {title_text}"

            self.headings.append({
                'text': clean_title,
                'level': level,
                'index': i,
                'original_paragraph': paragraph
            })

    def create_styles(self, doc):
        try:
            styles_to_create = [
                ('Heading 1 GOST', WD_STYLE_TYPE.PARAGRAPH, doc.styles['Heading 1'], {
                    'name': 'Times New Roman',
                    'size': Pt(18),
                    'bold': True,
                    'color': RGBColor(0, 0, 0),
                    'alignment': WD_ALIGN_PARAGRAPH.JUSTIFY,
                    'space_before': Pt(0),
                    'space_after': Pt(12),
                    'line_spacing': 1.5,
                    'first_line_indent': Inches(0.5)
                }),
                ('Heading 2 GOST', WD_STYLE_TYPE.PARAGRAPH, doc.styles['Heading 2'], {
                    'name': 'Times New Roman',
                    'size': Pt(16),
                    'bold': True,
                    'color': RGBColor(0, 0, 0),
                    'alignment': WD_ALIGN_PARAGRAPH.JUSTIFY,
                    'space_before': Pt(24),
                    'space_after': Pt(12),
                    'line_spacing': 1.5,
                    'first_line_indent': Inches(0.5),
                }),
                ('Heading 3 GOST', WD_STYLE_TYPE.PARAGRAPH, doc.styles['Heading 3'], {
                    'name': 'Times New Roman',
                    'size': Pt(14),
                    'bold': True,
                    'color': RGBColor(0, 0, 0),
                    'alignment': WD_ALIGN_PARAGRAPH.JUSTIFY,
                    'space_before': Pt(24),
                    'space_after': Pt(12),
                    'line_spacing': 1.5,
                    'first_line_indent': Inches(0.5),
                }),
            ]

            for style_name, style_type, base_style, params in styles_to_create:
                if style_name not in [s.name for s in doc.styles]:
                    style = doc.styles.add_style(style_name, style_type)
                    style.base_style = base_style
                    font = style.font
                    font.name = params['name']
                    font.size = params['size']
                    font.bold = params['bold']
                    font.color.rgb = RGBColor(0, 0, 0)
                    pf = style.paragraph_format
                    pf.alignment = params['alignment']
                    pf.space_before = params['space_before']
                    pf.space_after = params['space_after']
                    pf.line_spacing = params['line_spacing']
                    if 'first_line_indent' in params:
                        pf.first_line_indent = params['first_line_indent']

        except Exception as e:
            print(f"Не удалось создать стили: {e}")

    def add_heading_with_style(self, text, doc, level):
        SPECIAL_TITLES = {"СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
                          "СПИСОК ЛИТЕРАТУРЫ", "ПРИЛОЖЕНИЯ"}

        clean_text = text.strip()

        title_part = clean_text.split(' ', 1)[-1] if ' ' in clean_text else clean_text
        is_special = title_part.upper() in SPECIAL_TITLES

        if level == 1 or is_special:
            display_text = clean_text.upper()
        else:
            if ' ' in clean_text:
                num, title = clean_text.split(' ', 1)
                display_text = f"{num} {title.strip()[0].upper() + title.strip()[1:]}"
            else:
                display_text = clean_text[0].upper() + clean_text[1:]

        if level == 1 or is_special:
            style = 'Heading 1 GOST'
        elif level == 2:
            style = 'Heading 2 GOST'
        else:
            style = 'Heading 3 GOST'

        p = doc.add_paragraph(display_text, style=style)

        if is_special:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        for run in p.runs:
            run.font.color.rgb = RGBColor(0, 0, 0)

        return p

    def add_normal_text_gost(self, source_paragraph, target_doc):
        if not source_paragraph.text.strip():
            return None

        paragraph = target_doc.add_paragraph()

        for run in source_paragraph.runs:
            new_run = paragraph.add_run(run.text)
            new_run.font.name = 'Times New Roman'
            new_run.font.size = Pt(14)
            new_run.font.color.rgb = RGBColor(0, 0, 0)
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline

        pf = paragraph.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.line_spacing = 1.5
        pf.first_line_indent = Inches(0.49)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.left_indent = Inches(0)
        pf.right_indent = Inches(0)

        return paragraph

    def add_image_placeholder_gost(self, doc, image_number, caption="Иллюстрация"):
        placeholder = doc.add_paragraph()
        placeholder.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = placeholder.add_run("(место для изображения)")
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.italic = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        placeholder.paragraph_format.space_after = Pt(6)

        caption_para = doc.add_paragraph()
        caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption_run = caption_para.add_run(f"Рисунок {image_number} — {caption}")
        caption_run.font.name = 'Times New Roman'
        caption_run.font.size = Pt(12)
        caption_run.font.color.rgb = RGBColor(0, 0, 0)

        caption_para.paragraph_format.space_before = Pt(6)
        caption_para.paragraph_format.space_after = Pt(6)
        caption_para.paragraph_format.line_spacing = 1.0

    def add_table_with_caption_gost(self, source_table, target_doc, table_number, caption="Название таблицы"):
        caption_para = target_doc.add_paragraph()
        caption_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        caption_run = caption_para.add_run(f"Таблица {table_number} — {caption}")
        caption_run.font.name = 'Times New Roman'
        caption_run.font.size = Pt(12)
        caption_run.font.italic = True
        caption_run.font.color.rgb = RGBColor(0, 0, 0)

        caption_para.paragraph_format.space_before = Pt(6)
        caption_para.paragraph_format.space_after = Pt(3)
        caption_para.paragraph_format.line_spacing = 1.0

        rows = len(source_table.rows)
        cols = len(source_table.columns)
        if rows > 0 and cols > 0:
            new_table = target_doc.add_table(rows=rows, cols=cols)
            new_table.style = 'Table Grid'

            for row_idx, row in enumerate(source_table.rows):
                for col_idx, cell in enumerate(row.cells):
                    new_cell = new_table.cell(row_idx, col_idx)
                    new_cell.text = cell.text
                    for p in new_cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        for run in p.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(12)
                            run.font.color.rgb = RGBColor(0, 0, 0)

            for cell in new_table.rows[-1].cells:
                for p in cell.paragraphs:
                    if p.text.strip():
                        p.paragraph_format.space_after = Pt(6)

    def is_likely_image_paragraph(self, paragraph):
        if not paragraph.text.strip() and len(paragraph.runs) > 0:
            return True
        text = paragraph.text.strip().lower()
        return len(text) < 30 and any(word in text for word in ['рисун', 'изображен', 'фото', 'схем', 'график'])

    def create_page_number_footer(self, section):
        footer = section.footer

        for element in footer.paragraphs:
            p = element._element
            p.getparent().remove(p)

        paragraph = footer.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = paragraph.add_run()

        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')

        instrText = OxmlElement('w:instrText')
        instrText.text = "PAGE"
        instrText.set(qn('xml:space'), 'preserve')

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)

    def add_page_numbers_to_document(self, doc):
        for section in doc.sections:
            self.create_page_number_footer(section)

        for section in doc.sections:
            section.footer_distance = Inches(0.3)

    def convert_document(self, input_file, output_file):
        try:
            source_doc = Document(input_file)
            self.detect_headings(source_doc)

            target_doc = Document()
            self.create_styles(target_doc)

            section = target_doc.sections[0]
            section.top_margin = Inches(0.79)
            section.bottom_margin = Inches(0.79)
            section.left_margin = Inches(1.18)
            section.right_margin = Inches(0.39)

            self.add_page_numbers_to_document(target_doc)

            heading_indices = {h['index'] for h in self.headings}

            # Добавляем "СОДЕРЖАНИЕ" в самое начало
            if self.headings:
                self.add_heading_with_style("СОДЕРЖАНИЕ", target_doc, 1)
                target_doc.add_paragraph()

            # Флаг для отслеживания первого заголовка 1-го уровня после содержания
            first_level1_after_toc = True

            # Собираем информацию о списках заранее для определения последнего элемента
            list_info = []
            for i, paragraph in enumerate(source_doc.paragraphs):
                if self.is_list_item(paragraph):
                    list_type = self.detect_list_type(paragraph)
                    list_level = self.get_list_hierarchy(paragraph)
                    list_info.append({
                        'index': i,
                        'type': list_type,
                        'level': list_level,
                        'paragraph': paragraph
                    })

            # Определяем последний элемент каждого списка
            last_list_items = set()
            for idx, info in enumerate(list_info):
                is_last = True
                current_type = info['type']
                current_level = info['level']

                for next_idx in range(idx + 1, len(list_info)):
                    next_info = list_info[next_idx]
                    if next_info['type'] == current_type and next_info['level'] == current_level:
                        is_last = False
                        break
                    if next_info['level'] <= current_level:
                        break

                if is_last:
                    last_list_items.add(info['index'])

            in_list = False
            current_list_level = 0

            for i, paragraph in enumerate(source_doc.paragraphs):
                text = paragraph.text.strip()

                if not text and not self.is_likely_image_paragraph(paragraph):
                    continue

                if self.is_likely_image_paragraph(paragraph):
                    self.add_image_placeholder_gost(target_doc, self.image_counter)
                    self.image_counter += 1
                    continue

                if i in heading_indices:
                    heading = next(h for h in self.headings if h['index'] == i)
                    level = heading['level']

                    # Добавляем разрыв страницы перед заголовком 1-го уровня
                    # (кроме самого первого после содержания)
                    if level == 1 and not first_level1_after_toc:
                        target_doc.add_page_break()

                    # После добавления первого заголовка 1-го уровня сбрасываем флаг
                    if level == 1 and first_level1_after_toc:
                        first_level1_after_toc = False

                    self.add_heading_with_style(heading['text'], target_doc, level)
                    in_list = False
                    self.current_list_info = None
                else:
                    if self.is_list_item(paragraph):
                        list_level = self.get_list_hierarchy(paragraph)
                        list_type = self.detect_list_type(paragraph)
                        is_last = i in last_list_items
                        self.add_list_item_gost(paragraph, target_doc, list_level, list_type, is_last)
                        in_list = True
                        current_list_level = list_level
                    else:
                        if in_list:
                            target_doc.add_paragraph()
                            in_list = False
                            self.current_list_info = None

                        self.add_normal_text_gost(paragraph, target_doc)

            for i, table in enumerate(source_doc.tables):
                self.add_table_with_caption_gost(table, target_doc, self.table_counter)
                self.table_counter += 1

            target_doc.save(output_file)
            return True

        except Exception as e:
            print(f"Ошибка: {e}")
            import traceback
            traceback.print_exc()
            return False

    def is_list_item(self, paragraph):
        if paragraph._element.pPr is not None:
            num_pr = paragraph._element.pPr.numPr
            if num_pr is not None:
                return True

        text = paragraph.text.strip()

        # ИСПРАВЛЕНИЕ 1: Проверяем нумерованные списки (с точкой после номера)
        # Теперь ищем только формат "1. текст" или "1.1. текст" с точкой
        if re.match(r'^\d+(?:\.\d+)*\.\s+', text):
            return True

        # Проверяем маркированные списки
        if re.match(r'^[•\-—–]\s+', text):
            return True

        return False

    def detect_list_type(self, paragraph):
        """Определяет тип списка: 'numbered' или 'bulleted'"""
        text = paragraph.text.strip()

        # ИСПРАВЛЕНИЕ 2: Проверяем нумерованные списки только с точкой
        if re.match(r'^\d+(?:\.\d+)*\.\s+', text):
            return 'numbered'
        elif re.match(r'^[•\-—–]\s+', text):
            return 'bulleted'

        # По стилю Word
        if paragraph._element.pPr is not None:
            num_pr = paragraph._element.pPr.numPr
            if num_pr is not None:
                # По умолчанию считаем маркированным
                return 'bulleted'

        return 'bulleted'

    def get_list_hierarchy(self, paragraph):
        if paragraph._element.pPr is not None:
            num_pr = paragraph._element.pPr.numPr
            if num_pr is not None and num_pr.ilvl is not None:
                return num_pr.ilvl.val
            elif num_pr is not None:
                return 0

        text = paragraph.text.strip()

        # ИСПРАВЛЕНИЕ 3: Для нумерованных списков с вложенностью вида 1.1., 1.1.1. и т.д.
        if re.match(r'^(\d+(?:\.\d+)+)\.\s+', text):
            # Считаем уровень по количеству точек в номере
            match = re.match(r'^(\d+(?:\.\d+)+)\.', text)
            if match:
                number_part = match.group(1)
                return number_part.count('.') - 1  # 1.1 -> уровень 0, 1.1.1 -> уровень 1 и т.д.

        return 0

    def add_list_item_gost(self, source_paragraph, target_doc, list_hierarchy=0, list_type='bulleted', is_last=False):
        """Добавляет элемент списка с использованием настоящих списков Word и правильными отступами по ГОСТу"""
        text = source_paragraph.text.strip()

        # Извлекаем чистый текст без маркера/номера
        clean_text = text

        if list_type == 'numbered':
            # Удаляем номер в начале (например, "1. ", "1.1. ")
            clean_text = re.sub(r'^\d+(?:\.\d+)*\.\s*', '', text)
        elif list_type == 'bulleted':
            # Удаляем маркер в начале
            clean_text = re.sub(r'^[•\-—–]\s*', '', text)

        # Форматирование по ГОСТу:
        if clean_text:
            if list_type == 'numbered':
                # Нумерованный список: начинаем с прописной буквы
                if clean_text and clean_text[0].isalpha():
                    clean_text = clean_text[0].upper() + clean_text[1:]

                # ИСПРАВЛЕНИЕ 4: Заканчиваем точкой (добавляем только если нет знака препинания)
                # В нумерованных списках по ГОСТу всегда точка, не точка с запятой
                if not clean_text.endswith(('.', '!', '?')):
                    clean_text = clean_text + '.'
            else:
                # Маркированный список: начинаем с маленькой буквы
                if clean_text and clean_text[0].isalpha():
                    clean_text = clean_text[0].lower() + clean_text[1:]

                # Заканчиваем точкой с запятой или точкой для последнего элемента
                # Не добавляем точку с запятой, если уже есть знак препинания
                if not clean_text.endswith(('.', '!', '?', ';')):
                    if is_last:
                        clean_text = clean_text + '.'
                    else:
                        clean_text = clean_text + ';'

        # Создаем параграф с соответствующим стилем списка
        if list_type == 'numbered':
            # Нумерованный список
            paragraph = target_doc.add_paragraph(style='List Number')
            # Настраиваем уровень вложенности
            if list_hierarchy > 0:
                pPr = paragraph._element.get_or_add_pPr()
                numPr = pPr.get_or_add_numPr()
                ilvl = OxmlElement('w:ilvl')
                ilvl.set(qn('w:val'), str(list_hierarchy))
                numPr.append(ilvl)

                # Устанавливаем numId для вложенных списков
                numId = OxmlElement('w:numId')
                numId.set(qn('w:val'), '1')
                numPr.append(numId)
        else:
            # Маркированный список
            paragraph = target_doc.add_paragraph(style='List Bullet')
            if list_hierarchy > 0:
                pPr = paragraph._element.get_or_add_pPr()
                numPr = pPr.get_or_add_numPr()
                ilvl = OxmlElement('w:ilvl')
                ilvl.set(qn('w:val'), str(list_hierarchy))
                numPr.append(ilvl)

                numId = OxmlElement('w:numId')
                numId.set(qn('w:val'), '2')
                numPr.append(numId)

        # Добавляем текст
        run = paragraph.add_run(clean_text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(0, 0, 0)

        # Настраиваем форматирование параграфа
        pf = paragraph.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.line_spacing = 1.5

        # Устанавливаем отступы по ГОСТу согласно таблице 4.1
        # Положение маркера (номера): отступ 1,25 см = 0.49 дюйма
        # Положение текста: отступ 2,25 см = 0.89 дюйма

        # Вычисляем отступы в зависимости от уровня вложенности
        # Для каждого следующего уровня добавляем 1,25 см (0.49 дюйма)
        marker_indent = 0.49 + (0.49 * list_hierarchy)  # Положение маркера
        text_indent = 0.89 + (0.49 * list_hierarchy)  # Положение текста

        pf.left_indent = Inches(text_indent)  # Отступ текста
        pf.first_line_indent = Inches(-(text_indent - marker_indent))  # Висячий отступ для маркера

        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

        return paragraph


def main():
    if len(sys.argv) == 3:
        input_file, output_file = sys.argv[1], sys.argv[2]
    else:
        input_file = input("Входной файл: ").strip('"')
        output_file = input("Выходной файл: ").strip('"')

    if not input_file:
        input_file = "input.docx"
    if not output_file:
        output_file = "output_gost.docx"

    if not os.path.exists(input_file):
        print(f"Файл не найден: {input_file}")
        return

    converter = GOSTConverter()
    if converter.convert_document(input_file, output_file):
        print(f"Успешно создан: {output_file}")
    else:
        print("Ошибка конвертации")

    input("Нажмите Enter...")


if __name__ == "__main__":
    main()