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
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl

except ImportError as e:
    print(f"Ошибка импорта: {e}")
    input("Нажмите Enter для выхода...")
    sys.exit(1)


class GOSTConverter:
    def __init__(self):
        self.headings = []
        self.image_counter = 1
        self.table_counter = 1
        self.listing_counter = 1
        self.current_list_info = None
        self.current_section_for_listing = 1
        self.listing_detected = False
        self.last_listing_number = None
        self.section_numbers = {}
        self.list_numbers = {}  # Хранит текущие номера для каждого уровня и типа списка
        self.previous_list_paragraph_index = -1  # Индекс предыдущего параграфа списка
        self.list_groups = []  # Группы связанных списковых элементов
        self.current_list_group_id = 0  # ID текущей группы списков

    def detect_headings(self, doc):
        self.headings = []
        SPECIAL_TITLES = {"СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ",
                          "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
                          "СПИСОК ЛИТЕРАТУРЫ", "ПРИЛОЖЕНИЯ"}

        pattern = re.compile(
            r'^(\d+(?:\.\d+)*)'
            r'\s+'
            r'([А-ЯA-Z].*)'
            r'$',
            re.UNICODE
        )

        for i, paragraph in enumerate(doc.paragraphs):
            raw_text = paragraph.text.strip()
            if not raw_text:
                continue

            if raw_text.upper() in SPECIAL_TITLES:
                self.headings.append({
                    'text': raw_text,
                    'level': 1,
                    'index': i,
                    'original_paragraph': paragraph
                })
                continue

            match = pattern.match(raw_text)
            if not match:
                continue

            number_part = match.group(1)
            title_text = match.group(2).strip()

            if title_text.upper() in SPECIAL_TITLES:
                level = 1
            else:
                if not title_text[0].isupper():
                    continue
                if len(title_text) < 3:
                    continue
                if title_text.lower().startswith(('в ', 'на ', 'по ', 'из ', 'от ', 'к ', 'с ')):
                    continue

                level = number_part.count('.') + 1
                if level > 3:
                    level = 3

            clean_title = f"{number_part} {title_text}" if number_part else title_text

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
        SPECIAL_TITLES = {"СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ",
                          "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
                          "СПИСОК ЛИТЕРАТУРЫ", "ПРИЛОЖЕНИЯ"}

        clean_text = text.strip()

        is_special = clean_text.upper() in SPECIAL_TITLES

        title_part = clean_text
        if not is_special and ' ' in clean_text:
            parts = clean_text.split(' ', 1)
            if parts[0].replace('.', '').isdigit():
                title_part = parts[1] if len(parts) > 1 else parts[0]
            else:
                title_part = clean_text

        if title_part.upper() in SPECIAL_TITLES:
            is_special = True

        if level == 1 or is_special:
            display_text = clean_text.upper()
            style = 'Heading 1 GOST'
        elif level == 2:
            if ' ' in clean_text:
                num, title = clean_text.split(' ', 1)
                display_text = f"{num} {title.strip()[0].upper() + title.strip()[1:]}"
            else:
                display_text = clean_text[0].upper() + clean_text[1:]
            style = 'Heading 2 GOST'
        else:
            if ' ' in clean_text:
                num, title = clean_text.split(' ', 1)
                display_text = f"{num} {title.strip()[0].upper() + title.strip()[1:]}"
            else:
                display_text = clean_text[0].upper() + clean_text[1:]
            style = 'Heading 3 GOST'

        p = doc.add_paragraph(display_text, style=style)

        if is_special:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        for run in p.runs:
            run.font.color.rgb = RGBColor(0, 0, 0)

        if level == 1 and not is_special:
            match = re.match(r'^(\d+)', clean_text)
            if match:
                self.current_section_for_listing = int(match.group(1))
                self.listing_counter = 1
                # Сбрасываем группы списков при новом разделе
                self.list_groups = []
                self.current_list_group_id = 0
        elif is_special:
            self.current_section_for_listing = 0
            self.listing_counter = 1
            # Сбрасываем группы списков при специальном разделе
            self.list_groups = []
            self.current_list_group_id = 0

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

    def is_courier_new_paragraph(self, paragraph):
        if not paragraph.runs:
            return False

        for run in paragraph.runs:
            if run.font.name:
                font_name = run.font.name.lower()
                if 'courier' in font_name:
                    return True

        if paragraph.style and 'code' in paragraph.style.name.lower():
            return True

        return False

    def is_listing_caption(self, paragraph):
        text = paragraph.text.strip().lower()
        return 'листинг' in text or 'listing' in text

    def add_listing_caption_gost(self, target_doc, listing_number, caption_text=""):
        if caption_text and caption_text.strip():
            clean_caption = re.sub(r'^листинг\s*\d+(\.\d+)*\s*[—-]\s*', '', caption_text, flags=re.IGNORECASE)
            clean_caption = re.sub(r'^listing\s*\d+(\.\d+)*\s*[—-]\s*', '', clean_caption, flags=re.IGNORECASE)
            full_caption = f"Листинг {listing_number} — {clean_caption.strip()}"
        else:
            full_caption = f"Листинг {listing_number}"

        caption_para = target_doc.add_paragraph()

        caption_run = caption_para.add_run(full_caption)
        caption_run.font.name = 'Times New Roman'
        caption_run.font.size = Pt(12)
        caption_run.font.italic = True
        caption_run.font.color.rgb = RGBColor(0, 0, 0)

        pf = caption_para.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.left_indent = Inches(0)
        pf.right_indent = Inches(0)
        pf.first_line_indent = Inches(0)
        pf.space_before = Pt(6)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0

        return caption_para

    def add_listing_content_gost(self, source_paragraph, target_doc):
        listing_para = target_doc.add_paragraph()

        for run in source_paragraph.runs:
            new_run = listing_para.add_run(run.text)
            new_run.font.name = 'Courier New'
            new_run.font.size = Pt(10)
            new_run.font.color.rgb = RGBColor(0, 0, 0)
            new_run.bold = False
            new_run.italic = False
            new_run.underline = False

        pf = listing_para.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.left_indent = Inches(0)
        pf.right_indent = Inches(0)
        pf.first_line_indent = Inches(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0

        return listing_para

    def process_listing_paragraph(self, source_paragraph, target_doc):
        if self.current_section_for_listing == 0:
            listing_number = f"0.{self.listing_counter}"
        else:
            listing_number = f"{self.current_section_for_listing}.{self.listing_counter}"

        self.last_listing_number = listing_number

        caption_text = ""

        self.add_listing_caption_gost(target_doc, listing_number, caption_text)
        self.add_listing_content_gost(source_paragraph, target_doc)

        self.listing_counter += 1
        self.listing_detected = True

        return True

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

            first_level1_after_toc = True

            # Сначала анализируем структуру списков
            list_info = []
            for i, paragraph in enumerate(source_doc.paragraphs):
                if self.is_list_item(paragraph):
                    list_type = self.detect_list_type(paragraph)
                    list_level = self.get_list_hierarchy(paragraph)

                    # Определяем, является ли это началом нового списка
                    is_start_of_new_list = False
                    if list_type == 'numbered':
                        # Проверяем, начинается ли с 1
                        text = paragraph.text.strip()
                        if text.startswith('1.') and list_level == 0:
                            # Проверяем, был ли до этого обычный текст или другой список
                            if i == 0:
                                is_start_of_new_list = True
                            else:
                                # Проверяем предыдущий элемент
                                prev_paragraph = source_doc.paragraphs[i - 1]
                                if not self.is_list_item(prev_paragraph):
                                    is_start_of_new_list = True

                    list_info.append({
                        'index': i,
                        'type': list_type,
                        'level': list_level,
                        'paragraph': paragraph,
                        'is_start_of_new_list': is_start_of_new_list,
                        'list_id': None  # Будет заполнено позже
                    })

            # Группируем связанные элементы списков
            current_list_id = 0
            for idx, info in enumerate(list_info):
                if idx == 0:
                    # Первый элемент списка всегда начинает новую группу
                    info['list_id'] = current_list_id
                    current_list_id += 1
                else:
                    prev_info = list_info[idx - 1]

                    # Проверяем, является ли это продолжением предыдущего списка
                    if (info['is_start_of_new_list'] or
                            info['type'] != prev_info['type'] or
                            info['level'] < prev_info['level']):
                        # Начинаем новую группу
                        info['list_id'] = current_list_id
                        current_list_id += 1
                    else:
                        # Продолжаем предыдущую группу
                        info['list_id'] = prev_info['list_id']

            # Определяем последний элемент каждой группы
            last_list_items = set()
            for idx, info in enumerate(list_info):
                is_last = True
                for next_idx in range(idx + 1, len(list_info)):
                    next_info = list_info[next_idx]
                    if next_info['list_id'] == info['list_id']:
                        is_last = False
                        break

                if is_last:
                    last_list_items.add(info['index'])

            # Создаем числовые определения для разных групп списков
            num_counter = 3  # Начинаем с 3, так как 1 и 2 уже используются

            # Собираем уникальные группы
            unique_groups = set(info['list_id'] for info in list_info)
            group_num_ids = {}
            for group_id in unique_groups:
                if group_id is not None:
                    group_num_ids[group_id] = num_counter
                    num_counter += 1

            in_list = False
            current_list_level = 0
            self.listing_detected = False

            body_elements = list(source_doc.element.body)
            element_index = 0
            paragraph_index = 0
            table_index = 0

            ordered_elements = []

            for elem in body_elements:
                if isinstance(elem, CT_P):
                    if paragraph_index < len(source_doc.paragraphs):
                        ordered_elements.append({
                            'type': 'paragraph',
                            'element': source_doc.paragraphs[paragraph_index],
                            'index': paragraph_index
                        })
                        paragraph_index += 1
                elif isinstance(elem, CT_Tbl):
                    if table_index < len(source_doc.tables):
                        ordered_elements.append({
                            'type': 'table',
                            'element': source_doc.tables[table_index],
                            'index': table_index
                        })
                        table_index += 1

            for element_info in ordered_elements:
                if element_info['type'] == 'paragraph':
                    paragraph = element_info['element']
                    i = element_info['index']
                    text = paragraph.text.strip()

                    if not text and not self.is_likely_image_paragraph(paragraph) and not self.is_courier_new_paragraph(
                            paragraph):
                        continue

                    if self.is_likely_image_paragraph(paragraph):
                        self.add_image_placeholder_gost(target_doc, self.image_counter)
                        self.image_counter += 1
                        continue

                    if self.is_courier_new_paragraph(paragraph):
                        self.process_listing_paragraph(paragraph, target_doc)
                        continue

                    if i in heading_indices:
                        heading = next(h for h in self.headings if h['index'] == i)
                        level = heading['level']

                        if level == 1 and not first_level1_after_toc:
                            target_doc.add_page_break()

                        if level == 1 and first_level1_after_toc:
                            first_level1_after_toc = False

                        self.add_heading_with_style(heading['text'], target_doc, level)
                        in_list = False
                        self.current_list_info = None
                        self.listing_detected = False
                    else:
                        if self.is_list_item(paragraph):
                            list_level = self.get_list_hierarchy(paragraph)
                            list_type = self.detect_list_type(paragraph)

                            # Находим информацию о группе для этого элемента
                            list_item_info = next((info for info in list_info if info['index'] == i), None)
                            is_last = i in last_list_items

                            if list_item_info:
                                group_id = list_item_info['list_id']
                                self.add_list_item_gost(paragraph, target_doc, list_level, list_type, is_last, group_id,
                                                        group_num_ids.get(group_id))
                            else:
                                self.add_list_item_gost(paragraph, target_doc, list_level, list_type, is_last)

                            in_list = True
                            current_list_level = list_level
                        else:
                            if in_list:
                                target_doc.add_paragraph()
                                in_list = False
                                self.current_list_info = None

                            self.add_normal_text_gost(paragraph, target_doc)
                            self.listing_detected = False

                elif element_info['type'] == 'table':
                    table = element_info['element']

                    if self.is_courier_new_table(table):
                        self.process_listing_table(table, target_doc)
                    else:
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

        if re.match(r'^\d+(?:\.\d+)*\.\s+', text):
            return True

        if re.match(r'^[•\-—–]\s+', text):
            return True

        return False

    def detect_list_type(self, paragraph):
        text = paragraph.text.strip()

        # Принудительная проверка для первого элемента
        # Если начинается с "1." - это точно нумерованный список
        if re.match(r'^1\.\s+', text):
            return 'numbered'

        if re.match(r'^\d+(?:\.\d+)*\.\s+', text):
            return 'numbered'
        elif re.match(r'^[•\-—–]\s+', text):
            return 'bulleted'

        if paragraph._element.pPr is not None:
            num_pr = paragraph._element.pPr.numPr
            if num_pr is not None:
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

        if re.match(r'^(\d+(?:\.\d+)+)\.\s+', text):
            match = re.match(r'^(\d+(?:\.\d+)+)\.', text)
            if match:
                number_part = match.group(1)
                return number_part.count('.') - 1

        return 0

    def add_list_item_gost(self, source_paragraph, target_doc, list_hierarchy=0, list_type='bulleted', is_last=False,
                           group_id=None, num_id=None):
        """Добавляет элемент списка с независимой нумерацией для каждой группы"""
        text = source_paragraph.text.strip()

        clean_text = text

        if list_type == 'numbered':
            # Извлекаем номер из исходного текста
            match = re.match(r'^(\d+(?:\.\d+)*)\.\s*(.*)', text)
            if match:
                original_number = match.group(1)
                clean_text = match.group(2)
            else:
                clean_text = re.sub(r'^\d+(?:\.\d+)*\.\s*', '', text)
        elif list_type == 'bulleted':
            clean_text = re.sub(r'^[•\-—–]\s*', '', text)

        # Форматирование по ГОСТу:
        if clean_text:
            if list_type == 'numbered':
                if clean_text and clean_text[0].isalpha():
                    clean_text = clean_text[0].upper() + clean_text[1:]

                if not clean_text.endswith(('.', '!', '?')):
                    clean_text = clean_text + '.'
            else:
                if clean_text and clean_text[0].isalpha():
                    clean_text = clean_text[0].lower() + clean_text[1:]

                if not clean_text.endswith(('.', '!', '?', ';')):
                    if is_last:
                        clean_text = clean_text + '.'
                    else:
                        clean_text = clean_text + ';'

        # Создаем параграф
        if list_type == 'numbered':
            paragraph = target_doc.add_paragraph()

            # Используем указанный num_id или создаем новый
            if num_id is None:
                # Если группа не указана, создаем временную нумерацию
                pPr = paragraph._element.get_or_add_pPr()
                numPr = pPr.get_or_add_numPr()

                # Создаем новую нумерацию для этой группы
                ilvl = OxmlElement('w:ilvl')
                ilvl.set(qn('w:val'), str(list_hierarchy))
                numPr.append(ilvl)

                numId = OxmlElement('w:numId')
                # Используем высокий номер, чтобы не конфликтовать со стандартными
                numId.set(qn('w:val'), str(1000 + (group_id if group_id is not None else 0)))
                numPr.append(numId)
            else:
                # Используем существующий numId для группы
                pPr = paragraph._element.get_or_add_pPr()
                numPr = pPr.get_or_add_numPr()

                ilvl = OxmlElement('w:ilvl')
                ilvl.set(qn('w:val'), str(list_hierarchy))
                numPr.append(ilvl)

                numId = OxmlElement('w:numId')
                numId.set(qn('w:val'), str(num_id))
                numPr.append(numId)
        else:
            # Маркированные списки всегда используют один стиль
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

        marker_indent = 0.49 + (0.49 * list_hierarchy)
        text_indent = 0.89 + (0.49 * list_hierarchy)

        pf.left_indent = Inches(text_indent)
        pf.first_line_indent = Inches(-(text_indent - marker_indent))

        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

        return paragraph

    def is_courier_new_table(self, table):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if self.is_courier_new_paragraph(paragraph):
                        return True
        return False

    def process_listing_table(self, table, target_doc):
        if self.current_section_for_listing == 0:
            listing_number = f"0.{self.listing_counter}"
        else:
            listing_number = f"{self.current_section_for_listing}.{self.listing_counter}"

        self.add_listing_caption_gost(target_doc, listing_number, "код программы")

        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.text.strip():
                        temp_doc = Document()
                        temp_para = temp_doc.add_paragraph()

                        for run in paragraph.runs:
                            new_run = temp_para.add_run(run.text)
                            new_run.font.name = 'Courier New'
                            new_run.font.size = Pt(10)

                        self.add_listing_content_gost(temp_para, target_doc)

        self.listing_counter += 1
        self.listing_detected = True


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