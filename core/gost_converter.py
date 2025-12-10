import re
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.enum.section import WD_SECTION

from counters_manager import CountersManager
from element_recognizers import ElementRecognizer
from style_manager import StyleManager
from element_builders import ElementBuilder

class GOSTConverter:
    def __init__(self):
        self.headings = []
        self.counters_manager = CountersManager()
        self.element_builder = ElementBuilder(self.counters_manager)
        self.current_section = 0
        self.current_list_info = None
        self.listing_detected = False
        self.last_listing_number = None
        self.section_numbers = {}
        self.list_numbers = {}
        self.previous_list_paragraph_index = -1
        self.list_groups = []
        self.current_list_group_id = 0
        self.page_number_counter = 1  # Счетчик для нумерации страниц

        # Специальные заголовки, требующие начала с новой страницы
        self.SPECIAL_TITLES_NEW_PAGE = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ",
                                        "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
                                        "СПИСОК ЛИТЕРАТУРЫ", "ПРИЛОЖЕНИЯ", "АННОТАЦИЯ"}
        # Все специальные заголовки (включая СОДЕРЖАНИЕ)
        self.SPECIAL_TITLES = {"СОДЕРЖАНИЕ"} | self.SPECIAL_TITLES_NEW_PAGE

    def detect_headings(self, doc):
        """Обнаружение заголовков в документе"""
        self.headings = []

        pattern = re.compile(
            r'^(\d+(?:\.\d+)*)'
            r'\s+'
            r'([А-Яа-яA-Za-z].*)'
            r'$',
            re.UNICODE
        )

        for i, paragraph in enumerate(doc.paragraphs):
            raw_text = paragraph.text.strip()
            if not raw_text:
                continue

            if raw_text.upper() in self.SPECIAL_TITLES:
                self.headings.append({
                    'text': raw_text,
                    'level': 1,
                    'index': i,
                    'original_paragraph': paragraph,
                    'is_special': True,
                    'requires_new_page': raw_text.upper() in self.SPECIAL_TITLES_NEW_PAGE
                })
                continue

            match = pattern.match(raw_text)
            if not match:
                continue

            number_part = match.group(1)
            title_text = match.group(2).strip()

            if title_text.upper() in self.SPECIAL_TITLES:
                level = 1
                is_special = True
                requires_new_page = title_text.upper() in self.SPECIAL_TITLES_NEW_PAGE
            else:
                level = number_part.count('.') + 1
                if level > 3:
                    level = 3
                is_special = False
                requires_new_page = False

            clean_title = f"{number_part} {title_text}" if number_part else title_text

            self.headings.append({
                'text': clean_title,
                'level': level,
                'index': i,
                'original_paragraph': paragraph,
                'is_special': is_special,
                'requires_new_page': requires_new_page
            })

    def add_heading_with_style(self, text, doc, level, is_special=False, requires_new_page=False):
        """Добавление заголовка со стилем ГОСТ"""
        clean_text = text.strip()

        # Проверяем, нужно ли начинать с новой страницы
        should_start_new_page = False

        if requires_new_page:
            should_start_new_page = True
        elif level == 1 and not is_special:
            # Проверяем, не первый ли это заголовок после СОДЕРЖАНИЯ
            # Если это первый заголовок уровня 1, то это начало текста - нужно с новой страницы
            should_start_new_page = True

        if should_start_new_page:
            doc.add_page_break()
            self.page_number_counter += 1

        # Определяем отображаемый текст
        if is_special:
            display_text = clean_text.upper()
            style = 'Heading 1 GOST'
        elif level == 1:
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

        # Добавляем параграф
        p = doc.add_paragraph(display_text, style=style)

        # Выравнивание
        if is_special:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Цвет текста
        for run in p.runs:
            run.font.color.rgb = RGBColor(0, 0, 0)

        # Обновляем счетчики при смене раздела
        if level == 1 and not is_special:
            match = re.match(r'^(\d+)', clean_text)
            if match:
                section_num = int(match.group(1))
                self.current_section = section_num
                self.counters_manager.set_current_section(section_num)
        elif is_special:
            self.current_section = 0
            self.counters_manager.set_current_section(0)

        return p

    def process_listing_paragraph(self, source_paragraph, target_doc):
        """Обработка листинга с нумерацией в пределах раздела"""
        listing_number = self.counters_manager.get_next_number('listing')
        caption_text = ""

        self.element_builder.add_listing_caption_gost(target_doc, listing_number, caption_text)
        self.element_builder.add_listing_content_with_frame_gost(source_paragraph, target_doc)

        self.listing_detected = True

        return True

    def process_listing_table(self, table, target_doc):
        """Обработка таблицы, содержащей листинг"""
        listing_number = self.counters_manager.get_next_number('listing')

        self.element_builder.add_listing_caption_gost(target_doc, listing_number, "код программы")

        listing_table = target_doc.add_table(rows=1, cols=1)
        listing_table.style = 'Table Grid'
        listing_table.autofit = False

        cell = listing_table.cell(0, 0)
        cell.paragraphs[0].clear()

        for row in table.rows:
            for table_cell in row.cells:
                for paragraph in table_cell.paragraphs:
                    if paragraph.text.strip():
                        listing_para = cell.add_paragraph()

                        for run in paragraph.runs:
                            new_run = listing_para.add_run(run.text)
                            new_run.font.name = 'Courier New'
                            new_run.font.size = Pt(10)
                            new_run.font.color.rgb = RGBColor(0, 0, 0)

        target_doc.add_paragraph().paragraph_format.space_before = Pt(6)

        self.listing_detected = True

    def convert_document(self, input_file, output_file):
        """Основной метод конвертации документа"""
        try:
            source_doc = Document(input_file)
            self.detect_headings(source_doc)

            target_doc = Document()
            StyleManager.create_styles(target_doc)

            # РАЗДЕЛ 1: ТИТУЛЬНЫЙ ЛИСТ (без номера страницы)
            section1 = target_doc.sections[0]
            section1.top_margin = Inches(0.79)
            section1.bottom_margin = Inches(0.79)
            section1.left_margin = Inches(1.18)
            section1.right_margin = Inches(0.39)

            # Очищаем нижний колонтитул для первого раздела
            StyleManager.clear_footer(section1)

            # Добавляем пустой титульный лист
            target_doc.add_paragraph()

            # РАЗДЕЛ 2: СОДЕРЖАНИЕ (без номера страницы)
            target_doc.add_section(WD_SECTION.NEW_PAGE)
            section2 = target_doc.sections[1]
            section2.top_margin = Inches(0.79)
            section2.bottom_margin = Inches(0.79)
            section2.left_margin = Inches(1.18)
            section2.right_margin = Inches(0.39)

            # Очищаем нижний колонтитул для второго раздела
            StyleManager.clear_footer(section2)

            # Добавляем заголовок "СОДЕРЖАНИЕ"
            p = target_doc.add_paragraph("СОДЕРЖАНИЕ")
            p.style = 'Heading 1 GOST'
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.color.rgb = RGBColor(0, 0, 0)

            # Добавляем пустой параграф после заголовка
            target_doc.add_paragraph()

            # РАЗДЕЛ 3: ОСНОВНОЙ ТЕКСТ (нумерация начинается с 3)
            target_doc.add_section(WD_SECTION.NEW_PAGE)
            section3 = target_doc.sections[2]
            section3.top_margin = Inches(0.79)
            section3.bottom_margin = Inches(0.79)
            section3.left_margin = Inches(1.18)
            section3.right_margin = Inches(0.39)

            # Устанавливаем начальный номер страницы для третьего раздела = 3
            StyleManager.set_section_starting_page_number(section3, 3)

            # Добавляем нижний колонтитул с номером страницы для третьего раздела
            StyleManager.create_page_number_footer(section3)

            heading_indices = {h['index'] for h in self.headings}

            # Анализ структуры списков
            list_info = []
            for i, paragraph in enumerate(source_doc.paragraphs):
                if ElementRecognizer.is_list_item(paragraph):
                    list_type = ElementRecognizer.detect_list_type(paragraph)
                    list_level = ElementRecognizer.get_list_hierarchy(paragraph)

                    is_start_of_new_list = False
                    if list_type == 'numbered':
                        text = paragraph.text.strip()
                        if text.startswith('1.') and list_level == 0:
                            if i == 0:
                                is_start_of_new_list = True
                            else:
                                prev_paragraph = source_doc.paragraphs[i - 1]
                                if not ElementRecognizer.is_list_item(prev_paragraph):
                                    is_start_of_new_list = True

                    list_info.append({
                        'index': i,
                        'type': list_type,
                        'level': list_level,
                        'paragraph': paragraph,
                        'is_start_of_new_list': is_start_of_new_list,
                        'list_id': None
                    })

            # Группируем связанные элементы списков
            current_list_id = 0
            for idx, info in enumerate(list_info):
                if idx == 0:
                    info['list_id'] = current_list_id
                    current_list_id += 1
                else:
                    prev_info = list_info[idx - 1]

                    if (info['is_start_of_new_list'] or
                            info['type'] != prev_info['type'] or
                            info['level'] < prev_info['level']):
                        info['list_id'] = current_list_id
                        current_list_id += 1
                    else:
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
            num_counter = 3
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

            # Флаг для отслеживания, был ли добавлен первый заголовок в третьем разделе
            first_heading_in_section3 = True

            for element_info in ordered_elements:
                if element_info['type'] == 'paragraph':
                    paragraph = element_info['element']
                    i = element_info['index']
                    text = paragraph.text.strip()

                    if not text and not ElementRecognizer.is_likely_image_paragraph(paragraph) and not ElementRecognizer.is_courier_new_paragraph(
                            paragraph):
                        continue

                    if ElementRecognizer.is_likely_image_paragraph(paragraph):
                        self.element_builder.add_image_placeholder_gost(target_doc)
                        continue

                    if ElementRecognizer.is_courier_new_paragraph(paragraph):
                        self.process_listing_paragraph(paragraph, target_doc)
                        continue

                    if i in heading_indices:
                        heading = next(h for h in self.headings if h['index'] == i)
                        level = heading['level']
                        is_special = heading.get('is_special', False)
                        requires_new_page = heading.get('requires_new_page', False)

                        # Если это первый заголовок в третьем разделе, не добавляем разрыв страницы
                        if first_heading_in_section3:
                            first_heading_in_section3 = False
                            # Для первого заголовка не добавляем разрыв, даже если requires_new_page
                            requires_new_page = False

                        self.add_heading_with_style(
                            heading['text'], target_doc, level,
                            is_special=is_special, requires_new_page=requires_new_page
                        )
                        in_list = False
                        self.current_list_info = None
                        self.listing_detected = False
                    else:
                        if ElementRecognizer.is_list_item(paragraph):
                            list_level = ElementRecognizer.get_list_hierarchy(paragraph)
                            list_type = ElementRecognizer.detect_list_type(paragraph)

                            list_item_info = next((info for info in list_info if info['index'] == i), None)
                            is_last = i in last_list_items

                            if list_item_info:
                                group_id = list_item_info['list_id']
                                self.element_builder.add_list_item_gost(paragraph, target_doc, list_level, list_type, is_last, group_id,
                                                        group_num_ids.get(group_id))
                            else:
                                self.element_builder.add_list_item_gost(paragraph, target_doc, list_level, list_type, is_last)

                            in_list = True
                            current_list_level = list_level
                        else:
                            if in_list:
                                target_doc.add_paragraph()
                                in_list = False
                                self.current_list_info = None

                            self.element_builder.add_normal_text_gost(paragraph, target_doc)
                            self.listing_detected = False

                elif element_info['type'] == 'table':
                    table = element_info['element']

                    if ElementRecognizer.is_courier_new_table(table):
                        self.process_listing_table(table, target_doc)
                    else:
                        self.element_builder.add_table_with_caption_gost(table, target_doc)

            target_doc.save(output_file)
            return True

        except Exception as e:
            print(f"Ошибка: {e}")
            import traceback
            traceback.print_exc()
            return False