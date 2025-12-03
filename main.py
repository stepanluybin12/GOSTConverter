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

    def detect_headings(self, doc):
        self.headings = []

        pattern = re.compile(
            r'^(\d+(?:\.\d+)*)'
            r'\.?'
            r'\s*'
            r'([А-ЯA-Z].*)'
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
        """
        Создает футер с номером страницы
        """
        footer = section.footer

        # Удаляем существующие параграфы в футере
        for element in footer.paragraphs:
            p = element._element
            p.getparent().remove(p)

        # Создаем новый параграф для номера страницы
        paragraph = footer.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Добавляем поле номера страницы
        run = paragraph.add_run()

        # Создаем XML элементы для поля номера страницы
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

        # Форматирование номера страницы
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

        # Устанавливаем отступы
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)

    def add_page_numbers_to_document(self, doc):
        """
        Добавляет нумерацию страниц ко всем разделам документа
        """
        for section in doc.sections:
            self.create_page_number_footer(section)

        # Устанавливаем расстояние от текста до футера
        for section in doc.sections:
            section.footer_distance = Inches(0.3)

    def convert_document(self, input_file, output_file):
        try:
            source_doc = Document(input_file)
            self.detect_headings(source_doc)

            target_doc = Document()
            self.create_styles(target_doc)

            section = target_doc.sections[0]
            section.top_margin = Inches(0.79)  # 20 мм
            section.bottom_margin = Inches(0.79)  # 20 мм
            section.left_margin = Inches(1.18)  # 30 мм
            section.right_margin = Inches(0.39)  # 10 мм

            # Добавляем нумерацию страниц
            self.add_page_numbers_to_document(target_doc)

            heading_indices = {h['index'] for h in self.headings}

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
                    self.add_heading_with_style(heading['text'], target_doc, heading['level'])
                else:
                    self.add_normal_text_gost(paragraph, target_doc)

            for i, table in enumerate(source_doc.tables):
                self.add_table_with_caption_gost(table, target_doc, self.table_counter)
                self.table_counter += 1

            if self.headings:
                target_doc.add_page_break()
                self.add_heading_with_style("СОДЕРЖАНИЕ", target_doc, 1)

            target_doc.save(output_file)
            return True

        except Exception as e:
            print(f"Ошибка: {e}")
            import traceback
            traceback.print_exc()
            return False


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