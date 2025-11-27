import sys
import os
import re

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.style import WD_STYLE_TYPE
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
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text.strip()
            if not text:
                continue

            if re.match(r'^\d+\.', text):
                dots_count = text.count('.', 0, 10)

                if dots_count == 1:
                    level = 1
                elif dots_count == 2:
                    level = 2
                elif dots_count >= 3:
                    level = 3
                else:
                    level = 1

                self.headings.append({
                    'text': text,
                    'level': level,
                    'index': i,
                    'original_paragraph': paragraph
                })

    def create_styles(self, doc):
        try:
            if 'Heading 1 GOST' not in [s.name for s in doc.styles]:
                heading1_style = doc.styles.add_style('Heading 1 GOST', WD_STYLE_TYPE.PARAGRAPH)
                heading1_style.base_style = doc.styles['Heading 1']
                font = heading1_style.font
                font.name = 'Times New Roman'
                font.size = Pt(18)
                font.bold = True
                font.color.rgb = RGBColor(0, 0, 0)
                heading1_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                heading1_style.paragraph_format.space_before = Pt(0)
                heading1_style.paragraph_format.space_after = Pt(12)
                heading1_style.paragraph_format.line_spacing = 1.5

            if 'Heading 2 GOST' not in [s.name for s in doc.styles]:
                heading2_style = doc.styles.add_style('Heading 2 GOST', WD_STYLE_TYPE.PARAGRAPH)
                heading2_style.base_style = doc.styles['Heading 2']
                font = heading2_style.font
                font.name = 'Times New Roman'
                font.size = Pt(16)
                font.bold = True
                font.color.rgb = RGBColor(0, 0, 0)
                heading2_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                heading2_style.paragraph_format.space_before = Pt(24)
                heading2_style.paragraph_format.space_after = Pt(12)
                heading2_style.paragraph_format.line_spacing = 1.5
                heading2_style.paragraph_format.first_line_indent = Inches(0.49)

            if 'Heading 3 GOST' not in [s.name for s in doc.styles]:
                heading3_style = doc.styles.add_style('Heading 3 GOST', WD_STYLE_TYPE.PARAGRAPH)
                heading3_style.base_style = doc.styles['Heading 3']
                font = heading3_style.font
                font.name = 'Times New Roman'
                font.size = Pt(14)
                font.bold = True
                font.color.rgb = RGBColor(0, 0, 0)
                heading3_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                heading3_style.paragraph_format.space_before = Pt(18)
                heading3_style.paragraph_format.space_after = Pt(6)
                heading3_style.paragraph_format.line_spacing = 1.5
                heading3_style.paragraph_format.first_line_indent = Inches(0.49)

        except Exception as e:
            print(f"Не удалось создать стили: {e}")

    def add_heading_with_style(self, text, doc, level):
        if level == 1:
            paragraph = doc.add_paragraph(text.upper(), style='Heading 1 GOST')
        elif level == 2:
            paragraph = doc.add_paragraph(text, style='Heading 2 GOST')
        else:
            paragraph = doc.add_paragraph(text, style='Heading 3 GOST')

        for run in paragraph.runs:
            run.font.color.rgb = RGBColor(0, 0, 0)

        return paragraph

    def add_normal_text_gost(self, source_paragraph, target_doc):
        if not source_paragraph.text.strip():
            return

        paragraph = target_doc.add_paragraph()

        for run in source_paragraph.runs:
            if run.text.strip():
                new_run = paragraph.add_run(run.text)
                new_run.font.name = 'Times New Roman'
                new_run.font.size = Pt(14)
                new_run.font.color.rgb = RGBColor(0, 0, 0)

                if run.font.bold:
                    new_run.font.bold = True
                if run.font.italic:
                    new_run.font.italic = True

        paragraph_format = paragraph.paragraph_format
        paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph_format.line_spacing = 1.5
        paragraph_format.first_line_indent = Inches(0.49)
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)

    def add_image_placeholder_gost(self, doc, image_number):
        img_para = doc.add_paragraph()
        img_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        img_run = img_para.add_run(f"Рисунок {image_number}")
        img_run.font.name = 'Times New Roman'
        img_run.font.size = Pt(12)
        img_run.font.bold = True
        img_run.font.color.rgb = RGBColor(0, 0, 0)

        img_para.paragraph_format.space_before = Pt(0)
        img_para.paragraph_format.space_after = Pt(6)
        img_para.paragraph_format.line_spacing = 1.0

        placeholder_para = doc.add_paragraph()
        placeholder_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        placeholder_run = placeholder_para.add_run("(место для изображения)")
        placeholder_run.font.name = 'Times New Roman'
        placeholder_run.font.size = Pt(14)
        placeholder_run.font.italic = True
        placeholder_run.font.color.rgb = RGBColor(0, 0, 0)

    def add_table_with_caption_gost(self, source_table, target_doc, table_number):
        try:
            caption_para = target_doc.add_paragraph()
            caption_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            caption_run = caption_para.add_run(f"Таблица {table_number}")
            caption_run.font.name = 'Times New Roman'
            caption_run.font.size = Pt(12)
            caption_run.font.italic = True
            caption_run.font.color.rgb = RGBColor(0, 0, 0)

            caption_para.paragraph_format.space_before = Pt(6)
            caption_para.paragraph_format.space_after = Pt(0)
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

                        for paragraph in new_cell.paragraphs:
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            for run in paragraph.runs:
                                run.font.name = 'Times New Roman'
                                run.font.size = Pt(12)
                                run.font.color.rgb = RGBColor(0, 0, 0)

                target_doc.add_paragraph()

        except Exception as e:
            print(f"Ошибка при обработке таблицы {table_number}: {e}")

    def is_likely_image_paragraph(self, paragraph):
        if not paragraph.text.strip() and len(paragraph.runs) > 0:
            return True
        text = paragraph.text.strip().lower()
        if len(text) < 30 and any(word in text for word in ['рисун', 'изображен', 'фото', 'картинк']):
            return True
        return False

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

            heading_indices = {h['index'] for h in self.headings}
            image_count = 0
            table_count = 0

            for i, paragraph in enumerate(source_doc.paragraphs):
                text = paragraph.text.strip()

                if not text and not self.is_likely_image_paragraph(paragraph):
                    continue

                if self.is_likely_image_paragraph(paragraph):
                    image_count += 1
                    self.add_image_placeholder_gost(target_doc, image_count)
                    continue

                if i in heading_indices:
                    heading = next(h for h in self.headings if h['index'] == i)
                    self.add_heading_with_style(heading['text'], target_doc, heading['level'])
                else:
                    self.add_normal_text_gost(paragraph, target_doc)

            for i, table in enumerate(source_doc.tables, 1):
                self.add_table_with_caption_gost(table, target_doc, i)
                table_count += 1

            if self.headings:
                target_doc.add_page_break()
                self.add_heading_with_style("СОДЕРЖАНИЕ", target_doc, 1)
                target_doc.add_paragraph()
                toc_paragraph = target_doc.add_paragraph()
                toc_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = toc_paragraph.add_run()
                run.add_break()

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

    if not input_file: input_file = "input.docx"
    if not output_file: output_file = "output_gost.docx"

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