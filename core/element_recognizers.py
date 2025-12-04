import re



class ElementRecognizer:
    """Класс для распознавания различных элементов документа"""

    @staticmethod
    def is_likely_image_paragraph(paragraph):
        """Упрощенный метод определения изображений"""
        if paragraph.text.strip():
            return False

        if not paragraph.runs:
            return False

        all_runs_empty = all(not run.text.strip() for run in paragraph.runs)

        if all_runs_empty:
            try:
                xml_str = str(paragraph._element.xml)
                graphic_indicators = ['drawing', 'graphic', 'picture', 'shape', 'imagedata', 'blip']
                if any(indicator in xml_str.lower() for indicator in graphic_indicators):
                    return True
            except:
                pass

        return False

    @staticmethod
    def is_courier_new_paragraph(paragraph):
        """Определение, является ли параграф листингом (по шрифту Courier New)"""
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

    @staticmethod
    def is_courier_new_table(table):
        """Определение, содержит ли таблица листинг (по шрифту)"""
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if ElementRecognizer.is_courier_new_paragraph(paragraph):
                        return True
        return False

    @staticmethod
    def is_list_item(paragraph):
        """Определение, является ли параграф элементом списка"""
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

    @staticmethod
    def detect_list_type(paragraph):
        """Определение типа списка (нумерованный/маркированный)"""
        text = paragraph.text.strip()

        # Проверяем по тексту параграфа в первую очередь
        if re.match(r'^\d+(?:\.\d+)*\.\s+', text):
            return 'numbered'

        if re.match(r'^[•\-—–]\s+', text):
            return 'bulleted'

        # Если текст не содержит явных маркеров, проверяем XML структуру
        if paragraph._element.pPr is not None:
            num_pr = paragraph._element.pPr.numPr
            if num_pr is not None:
                # Проверяем, есть ли в XML указание на тип нумерации
                xml_str = str(paragraph._element.xml)

                # Ищем признаки нумерованного списка в XML
                # Нумерованные списки обычно имеют конкретные numId
                if 'numId' in xml_str:
                    # Пытаемся определить по numId
                    num_id_match = re.search(r'w:numId\s+w:val="(\d+)"', xml_str)
                    if num_id_match:
                        num_id = num_id_match.group(1)
                        # Обычно numId=1 или небольшие числа для нумерованных списков
                        # numId=2 часто для маркированных, но это может варьироваться
                        # Более надежный способ - искать конкретные элементы форматирования
                        if 'w:lvlText w:val="%' in xml_str:
                            return 'numbered'
                        elif 'w:lvlText w:val="•"' in xml_str or 'w:lvlText w:val="-"' in xml_str:
                            return 'bulleted'

                # По умолчанию считаем маркированным, если есть numPr но не удалось определить точно
                return 'bulleted'

        # По умолчанию, если не удалось определить - считаем маркированным
        # но это можно изменить на 'numbered' в зависимости от ваших потребностей
        return 'numbered'  # Изменено с 'bulleted' на 'numbered'

    @staticmethod
    def get_list_hierarchy(paragraph):
        """Определение уровня вложенности списка"""
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