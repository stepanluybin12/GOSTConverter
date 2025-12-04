class CountersManager:
    """Менеджер сквозной нумерации элементов в пределах разделов"""

    def __init__(self):
        self.counters = {
            'table': {},  # {раздел: счетчик}
            'figure': {},  # {раздел: счетчик}
            'listing': {},  # {раздел: счетчик}
            'formula': {}  # {раздел: счетчик} - на будущее
        }
        self.current_section = 0  # 0 - специальные разделы (ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ и т.д.)

    def set_current_section(self, section_number):
        """Устанавливает текущий раздел и инициализирует счетчики"""
        self.current_section = section_number

        # Инициализируем счетчики для нового раздела если нужно
        if section_number not in self.counters['table']:
            self.counters['table'][section_number] = 1
        if section_number not in self.counters['figure']:
            self.counters['figure'][section_number] = 1
        if section_number not in self.counters['listing']:
            self.counters['listing'][section_number] = 1

    def get_next_number(self, element_type):
        """Получает следующий номер для элемента в текущем разделе"""
        if element_type not in self.counters:
            raise ValueError(f"Неизвестный тип элемента: {element_type}")

        if self.current_section not in self.counters[element_type]:
            self.counters[element_type][self.current_section] = 1

        number = self.counters[element_type][self.current_section]
        self.counters[element_type][self.current_section] += 1

        # Форматируем номер в зависимости от типа раздела
        if self.current_section == 0:
            return f"0.{number}"
        else:
            return f"{self.current_section}.{number}"

    def get_current_number(self, element_type):
        """Получает текущий номер без увеличения счетчика"""
        if element_type not in self.counters:
            raise ValueError(f"Неизвестный тип элемента: {element_type}")

        if self.current_section not in self.counters[element_type]:
            return f"{self.current_section}.1"
        return f"{self.current_section}.{self.counters[element_type][self.current_section]}"