import sys
import os
from core.gost_converter import GOSTConverter

def main():
    """Основная функция программы"""
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

if __name__ == "__main__":
    main()