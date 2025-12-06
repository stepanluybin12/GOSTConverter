import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import sys
import os

# Добавляем путь для импорта core
current_dir = os.path.dirname(os.path.abspath(__file__))
core_dir = os.path.join(current_dir, 'core')
if core_dir not in sys.path:
    sys.path.append(core_dir)

# Импортируем конвертер
import gost_converter


class GOSTConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер ГОСТ")
        self.root.geometry("500x400")

        # Центрируем окно
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'500x400+{x}+{y}')

        # Переменные для файлов
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()

        # Создаем интерфейс
        self.create_ui()

    def create_ui(self):
        """Создаем простой интерфейс"""
        # Заголовок
        tk.Label(
            self.root,
            text="Конвертер в формат ГОСТ",
            font=('Arial', 14, 'bold')
        ).pack(pady=10)

        # Входной файл
        tk.Label(self.root, text="Выберите документ:").pack(anchor='w', padx=20)

        frame1 = tk.Frame(self.root)
        frame1.pack(pady=5, padx=20, fill='x')

        tk.Entry(frame1, textvariable=self.input_file, width=40).pack(side='left', padx=(0, 10))
        tk.Button(frame1, text="📁", command=self.select_input).pack()

        # Выходной файл
        tk.Label(self.root, text="Сохранить как:").pack(anchor='w', padx=20, pady=(10, 0))

        frame2 = tk.Frame(self.root)
        frame2.pack(pady=5, padx=20, fill='x')

        tk.Entry(frame2, textvariable=self.output_file, width=40).pack(side='left', padx=(0, 10))
        tk.Button(frame2, text="📁", command=self.select_output).pack()

        # Кнопка конвертации
        self.convert_btn = tk.Button(
            self.root,
            text="Конвертировать",
            font=('Arial', 12, 'bold'),
            bg='green',
            fg='white',
            command=self.start_conversion,
            height=2,
            width=20
        )
        self.convert_btn.pack(pady=20)

        # Статус
        self.status_label = tk.Label(self.root, text="", font=('Arial', 10))
        self.status_label.pack()

        # Прогресс
        self.progress = ttk.Progressbar(self.root, mode='indeterminate', length=400)
        self.progress.pack(pady=10)

        # Лог
        self.log_text = tk.Text(self.root, height=8, width=55, font=('Courier New', 8))
        self.log_text.pack(pady=10, padx=20)

    def select_input(self):
        """Выбор входного файла"""
        filename = filedialog.askopenfilename(
            title="Выберите документ Word",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            # Автоматически предлагаем имя для выходного файла
            if not self.output_file.get():
                base = os.path.splitext(filename)[0]
                self.output_file.set(f"{base}_ГОСТ.docx")

    def select_output(self):
        """Выбор выходного файла"""
        filename = filedialog.asksaveasfilename(
            title="Сохранить результат",
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)

    def log(self, message):
        """Добавляет сообщение в лог"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)

    def start_conversion(self):
        """Начинает конвертацию"""
        input_path = self.input_file.get()
        output_path = self.output_file.get()

        if not input_path:
            messagebox.showerror("Ошибка", "Выберите входной файл!")
            return

        if not output_path:
            messagebox.showerror("Ошибка", "Выберите выходной файл!")
            return

        # Отключаем кнопку, показываем прогресс
        self.convert_btn.config(state='disabled', text="⏳ Конвертация...")
        self.progress.start()
        self.log_text.delete(1.0, tk.END)
        self.log("Начинаю конвертацию...")

        # Запускаем в отдельном потоке
        thread = threading.Thread(target=self.convert, args=(input_path, output_path))
        thread.daemon = True
        thread.start()

    def convert(self, input_path, output_path):
        """Выполняет конвертацию"""
        try:
            converter = gost_converter.GOSTConverter()
            success = converter.convert_document(input_path, output_path)

            if success:
                self.log("✅ Конвертация завершена успешно!")
                messagebox.showinfo("Готово", f"Документ сохранен:\n{output_path}")
            else:
                self.log("❌ Ошибка при конвертации")
                messagebox.showerror("Ошибка", "Не удалось сконвертировать документ")

        except Exception as e:
            self.log(f"❌ Ошибка: {str(e)}")
            import traceback
            self.log(traceback.format_exc())
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

        finally:
            # Восстанавливаем интерфейс
            self.root.after(0, self.conversion_finished)

    def conversion_finished(self):
        """Завершение конвертации"""
        self.progress.stop()
        self.convert_btn.config(state='normal', text="Конвертировать")