import pandas as pd
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class CodeProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализатор кодов ФЕР")
        self.root.geometry("800x600")

        self.max_level = tk.IntVar(value=3)  # Значение по умолчанию
        self.file_path = None

        self.create_widgets()

    def create_widgets(self):
        # Фрейм для настроек
        settings_frame = ttk.LabelFrame(self.root, text="Настройки", padding=10)
        settings_frame.pack(fill=tk.X, padx=5, pady=5)

        # Поле для выбора максимального уровня
        ttk.Label(settings_frame, text="Максимальный уровень:").grid(row=0, column=0, sticky=tk.W)
        self.level_spinbox = ttk.Spinbox(
            settings_frame,
            from_=1,
            to=10,
            textvariable=self.max_level,
            width=5
        )
        self.level_spinbox.grid(row=0, column=1, sticky=tk.W, padx=5)

        # Кнопка выбора файла
        self.file_btn = ttk.Button(
            settings_frame,
            text="Выбрать файл Excel",
            command=self.select_file
        )
        self.file_btn.grid(row=1, column=0, columnspan=2, pady=5)

        # Кнопка обработки
        self.process_btn = ttk.Button(
            settings_frame,
            text="Обработать коды",
            command=self.process_codes,
            state=tk.DISABLED
        )
        self.process_btn.grid(row=2, column=0, columnspan=2, pady=5)

        # Фрейм для результатов
        result_frame = ttk.LabelFrame(self.root, text="Результаты", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Текстовое поле с прокруткой
        self.result_text = tk.Text(result_frame, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(result_frame, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.pack(fill=tk.BOTH, expand=True)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(
            title="Выберите Excel файл с кодами",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if self.file_path:
            self.process_btn.config(state=tk.NORMAL)
        else:
            self.process_btn.config(state=tk.DISABLED)

    def process_codes(self):
        if not self.file_path:
            messagebox.showwarning("Предупреждение", "Сначала выберите файл!")
            return

        try:
            # Загрузка данных из Excel
            df = pd.read_excel(self.file_path, header=None)

            # Собираем все коды из всех ячеек
            all_codes = []
            for cell in df[0]:
                # Обрабатываем каждую ячейку - разбиваем по переносам строк
                if pd.notna(cell):
                    codes_in_cell = str(cell).split('\n')
                    for code in codes_in_cell:
                        code = code.strip()
                        if code:  # Игнорируем пустые строки
                            all_codes.append(code)

            # Обработка кодов
            tree = self.build_tree(all_codes)

            # Очистка предыдущих результатов
            self.result_text.delete(1.0, tk.END)

            # Вывод результатов
            self.print_tree_to_widget(tree)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")

    def build_tree(self, codes):
        tree = defaultdict(dict)
        max_level = self.max_level.get()

        for code in codes:
            parts = code.split('-')
            current_level = tree

            for i in range(min(len(parts), max_level)):
                part = parts[i].strip()  # Удаляем лишние пробелы
                if part not in current_level:
                    current_level[part] = defaultdict(dict)
                current_level = current_level[part]

        return tree

    def print_tree_to_widget(self, node, level=0, prefix=""):
        for part, child in sorted(node.items()):
            # Определяем отступы
            indent = '    ' * level  # Используем 4 пробела для каждого уровня

            # Формируем отображаемую часть
            if level == 0:
                display_part = part
            else:
                display_part = part.split('-')[-1] if '-' in part else part

            # Добавляем в текстовое поле
            self.result_text.insert(tk.END, f"{indent}{display_part} (ур {level + 1})\n")

            # Рекурсивно обрабатываем дочерние элементы
            self.print_tree_to_widget(child, level + 1, f"{prefix}-{part}" if prefix else part)

        # Автоматическая прокрутка в начало
        self.result_text.see(tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    app = CodeProcessorApp(root)
    root.mainloop()